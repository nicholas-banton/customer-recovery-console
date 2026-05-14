import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import './styles.css';

const EMAIL_RE = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;

function uniqueEmailsFromText(value) {
  return Array.from(new Set(String(value || '').match(EMAIL_RE) || []))
    .map((email) => email.toLowerCase());
}

function cleanDisplayName(value) {
  return String(value || '')
    .replace(/<[^>]+>/g, '')
    .replace(/["']/g, '')
    .trim();
}

function parseEmailHeaderBlock(messageText) {
  const normalized = String(messageText || '').replace(/\r\n/g, '\n');
  const headerText = normalized.split('\n\n')[0] || '';
  const lines = headerText.split('\n');

  const headers = {};
  let currentKey = '';

  for (const line of lines) {
    if (/^\s/.test(line) && currentKey) {
      headers[currentKey] += ' ' + line.trim();
      continue;
    }

    const match = line.match(/^([^:]+):\s*(.*)$/);
    if (match) {
      currentKey = match[1].toLowerCase();
      headers[currentKey] = match[2].trim();
    }
  }

  return headers;
}


function extractUrlsFromText(value) {
  return Array.from(new Set(String(value || '').match(/https?:\/\/[^\s"'<>]+/gi) || []))
    .map((url) => url.replace(/[)\].,]+$/g, ''));
}

function extractEtsyLinksFromText(value) {
  return extractUrlsFromText(value).filter((url) => /etsy\.com/i.test(url));
}

function getPrimaryEtsyLink(value) {
  const links = extractEtsyLinksFromText(value);
  return links.find((url) => /(order|orders|sold|transaction|receipt|buyer|conversation)/i.test(url)) || links[0] || '';
}

function isSystemOrSellerEmail(email) {
  const e = String(email || '').toLowerCase().trim();

  if (!e) return true;
  if (e.endsWith('@etsy.com')) return true;
  if (e.includes('no-reply')) return true;
  if (e.includes('noreply')) return true;
  if (e.includes('donotreply')) return true;
  if (e === 'seller@example.com') return true;

  return false;
}

function getMessageBody(messageText) {
  const normalized = String(messageText || '').replace(/\r\n/g, '\n');
  const parts = normalized.split(/\n\n/);
  return parts.slice(1).join('\n\n');
}

function extractBuyerEmailsFromMessage(messageText) {
  const fullText = String(messageText || '');
  const body = getMessageBody(fullText);
  const headers = parseEmailHeaderBlock(fullText);

  const buyerLines = body
    .split(/\r?\n/)
    .filter((line) => /(buyer|customer|contact|email|reply|message)/i.test(line));

  const labeledEmails = uniqueEmailsFromText(buyerLines.join('\n'))
    .filter((email) => !isSystemOrSellerEmail(email));

  if (labeledEmails.length) {
    return Array.from(new Set(labeledEmails));
  }

  const headerEmails = uniqueEmailsFromText([
    headers.from || '',
    headers.to || '',
    headers.cc || '',
    headers.bcc || ''
  ].join(' '));

  const allBodyEmails = uniqueEmailsFromText(body)
    .filter((email) => !isSystemOrSellerEmail(email))
    .filter((email) => !headerEmails.includes(email));

  return Array.from(new Set(allBodyEmails));
}

function inferBuyerNameFromMessage(messageText, email) {
  const body = getMessageBody(messageText);
  const lines = body.split(/\r?\n/);

  for (const line of lines) {
    const match = line.match(/(?:buyer|customer|name|contact buyer)\s*:?\s*([^<\n]+)/i);
    if (match && match[1]) {
      const cleaned = cleanDisplayName(match[1])
        .replace(email || '', '')
        .replace(/contact buyer/i, '')
        .trim();

      if (cleaned && !cleaned.includes('@') && cleaned.length <= 80) {
        return cleaned;
      }
    }
  }

  if (email) {
    return email.split('@')[0].replace(/[._-]+/g, ' ');
  }

  return 'Email missing — review Etsy link';
}

function normalizeMoneyValue(value) {
  const raw = cleanText(value).replace(/,/g, '');
  const match =
    raw.match(/(?:USD\s*)?\$\s*([0-9]+(?:\.[0-9]{2})?)/i) ||
    raw.match(/(?:USD\s*)?([0-9]+(?:\.[0-9]{2}))/i);

  if (!match?.[1]) return '';

  const amount = Number(match[1]);
  if (!Number.isFinite(amount) || amount <= 0 || amount > 100000) return '';

  return amount.toFixed(2);
}

function extractOrderTotal(messageText) {
  const text = String(messageText || '')
    .replace(/=0D=0A|=0A|=0D/gi, '\n')
    .replace(/=20/gi, ' ')
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  const patterns = [
    /(?:order\s*total|grand\s*total|total\s*paid|amount\s*paid|total\s*charged|payment\s*total|you\s*paid)\s*[:\-]?\s*((?:USD\s*)?\$?\s*[0-9][0-9,]*(?:\.[0-9]{2})?)/i,
    /((?:USD\s*)?\$\s*[0-9][0-9,]*(?:\.[0-9]{2})?)\s*(?:order\s*total|total\s*paid|amount\s*paid|paid)/i
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    const amount = normalizeMoneyValue(match?.[1] || '');
    if (amount) return amount;
  }

  return '';
}

function inferEtsyOrderId(messageText, fallbackIndex) {
  const value = String(messageText || '');

  const urlOrderMatch = value.match(/[?&](?:order_id|receipt_id|transaction_id)=(\d+)/i);
  if (urlOrderMatch) return urlOrderMatch[1];

  const orderMatch = value.match(/(?:order|order_id|receipt|transaction)[^\d]{0,12}(\d{4,})/i);
  if (orderMatch) return orderMatch[1];

  return '';
}

function buildEtsyRecoveryRow({ email, messageIndex, headers, message, fileName }) {
  const etsyLink = getPrimaryEtsyLink(message);
  const buyerName = inferBuyerNameFromMessage(message, email);
  const hasBuyerEmail = Boolean(email);

  return {
    'Customer Name': buyerName,
    'Customer Email': email || '',
    'Order ID': inferEtsyOrderId(message, messageIndex),
    'Order Date': headers.date || '',
    'Order Total': extractOrderTotal(message),
    'Marketing Consent': 'Unknown',
    'Source Type': 'Etsy Gmail Takeout',
    'Subject': headers.subject || '',
    'Raw Source File': fileName,
    'Etsy Link': etsyLink,
    'Etsy Recovery Status': hasBuyerEmail ? 'Email Found' : 'Link Review Needed',
    'Review Status': hasBuyerEmail ? 'Email Found' : 'Link Review Needed'
  };
}

function mboxToRows(text, fileName = 'Gmail Takeout.mbox') {
  const raw = String(text || '');

  const messages = raw
    .split(/\n(?=From [^\n]+\n)/g)
    .filter((part) => part.trim().length > 0);

  const rows = [];
  let messageCount = 0;

  for (const message of messages) {
    messageCount += 1;
    const headers = parseEmailHeaderBlock(message);
    const subject = headers.subject || '';
    const isEtsySale = /etsy/i.test(message) && /(transaction|sale|sold|order|receipt|buyer)/i.test(`${message} ${subject}`);

    if (!isEtsySale) {
      continue;
    }

    const buyerEmails = extractBuyerEmailsFromMessage(message);

    if (buyerEmails.length) {
      for (const email of buyerEmails) {
        rows.push(buildEtsyRecoveryRow({
          email,
          messageIndex: messageCount,
          headers,
          message,
          fileName
        }));
      }
    } else {
      rows.push(buildEtsyRecoveryRow({
        email: '',
        messageIndex: messageCount,
        headers,
        message,
        fileName
      }));
    }
  }

  return rows;
}

function hasDesktopMboxPicker() {
  return (
    typeof window !== 'undefined' &&
    window.crcApp &&
    typeof window.crcApp.selectMboxFile === 'function'
  );
}

function emlToRows(text, fileName = 'Imported Email.eml') {
  const message = String(text || '');
  const headers = parseEmailHeaderBlock(message);
  const subject = headers.subject || '';
  const isEtsySale = /etsy/i.test(message) && /(transaction|sale|sold|order|receipt|buyer)/i.test(`${message} ${subject}`);

  if (!isEtsySale) {
    return [];
  }

  const buyerEmails = extractBuyerEmailsFromMessage(message);

  if (buyerEmails.length) {
    return buyerEmails.map((email, index) => buildEtsyRecoveryRow({
      email,
      messageIndex: index + 1,
      headers,
      message,
      fileName
    }));
  }

  return [buildEtsyRecoveryRow({
    email: '',
    messageIndex: 1,
    headers,
    message,
    fileName
  })];
}

const FIELD_ALIASES = {
  email: ['email', 'e-mail', 'buyer email', 'customer email', 'ship email', 'contact email'],
  name: ['name', 'buyer', 'buyer name', 'customer', 'customer name', 'ship name', 'recipient'],
  orderId: ['order id', 'order', 'order number', 'receipt id', 'transaction id', 'sale id'],
  amount: ['amount', 'total', 'order total', 'price', 'gross', 'net', 'revenue'],
  date: ['date', 'order date', 'sale date', 'created', 'created at', 'purchase date'],
  consent: ['consent', 'marketing consent', 'opt in', 'opt-in', 'subscribed', 'newsletter'],
  sourceType: ['source type', 'source', 'import type'],
  subject: ['subject', 'email subject'],
  etsyLinks: ['etsy links', 'etsy link', 'transaction links', 'buyer links', 'review links'],
  linkReviewStatus: ['link review status', 'review status'],
  rawSourceFile: ['raw source file', 'source file']
};

const SAMPLE_ROWS = [
  {
    'Buyer Name': 'Jane Smith',
    'Buyer Email': 'jane@example.com',
    'Order ID': 'ETSY-1001',
    'Order Total': '48.00',
    'Order Date': '2024-03-14',
    'Marketing Consent': 'Unknown'
  },
  {
    'Buyer Name': 'Jane Smith',
    'Buyer Email': 'JANE@example.com',
    'Order ID': 'ETSY-1017',
    'Order Total': '32.00',
    'Order Date': '2024-05-03',
    'Marketing Consent': 'Unknown'
  },
  {
    'Buyer Name': 'Marcus Lee',
    'Buyer Email': 'marcus@example.net',
    'Order ID': 'ETSY-1002',
    'Order Total': '116.50',
    'Order Date': '2024-07-22',
    'Marketing Consent': 'Yes'
  },
  {
    'Buyer Name': 'Olivia Chen',
    'Buyer Email': 'olivia.invalid',
    'Order ID': 'ETSY-1003',
    'Order Total': '71.25',
    'Order Date': '2023-11-11',
    'Marketing Consent': 'No'
  }
];

function normalizeKey(key) {
  return String(key || '').trim().toLowerCase();
}

function findField(row, type) {
  const keys = Object.keys(row || {});
  const aliases = FIELD_ALIASES[type];
  const direct = keys.find((key) => aliases.includes(normalizeKey(key)));
  if (direct) return row[direct];
  const fuzzy = keys.find((key) => aliases.some((alias) => normalizeKey(key).includes(alias)));
  return fuzzy ? row[fuzzy] : '';
}

function extractEmail(row) {
  const direct = findField(row, 'email');
  const directMatch = String(direct || '').match(EMAIL_RE);
  if (directMatch?.[0]) return directMatch[0].toLowerCase();

  for (const value of Object.values(row || {})) {
    const match = String(value || '').match(EMAIL_RE);
    if (match?.[0]) return match[0].toLowerCase();
  }
  return '';
}

function normalizeConsent(value) {
  const raw = String(value || '').trim().toLowerCase();
  if (['yes', 'true', 'y', 'subscribed', 'opted in', 'opt-in', 'marketing eligible'].includes(raw)) return 'Marketing Eligible';
  if (['no', 'false', 'n', 'unsubscribed', 'opted out', 'opt-out', 'do not contact'].includes(raw)) return 'Do Not Contact';
  return 'Unknown';
}

function toNumber(value) {
  const n = Number(String(value || '').replace(/[^0-9.-]/g, ''));
  return Number.isFinite(n) ? n : 0;
}

function buildCustomers(rows, fileName = 'Imported File') {
  const map = new Map();
  let invalid = 0;

  rows.forEach((row, index) => {
    const email = extractEmail(row);
    const name = String(findField(row, 'name') || '').trim();
    const orderId = String(findField(row, 'orderId') || '').trim();
    const amount = toNumber(findField(row, 'amount'));
    const date = String(findField(row, 'date') || '').trim();
    const consent = normalizeConsent(findField(row, 'consent'));
    const sourceType = String(findField(row, 'sourceType') || 'Imported File').trim();
    const subject = String(findField(row, 'subject') || '').trim();
    const etsyLinks = String(findField(row, 'etsyLinks') || '').trim();
    const linkReviewStatus = String(findField(row, 'linkReviewStatus') || '').trim();
    const rawSourceFile = String(findField(row, 'rawSourceFile') || fileName).trim();
    const isEtsySale = /etsy/i.test(`${sourceType} ${subject} ${etsyLinks} ${rawSourceFile}`);
    const rowNotes = String(row.Notes || row.notes || '').trim();

    if (!email) invalid += 1;
    const key = email || `missing-email-${fileName}-${index}`;
    const existing = map.get(key);

    if (existing) {
      existing.orders += orderId ? 1 : 0;
      existing.totalSpend += amount;
      existing.duplicates += 1;
      existing.sources.add(fileName);
      existing.orderIds.push(orderId || `row-${index + 1}`);
      if (etsyLinks) existing.etsyLinks = Array.from(new Set(`${existing.etsyLinks || ''} | ${etsyLinks}`.split('|').map((s) => s.trim()).filter(Boolean))).join(' | ');
      if (subject && !existing.subject) existing.subject = subject;
      if (sourceType && !existing.sourceType) existing.sourceType = sourceType;
      if (linkReviewStatus && existing.linkReviewStatus !== 'Email Found') existing.linkReviewStatus = linkReviewStatus;
      if (isEtsySale) existing.isEtsySale = true;
      if (rowNotes) existing.notes = existing.notes ? `${existing.notes} ${rowNotes}` : rowNotes;
      if (!existing.name && name) existing.name = name;
      if (consent === 'Do Not Contact') existing.consent = 'Do Not Contact';
      if (consent === 'Marketing Eligible' && existing.consent === 'Unknown') existing.consent = 'Marketing Eligible';
      if (date && (!existing.lastPurchase || date > existing.lastPurchase)) existing.lastPurchase = date;
    } else {
      map.set(key, {
        id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${index}`,
        name: name || 'Name missing',
        email,
        orders: orderId ? 1 : 0,
        totalSpend: amount,
        lastPurchase: date,
        consent,
        status: email ? 'Needs Review' : isEtsySale ? 'Needs Review' : 'Do Not Contact',
        notes: rowNotes || (email ? 'Imported from authorized file. Consent must be verified before marketing use.' : isEtsySale ? 'Etsy sales email found, but no buyer email was detected in the exported message source.' : 'No valid email found.'),
        duplicates: 0,
        sourceRows: 1,
        sources: new Set([fileName]),
        orderIds: orderId ? [orderId] : [],
        sourceType,
        subject,
        etsyLinks,
        linkReviewStatus: linkReviewStatus || (isEtsySale && !email ? 'Link Review Needed' : email ? 'Email Found' : 'Not Applicable'),
        rawSourceFile,
        isEtsySale
      });
    }
  });

  const customers = Array.from(map.values()).map((customer) => {
    const sources = Array.from(customer.sources).join(', ');
    let status = customer.status;
    if (customer.consent === 'Do Not Contact') status = 'Do Not Contact';
    else if (!customer.email && !customer.isEtsySale) status = 'Do Not Contact';
    else if (customer.consent === 'Marketing Eligible') status = 'Marketing Eligible';
    else if (customer.isEtsySale && (!customer.email || customer.linkReviewStatus === 'Link Review Needed')) status = 'Needs Review';
    else if (customer.duplicates > 0) status = 'Needs Review';
    else status = 'Transactional Only';

    return {
      ...customer,
      status,
      sources,
      totalSpend: Number(customer.totalSpend.toFixed(2))
    };
  });

  return { customers, invalid };
}

function mergeCustomers(existing, incoming) {
  const byEmail = new Map();
  [...existing, ...incoming].forEach((customer) => {
    const key = customer.email || customer.id;
    const current = byEmail.get(key);
    if (!current) {
      byEmail.set(key, { ...customer });
      return;
    }
    byEmail.set(key, {
      ...current,
      name: current.name !== 'Name missing' ? current.name : customer.name,
      orders: Array.from(new Set([...(current.orderIds || []), ...(customer.orderIds || [])].filter(Boolean))).length,
      totalSpend: Number((current.totalSpend + customer.totalSpend).toFixed(2)),
      lastPurchase: [current.lastPurchase, customer.lastPurchase].filter(Boolean).sort().pop() || '',
      duplicates: current.duplicates + customer.duplicates + 1,
      sources: Array.from(new Set(`${current.sources}, ${customer.sources}`.split(',').map((s) => s.trim()).filter(Boolean))).join(', '),
      sourceType: current.sourceType || customer.sourceType || '',
      subject: current.subject || customer.subject || '',
      etsyLinks: Array.from(new Set(`${current.etsyLinks || ''} | ${customer.etsyLinks || ''}`.split('|').map((s) => s.trim()).filter(Boolean))).join(' | '),
      linkReviewStatus: current.linkReviewStatus === 'Email Found' || customer.linkReviewStatus === 'Email Found' ? 'Email Found' : (current.linkReviewStatus || customer.linkReviewStatus || 'Not Applicable'),
      isEtsySale: Boolean(current.isEtsySale || customer.isEtsySale),
      consent: current.consent === 'Do Not Contact' || customer.consent === 'Do Not Contact'
        ? 'Do Not Contact'
        : current.consent === 'Marketing Eligible' || customer.consent === 'Marketing Eligible'
          ? 'Marketing Eligible'
          : 'Unknown',
      status: current.status === 'Do Not Contact' || customer.status === 'Do Not Contact'
        ? 'Do Not Contact'
        : current.status === 'Marketing Eligible' || customer.status === 'Marketing Eligible'
          ? 'Marketing Eligible'
          : 'Needs Review'
    });
  });
  return Array.from(byEmail.values());
}

function csvEscape(value) {
  const text = String(value ?? '');
  if (/[",\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
  return text;
}

function downloadText(filename, text, mime = 'text/plain') {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function customersToCsv(customers) {
  const headers = ['name', 'email', 'orders', 'totalSpend', 'lastPurchase', 'consent', 'status', 'sourceType', 'subject', 'etsyLinks', 'linkReviewStatus', 'isEtsySale', 'sources', 'notes'];
  const lines = [headers.join(',')];
  customers.forEach((customer) => {
    lines.push(headers.map((header) => csvEscape(customer[header])).join(','));
  });
  return lines.join('\n');
}

function auditMarkdown(customers, importHistory) {
  const total = customers.length;
  const marketing = customers.filter((c) => c.status === 'Marketing Eligible').length;
  const review = customers.filter((c) => c.status === 'Needs Review').length;
  const dnc = customers.filter((c) => c.status === 'Do Not Contact').length;
  const transactional = customers.filter((c) => c.status === 'Transactional Only').length;
  const duplicates = customers.reduce((sum, c) => sum + c.duplicates, 0);

  return `# Customer Recovery Audit Report\n\nGenerated: ${new Date().toLocaleString()}\n\n## Summary\n\n- Unique customer records: ${total}\n- Marketing eligible: ${marketing}\n- Transactional/support only: ${transactional}\n- Needs manual review: ${review}\n- Do not contact: ${dnc}\n- Duplicate rows collapsed: ${duplicates}\n\n## Imported Files\n\n${importHistory.length ? importHistory.map((item) => `- ${item.file}: ${item.rows} rows processed, ${item.customers} customer records created`).join('\n') : '- No files imported yet.'}\n\n## Compliance Note\n\nThis report is based on files provided by the operator. Marketing eligibility should be verified against consent records, platform terms, and applicable email/privacy law before any outreach.\n`;
}

function StatCard({ label, value, tone = 'neutral' }) {
  return (
    <section className={`stat-card ${tone}`}>
      <span>{label}</span>
      <strong>{value}</strong>
    </section>
  );
}

function App() {
  const [activeTab, setActiveTab] = useState('home');
  const [customers, setCustomers] = useState(() => {
    try {
      return JSON.parse(localStorage.getItem('crc.customers') || '[]');
    } catch {
      return [];
    }
  });
  const [importHistory, setImportHistory] = useState(() => {
    try {
      return JSON.parse(localStorage.getItem('crc.importHistory') || '[]');
    } catch {
      return [];
    }
  });
  const [query, setQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState('All');
  const [scanMode, setScanMode] = useState('standard-etsy');
  const [advancedSearchInput, setAdvancedSearchInput] = useState('');
  const [advancedSelectedTerms, setAdvancedSelectedTerms] = useState([]);
  const [advancedMatchMode, setAdvancedMatchMode] = useState('any');
  const [isDragging, setIsDragging] = useState(false);
  const [lastMessage, setLastMessage] = useState('Ready to import a Gmail Takeout .mbox or authorized Etsy customer file.');
  const [importProgress, setImportProgress] = useState(null);
  const [selectedExports, setSelectedExports] = useState({
    recoveredEmails: true,
    etsyLinks: true,
    missingEmail: true,
    audit: true,
    dnc: false,
    transactional: false,
    marketing: false,
    full: false
  });

  useEffect(() => {
    localStorage.setItem('crc.customers', JSON.stringify(customers));
  }, [customers]);

  useEffect(() => {
    localStorage.setItem('crc.importHistory', JSON.stringify(importHistory));
  }, [importHistory]);

  useEffect(() => {
    if (
      typeof window === 'undefined' ||
      !window.crcApp ||
      typeof window.crcApp.onMboxImportProgress !== 'function'
    ) {
      return undefined;
    }

    return window.crcApp.onMboxImportProgress((progress) => {
      setImportProgress(progress);
      if (progress?.message) {
        setLastMessage(progress.message);
      }
    });
  }, []);

  useEffect(() => {
    if (
      typeof window === 'undefined' ||
      !window.crcApp ||
      typeof window.crcApp.onMboxImportComplete !== 'function'
    ) {
      return undefined;
    }

    return window.crcApp.onMboxImportComplete((result) => {
      handleDesktopMboxImportComplete(result);
    });
  }, []);


  function resetProject() {
    if (!confirm('Reset this local project and remove imported records from this device?')) return;
    setCustomers([]);
    setImportHistory([]);
    localStorage.removeItem('crc.customers');
    localStorage.removeItem('crc.importHistory');
    setLastMessage('Local project reset. Ready for a new recovery import.');
    setActiveTab('home');
  }

  const advancedSearchSuggestions = [
    'law',
    'legal',
    'attorney',
    'refund',
    'shipping',
    'tracking',
    'invoice',
    'payment',
    'order',
    'customer',
    'gmail',
    'yahoo'
  ];

  function parseAdvancedSearchTerms(value) {
    return String(value || '')
      .split(',')
      .map((term) => term.trim())
      .filter(Boolean);
  }

  function uniqueTerms(terms) {
    return Array.from(new Set(terms.map((term) => term.trim()).filter(Boolean)));
  }

  const advancedSearchTerms = uniqueTerms([
    ...parseAdvancedSearchTerms(advancedSearchInput),
    ...advancedSelectedTerms
  ]);

  function toggleAdvancedSearchTerm(term) {
    setAdvancedSelectedTerms((prev) => {
      if (prev.includes(term)) {
        return prev.filter((item) => item !== term);
      }
      return [...prev, term];
    });
  }

  function buildScanOptions() {
    return {
      scanMode,
      advancedSearch: {
        terms: advancedSearchTerms,
        matchMode: advancedMatchMode,
        includeHeaders: true,
        includeBody: true
      }
    };
  }

  const stats = useMemo(() => {
    const totalSpend = customers.reduce((sum, c) => sum + c.totalSpend, 0);
    const latestDesktopImport = [...importHistory].reverse().find((entry) => entry.exportPath) || null;
    const fullRecoveredCount = latestDesktopImport?.rows ?? customers.length;
    const previewRows = latestDesktopImport?.customers ?? customers.length;
    const hasFullDesktopExport = Boolean(latestDesktopImport?.exportPath);

    return {
      total: fullRecoveredCount,
      previewRows,
      fullRecoveredCount,
      fullExportPath: latestDesktopImport?.exportPath || '',
      latestResultsDir: latestDesktopImport?.resultsDir || '',
      latestReadmePath: latestDesktopImport?.readmePath || '',
      latestAuditPath: latestDesktopImport?.auditPath || '',
      latestLinkPreviewPath: latestDesktopImport?.linkPreviewPath || '',
      latestScanMode: latestDesktopImport?.scanMode || 'standard-etsy',
      hasFullDesktopExport,
      marketing: customers.filter((c) => c.status === 'Marketing Eligible').length,
      review: customers.filter((c) => c.status === 'Needs Review').length,
      dnc: customers.filter((c) => c.status === 'Do Not Contact').length,
      transactional: hasFullDesktopExport ? fullRecoveredCount : customers.filter((c) => c.status === 'Transactional Only').length,
      duplicates: customers.reduce((sum, c) => sum + c.duplicates, 0),
      etsySales: hasFullDesktopExport ? fullRecoveredCount : customers.filter((c) => c.isEtsySale).length,
      linkReview: customers.filter((c) => c.isEtsySale && (!c.email || c.linkReviewStatus === 'Link Review Needed')).length,
      recoveredEmails: hasFullDesktopExport ? fullRecoveredCount : customers.filter((c) => c.isEtsySale && c.email).length,
      missingEmails: customers.filter((c) => c.isEtsySale && !c.email).length,
      linksCaptured: customers.filter((c) => c.isEtsySale && c.etsyLinks).length,
      auditReports: importHistory.length ? 1 : 0,
      totalSpend: totalSpend.toLocaleString(undefined, { style: 'currency', currency: 'USD' })
    };
  }, [customers, importHistory]);

  const filtered = useMemo(() => {
    return customers.filter((customer) => {
      let matchesStatus = statusFilter === 'All' || customer.status === statusFilter;
      if (statusFilter === 'Etsy Sales Emails') matchesStatus = Boolean(customer.isEtsySale);
      if (statusFilter === 'Email Found') matchesStatus = Boolean(customer.email);
      if (statusFilter === 'Email Missing') matchesStatus = !customer.email;
      if (statusFilter === 'Link Review Needed') matchesStatus = Boolean(customer.isEtsySale && (!customer.email || customer.linkReviewStatus === 'Link Review Needed'));
      const text = `${customer.name} ${customer.email} ${customer.sources} ${customer.notes} ${customer.subject || ''} ${customer.etsyLinks || ''}`.toLowerCase();
      return matchesStatus && text.includes(query.toLowerCase());
    });
  }, [customers, query, statusFilter]);

  async function parseFile(file) {
    const fileName = file.name;
    const lower = fileName.toLowerCase();

    if (lower.endsWith('.csv') || lower.endsWith('.txt')) {
      const text = await file.text();
      return new Promise((resolve, reject) => {
        Papa.parse(text, {
          header: true,
          skipEmptyLines: true,
          complete: (result) => resolve(result.data),
          error: reject
        });
      });
    }

    if (lower.endsWith('.mbox')) {
      if (hasDesktopMboxPicker()) {
        throw new Error('For Gmail Takeout .mbox files, use the “Choose .mbox file” button so the desktop app can read the file safely.');
      }

      const text = await file.text();
      return mboxToRows(text, file.name);
    }

    if (lower.endsWith('.eml')) {
      const text = await file.text();
      return emlToRows(text, file.name);
    }

    if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const firstSheet = workbook.SheetNames[0];
      return XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: '' });
    }

    throw new Error('Unsupported file type. Use CSV, TXT, XLS, XLSX, MBOX, or EML for this MVP.');
  }

  async function handleDesktopMboxImportComplete(result) {
    try {
      if (result?.canceled) {
        setImportProgress(null);
        setLastMessage('Import canceled. No local records were changed.');
        return;
      }

      if (!result?.ok) {
        throw new Error(result?.error || 'The .mbox file could not be imported.');
      }

      const rows = Array.isArray(result.rows)
        ? result.rows
        : mboxToRows(result.text || '', result.fileName);

      const totalRecoveredRows = result.stats?.recoveredRows ?? rows.length;
      const previewRows = result.stats?.previewRows ?? rows.length;
      const rawExportPath = result.exportPath || result.stats?.outputPath || '';

      const { customers: imported } = buildCustomers(rows, result.fileName);

      const etsyLinkPreviewRows = imported.filter((customer) => customer.isEtsySale && customer.etsyLinks);
      const etsyLinkPreviewCsv = customersToCsv(etsyLinkPreviewRows);

      const activeScanMode = result.stats?.scanMode || (typeof scanMode !== 'undefined' ? scanMode : 'standard-etsy');
      const activeSearchTerms = result.stats?.searchTerms || (typeof advancedSearchTerms !== 'undefined' ? advancedSearchTerms : []);
      const activeMatchMode = result.stats?.matchMode || (typeof advancedMatchMode !== 'undefined' ? advancedMatchMode : 'any');

      let exportPath = rawExportPath;
      let resultsDir = '';
      let readmePath = '';
      let auditPath = '';
      let linkPreviewPath = '';

      if (
        rawExportPath &&
        window.crcApp &&
        typeof window.crcApp.prepareResultsPackage === 'function'
      ) {
        try {
          const packageResult = await window.crcApp.prepareResultsPackage({
            sourcePath: rawExportPath,
            recordCount: totalRecoveredRows,
            messagesScanned: result.stats?.messagesScanned || 0,
            previewRows,
            linkPreviewCsv: etsyLinkPreviewCsv,
            linkPreviewCount: etsyLinkPreviewRows.length,
            fileName: result.fileName,
            scanMode: activeScanMode,
            searchTerms: activeSearchTerms,
            matchMode: activeMatchMode
          });

          if (packageResult?.ok) {
            exportPath = packageResult.recoveredCsvPath || rawExportPath;
            resultsDir = packageResult.resultsDir || '';
            readmePath = packageResult.readmePath || '';
            auditPath = packageResult.auditPath || '';
            linkPreviewPath = packageResult.linkPreviewPath || '';

            if (typeof window.crcApp.openResultsFolder === 'function' && resultsDir) {
              window.crcApp.openResultsFolder(resultsDir);
            }
          } else if (packageResult?.canceled) {
            setLastMessage('Recovery completed, but no save folder was selected. Choose Export again to save the result files.');
          }
        } catch (packageError) {
          console.error('Could not prepare selected results folder:', packageError);
        }
      }

      setCustomers(imported);
      setQuery('');
      setStatusFilter('All');

      setImportHistory((prev) => [
        ...prev,
        {
          file: result.fileName,
          rows: totalRecoveredRows,
          customers: imported.length,
          size: result.sizeLabel || '',
          exportPath,
          rawExportPath,
          resultsDir,
          readmePath,
          auditPath,
          linkPreviewPath,
          scanMode: activeScanMode,
          searchTerms: activeSearchTerms,
          matchMode: activeMatchMode
        }
      ]);

      setImportProgress(null);

      const modeLabel = activeScanMode === 'advanced-mailbox'
        ? 'Advanced search complete'
        : 'Import complete';

      const resultCountLabel = activeScanMode === 'advanced-mailbox'
        ? 'advanced search match(es)'
        : 'recovered record(s)';

      setLastMessage(
        `${modeLabel}: ${result.fileName}. Saved ${totalRecoveredRows.toLocaleString()} ${resultCountLabel}. Loaded ${imported.length.toLocaleString()} record(s) into Review preview. Results folder: ${resultsDir || 'selected save location'}. Main CSV: ${exportPath || 'saved locally'}.`
      );

      setActiveTab('review');
    } catch (error) {
      setImportProgress(null);
      setLastMessage(`Could not import .mbox file: ${error.message}`);
    }
  }

  async function handleDesktopMboxImport() {
    if (scanMode === 'advanced-mailbox' && advancedSearchTerms.length === 0) {
      setLastMessage('Enter at least one Advanced Search term before running a full-mailbox search.');
      setActiveTab('import');
      return;
    }

    const scanOptions = buildScanOptions();

    if (!hasDesktopMboxPicker()) {
      setLastMessage('Desktop .mbox importer is not available in this runtime. Use CSV, XLSX, or the web fallback instead.');
      return;
    }

    if (window.crcApp && typeof window.crcApp.startMboxImport === 'function') {
      setImportProgress({
        phase: 'selecting',
        percent: 0,
        messageCount: 0,
        recoveredRows: 0,
        message: scanMode === 'advanced-mailbox'
          ? 'Waiting for you to choose a Gmail Takeout .mbox file for advanced full-mailbox search...'
          : 'Waiting for you to choose a Gmail Takeout .mbox file...'
      });
      setLastMessage(scanMode === 'advanced-mailbox'
        ? `Choose the Gmail Takeout .mbox file. Advanced search will scan all messages for: ${advancedSearchTerms.join(', ')}.`
        : 'Choose the Gmail Takeout .mbox file from your computer. The file will be read locally only.');
      window.crcApp.startMboxImport(scanOptions);
      return;
    }

    try {
      setImportProgress({
        phase: 'selecting',
        percent: 0,
        messageCount: 0,
        recoveredRows: 0,
        message: 'Waiting for you to choose a Gmail Takeout .mbox file...'
      });
      setLastMessage('Choose the Gmail Takeout .mbox file from your computer. The file will be read locally only.');

      const result = await window.crcApp.selectMboxFile(scanOptions);
      handleDesktopMboxImportComplete(result);
    } catch (error) {
      setImportProgress(null);
      setLastMessage(`Could not import .mbox file: ${error.message}`);
    }
  }


  async function handleFiles(files) {
    const list = Array.from(files || []);
    if (!list.length) return;

    for (const file of list) {
      try {
        const rows = await parseFile(file);
        const { customers: imported } = buildCustomers(rows, file.name);
        setCustomers((prev) => mergeCustomers(prev, imported));
        setImportHistory((prev) => [...prev, { file: file.name, rows: rows.length, customers: imported.length }]);
        setLastMessage(`Import complete: ${file.name}. Scanned ${rows.length} row(s), created ${imported.length} recovery record(s). Review recovered emails and link-review items next.`);
      } catch (error) {
        setLastMessage(`Could not import ${file.name}: ${error.message}`);
      }
    }
    setActiveTab('review');
  }

  function loadSample() {
    const { customers: imported } = buildCustomers(SAMPLE_ROWS, 'Sample Etsy Orders.csv');
    setCustomers(imported);
    setImportHistory([{ file: 'Sample Etsy Orders.csv', rows: SAMPLE_ROWS.length, customers: imported.length }]);
    setLastMessage('Loaded sample Etsy recovery project. You can test review queues and multi-export safely.');
    setActiveTab('review');
  }

  function updateCustomer(id, updates) {
    setCustomers((prev) => prev.map((customer) => customer.id === id ? { ...customer, ...updates } : customer));
  }

  async function exportCsv(type, silent = false) {
    if (type === 'recovered-emails' && stats.hasFullDesktopExport) {
      const filename = `etsy-recovered-emails-full-${stats.fullRecoveredCount}-records.csv`;

      if (window.crcApp && typeof window.crcApp.copyFullExportFile === 'function') {
        try {
          const result = await window.crcApp.copyFullExportFile({
            sourcePath: stats.fullExportPath,
            filename,
            count: stats.fullRecoveredCount
          });

          if (!silent) {
            setLastMessage(`Exported full Recovered Etsy Emails CSV with ${stats.fullRecoveredCount.toLocaleString()} records: ${result.path}`);
          }

          return {
            filename: result.path,
            count: stats.fullRecoveredCount,
            copied: true
          };
        } catch (error) {
          if (!silent) {
            setLastMessage(`Could not export full Recovered Etsy Emails CSV: ${error.message}`);
          }

          return {
            filename: stats.fullExportPath,
            count: stats.fullRecoveredCount,
            error: error.message
          };
        }
      }

      if (!silent) {
        setLastMessage(`Full Recovered Etsy Emails CSV already exists with ${stats.fullRecoveredCount.toLocaleString()} records: ${stats.fullExportPath}`);
      }

      return {
        filename: stats.fullExportPath,
        count: stats.fullRecoveredCount,
        existing: true
      };
    }

    let rows = customers;
    let filename = 'customer-recovery-preview.csv';

    if (type === 'etsy-links-preview') {
      if (stats.latestLinkPreviewPath && window.crcApp && typeof window.crcApp.revealPath === 'function') {
        try {
          await window.crcApp.revealPath(stats.latestLinkPreviewPath);
        } catch (error) {
          console.error('Could not reveal Etsy Link Review Preview:', error);
        }

        if (!silent) {
          setLastMessage(`Etsy Link Review Preview is ready: ${stats.latestLinkPreviewPath}`);
        }

        return {
          filename: stats.latestLinkPreviewPath,
          count: stats.linksCaptured,
          existing: true
        };
      }

      rows = customers.filter((c) => c.isEtsySale && c.etsyLinks);
      filename = 'etsy-link-review-preview-500-row-sample.csv';
    }

    if (type === 'recovered-emails') {
      rows = customers.filter((c) => c.isEtsySale && c.email);
      filename = 'etsy-recovered-emails-preview-only.csv';
    }

    downloadText(filename, customersToCsv(rows), 'text/csv');

    if (!silent) {
      setLastMessage(`Exported ${filename} with ${rows.length.toLocaleString()} preview record(s).`);
    }

    return { filename, count: rows.length };
  }


  function exportAudit(silent = false) {
    downloadText('customer-recovery-audit-report.md', auditMarkdown(customers, importHistory), 'text/markdown');
    if (!silent) setLastMessage('Exported customer-recovery-audit-report.md.');
    return { filename: 'customer-recovery-audit-report.md', count: stats.auditReports };
  }

  const exportOptions = [
    {
      key: 'recoveredEmails',
      label: 'Recovered Etsy Emails',
      detail: stats.hasFullDesktopExport
        ? `Full client CSV. Exports all ${stats.fullRecoveredCount.toLocaleString()} recovered records, not just the 500-row preview.`
        : 'Buyer emails found in Etsy sale messages.',
      count: stats.recoveredEmails,
      action: (silent) => exportCsv('recovered-emails', silent)
    },
    {
      key: 'etsyLinks',
      label: 'Etsy Link Review Preview',
      detail: 'Optional QA file only. This count comes from the 500-row Review preview, not the full dataset.',
      count: stats.linksCaptured,
      action: (silent) => exportCsv('etsy-links-preview', silent)
    },
    {
      key: 'audit',
      label: 'Audit Report',
      detail: 'One markdown summary of import counts, preview limits, and safety notes.',
      count: stats.auditReports || 1,
      action: (silent) => exportAudit(silent)
    }
  ];


  function toggleExportOption(key) {
    setSelectedExports((prev) => ({ ...prev, [key]: !prev[key] }));
  }

  async function exportSelectedFiles() {
    const selected = exportOptions.filter((option) => selectedExports[option.key]);

    if (!selected.length) {
      setLastMessage('Select at least one export file before exporting.');
      return;
    }

    const results = await Promise.all(
      selected.map(async (option) => ({
        option,
        result: await option.action(true)
      }))
    );

    const errors = results.filter((item) => item.result?.error);
    if (errors.length) {
      setLastMessage(`Some exports could not be completed: ${errors.map((item) => `${item.option.label}: ${item.result.error}`).join(' | ')}`);
      return;
    }

    const summary = results
      .filter((item) => item.result)
      .map((item) => `${item.option.label}: ${Number(item.result.count || 0).toLocaleString()} record(s) → ${item.result.filename}`)
      .join(' | ');

    setLastMessage(`Export complete. ${summary}`);
  }


  return (
    <div className="app-shell">
      <aside className="sidebar" aria-label="Primary navigation">
        <div className="brand-block">
          <div className="brand-mark">CRC</div>
          <div>
            <h1>Advanced Email Recovery Tool</h1>
            <p>Local mailbox recovery desk</p>
          </div>
        </div>

        <nav className="nav-list">
          {[
            ['home', 'Dashboard'],
            ['import', 'Import'],
            ['review', 'Review'],
            ['export', 'Export']
          ].map(([id, label]) => (
            <button key={id} className={activeTab === id ? 'active' : ''} onClick={() => setActiveTab(id)}>{label}</button>
          ))}
        </nav>

      </aside>

      <main className="main-content">
        <header className="topbar">
          <div>
            <p className="eyebrow">Version 1.4 Full-Mailbox Advanced Search Build</p>
            <h2>{activeTab === 'home' ? 'Dashboard' : activeTab === 'import' ? 'Import Mailbox' : activeTab === 'review' ? 'Review Records' : 'Export Files'}</h2>
          </div>
          <div className="status-stack">
            <div className="status-pill">{lastMessage}</div>
            {importProgress && (
              <div className="global-import-progress" role="status" aria-live="polite">
                <div className="global-progress-top">
                  <strong>Importing mailbox</strong>
                  <span>{Math.min(100, Math.max(0, Number(importProgress.percent) || 0))}%</span>
                </div>
                <div className="global-progress-track" aria-label="Mailbox import progress">
                  <div
                    className="global-progress-fill"
                    style={{
                      width: `${Math.min(100, Math.max(0, Number(importProgress.percent) || 0))}%`
                    }}
                  />
                </div>
                <div className="global-progress-meta">
                  <span>{(importProgress.messageCount || 0).toLocaleString()} messages checked</span>
                  <span>{(importProgress.recoveredRows || 0).toLocaleString()} emails found</span>
                  {importProgress.heapUsedMb ? <span>{importProgress.heapUsedMb} MB memory</span> : null}
                </div>
                <small>{importProgress.message || 'Scanning mailbox locally. Keep this window open.'}</small>
              </div>
            )}
          </div>
        </header>

        {activeTab === 'home' && (
          <section className="panel-grid">
            <div className="hero-panel">
              <p className="eyebrow">Guided local recovery</p>
              <h3>Advanced Email Recovery Tool</h3>
              <p>Import a Gmail Takeout file, recover available Etsy buyer emails, capture Etsy order links, and export a clean review package.</p>
              <div className="workflow-steps" aria-label="Recovery workflow">
                <span>1. Import</span>
                <span>2. Review</span>
                <span>3. Export</span>
              </div>
              <div className="hero-actions">
                <button className="primary" onClick={() => setActiveTab('import')}>Start Recovery</button>
                <button className="secondary" onClick={loadSample}>Try Sample Project</button>
                <button className="secondary" onClick={resetProject}>Reset Local Project</button>
              </div>
            </div>
            <div className="stat-grid">
              <StatCard label="Etsy sales emails found" value={stats.etsySales} />
              <StatCard label="Buyer emails recovered" value={stats.recoveredEmails} tone="good" />
              <StatCard label="Links captured" value={stats.linksCaptured} />
              <StatCard label="Link review needed" value={stats.linkReview} tone="warn" />
              <StatCard label="Duplicates collapsed" value={stats.duplicates} />
              <StatCard label="Do not contact" value={stats.dnc} tone="danger" />
            </div>
          </section>
        )}

        {activeTab === 'import' && (
          <section className="content-card">
            <div className="section-heading">
              <h3>Import Mailbox</h3>
              <p>Choose Standard Etsy Recovery or Advanced Full-Mailbox Search before selecting the Gmail Takeout .mbox file.</p>
            </div>

            <div className="scan-mode-panel">
              <div className="scan-mode-grid">
                <label className={scanMode === 'standard-etsy' ? 'scan-mode-card active' : 'scan-mode-card'}>
                  <input
                    type="radio"
                    name="scanMode"
                    value="standard-etsy"
                    checked={scanMode === 'standard-etsy'}
                    onChange={() => setScanMode('standard-etsy')}
                  />
                  <strong>Standard Etsy Recovery</strong>
                  <span>Finds Etsy buyer/customer emails using the proven recovery logic.</span>
                </label>

                <label className={scanMode === 'advanced-mailbox' ? 'scan-mode-card active' : 'scan-mode-card'}>
                  <input
                    type="radio"
                    name="scanMode"
                    value="advanced-mailbox"
                    checked={scanMode === 'advanced-mailbox'}
                    onChange={() => setScanMode('advanced-mailbox')}
                  />
                  <strong>Advanced Full-Mailbox Search</strong>
                  <span>Scans all mailbox messages for your selected terms, not only Etsy messages.</span>
                </label>
              </div>

              {scanMode === 'advanced-mailbox' && (
                <div className="advanced-mailbox-search">
                  <label>
                    <span>Search terms before scan</span>
                    <input
                      value={advancedSearchInput}
                      onChange={(event) => setAdvancedSearchInput(event.target.value)}
                      placeholder="Example: law, legal, attorney, refund, shipping"
                    />
                  </label>

                  <label>
                    <span>Match mode</span>
                    <select value={advancedMatchMode} onChange={(event) => setAdvancedMatchMode(event.target.value)}>
                      <option value="any">Any selected term</option>
                      <option value="all">All selected terms</option>
                      <option value="exact">Exact phrase / selected phrase</option>
                    </select>
                  </label>

                  <div className="search-chip-section">
                    <span>Suggested terms</span>
                    <div className="search-chips">
                      {advancedSearchSuggestions.map((term) => (
                        <button
                          key={term}
                          type="button"
                          className={advancedSelectedTerms.includes(term) ? 'search-chip active' : 'search-chip'}
                          onClick={() => toggleAdvancedSearchTerm(term)}
                        >
                          {term}
                        </button>
                      ))}
                    </div>
                  </div>

                  <p className="advanced-search-note">
                    Full CSV export is uncapped. The Review portal only shows a safe preview for performance.
                    Active terms: {advancedSearchTerms.length ? advancedSearchTerms.join(', ') : 'none selected yet'}.
                  </p>
                </div>
              )}
            </div>

            <div className="import-actions">
              <button className="primary" type="button" onClick={handleDesktopMboxImport}>
                {scanMode === 'advanced-mailbox' ? 'Choose .mbox and Run Advanced Search' : 'Choose .mbox and Run Etsy Recovery'}
              </button>
              <p>
                {scanMode === 'advanced-mailbox'
                  ? 'Advanced Search scans all mailbox messages for your terms and exports every matching email record found.'
                  : 'Standard Recovery scans Etsy sale emails, extracts buyer emails, and creates a client-ready results folder.'}
              </p>
              {importProgress && (
                <div className="import-progress">
                  <div className="progress-header">
                    <strong>{importProgress.percent ?? 0}% scanned</strong>
                    <small>{(importProgress.recoveredRows || 0).toLocaleString()} emails found</small>
                  </div>
                  <div className="progress-track" aria-label="Mailbox import progress">
                    <div
                      className={`progress-fill ${(Number(importProgress.percent) || 0) <= 0 ? 'is-pending' : ''}`}
                      style={{
                        width: `${Math.min(100, Math.max(0, Number(importProgress.percent) || 0))}%`
                      }}
                    />
                  </div>
                  <span>{importProgress.message || 'Scanning mailbox locally...'}</span>
                  <small>
                    {(importProgress.messageCount || 0).toLocaleString()} messages checked
                    {importProgress.outputPath ? ` · Saving CSV locally` : ''}
                  </small>
                </div>
              )}
            </div>

            <label
              className={`drop-zone ${isDragging ? 'dragging' : ''}`}
              onDragOver={(event) => { event.preventDefault(); setIsDragging(true); }}
              onDragLeave={() => setIsDragging(false)}
              onDrop={(event) => {
                event.preventDefault();
                setIsDragging(false);
                handleFiles(event.dataTransfer.files);
              }}
            >
              <input type="file" multiple accept=".csv,.txt,.xlsx,.xls,.mbox,.eml" onChange={(event) => handleFiles(event.target.files)} />
              <span className="upload-icon">⇪</span>
              <strong>Drop CSV, TXT, XLS, XLSX, or EML files here</strong>
              <small>For Gmail Takeout .mbox files, use the Choose .mbox file button above.</small>
            </label>

            <div className="help-grid">
              <article>
                <strong>Best file</strong>
                <p>Gmail Takeout Mail .mbox. If Takeout gives you a .zip, unzip it first, then use the Choose .mbox file button above.</p>
              </article>
              <article>
                <strong>Privacy posture</strong>
                <p>Files are processed locally in the app. Do not import data the client is not authorized to access.</p>
              </article>
              <article>
                <strong>What happens next</strong>
                <p>After import, review Email Found, Link Review Needed, and Missing Email queues before exporting the package.</p>
              </article>
            </div>
          </section>
        )}

        {activeTab === 'review' && (
          <section className="content-card full-height">
            <div className="section-heading split">
              <div>
                <h3>{stats.latestScanMode === 'advanced-mailbox' ? 'Review advanced search matches' : 'Review recovery queues'}</h3>
                <p>{stats.latestScanMode === 'advanced-mailbox'
                  ? 'Confirm advanced search matches, inspect source context, and mark records before client handoff.'
                  : 'Confirm recovered buyer emails, inspect Etsy links, and mark records before export.'}</p>
              </div>
              <div className="review-controls">
                <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search customers..." />
                <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
                  {['All', 'Etsy Sales Emails', 'Email Found', 'Link Review Needed', 'Email Missing', 'Transactional Only', 'Needs Review', 'Do Not Contact', 'Marketing Eligible'].map((option) => <option key={option}>{option}</option>)}
                </select>
              </div>
            </div>

            <div className="records-list">
              {filtered.length === 0 && <div className="empty-state">No records match this view. Import a file or load the sample project.</div>}
              {filtered.map((customer) => (
                <article key={customer.id} className={`record-card ${customer.status.toLowerCase().replaceAll(' ', '-')}`}>
                  <div className="record-main">
                    <div>
                      <h4>{customer.name}</h4>
                      <p>{customer.email || 'No valid email detected'}</p>
                    </div>
                    <div className="status-stack">
                      <span className="record-status">{customer.status}</span>
                      {customer.isEtsySale && <span className="record-status secondary-status">{customer.linkReviewStatus || 'Review pending'}</span>}
                    </div>
                  </div>
                  <div className="record-meta">
                    <span>Orders: {customer.orders || 'Unknown'}</span>
                    <span>Spend: {customer.totalSpend > 0 ? customer.totalSpend.toLocaleString(undefined, { style: 'currency', currency: 'USD' }) : 'Unknown'}</span>
                    <span>Last purchase: {customer.lastPurchase || 'Unknown'}</span>
                    <span>Duplicates: {customer.duplicates}</span>
                  </div>
                  {customer.subject && <p className="source-line">Subject: {customer.subject}</p>}
                  {customer.etsyLinks && <p className="source-line">Etsy links: {customer.etsyLinks.split(' | ').slice(0, 3).map((link, index) => <a key={link} href={link} target="_blank" rel="noreferrer">Open link {index + 1}</a>)}</p>}
                  <p className="source-line">Source: {customer.sources}</p>
                  <div className="record-actions">
                    {customer.email && <button onClick={() => { navigator.clipboard?.writeText(customer.email); setLastMessage(`Copied ${customer.email} to clipboard.`); }}>Copy Email</button>}
                    {customer.etsyLinks && <button onClick={() => window.open(customer.etsyLinks.split(' | ')[0], '_blank', 'noopener,noreferrer')}>Open Etsy Link</button>}
                    <button onClick={() => updateCustomer(customer.id, { status: 'Transactional Only', consent: 'Unknown', notes: 'Limited to order/support/admin use unless consent is later confirmed.' })}>Transactional Only</button>
                    <button onClick={() => updateCustomer(customer.id, { status: 'Do Not Contact', consent: 'Do Not Contact', notes: 'Marked as suppressed/do-not-contact.' })}>Do Not Contact</button>
                    {customer.isEtsySale && <button onClick={() => updateCustomer(customer.id, { linkReviewStatus: 'Reviewed - No Email Available', notes: 'Operator reviewed Etsy link; no buyer email available in exported source.' })}>Reviewed / No Email</button>}
                    <button onClick={() => updateCustomer(customer.id, { status: 'Marketing Eligible', consent: 'Marketing Eligible', notes: 'Operator marked consent as confirmed.' })}>Mark Marketing Eligible</button>
                  </div>
                </article>
              ))}
            </div>
          </section>
        )}

        {activeTab === 'outreach' && (
          <section className="content-card">
            <div className="section-heading">
              <h3>Outreach prep</h3>
              <p>This screen prepares responsible re-engagement. It does not send email.</p>
            </div>
            <div className="template-grid">
              <article>
                <strong>Consent confirmation</strong>
                <p>Use for contacts where prior marketing consent is unclear. Ask them to opt in before sending promotional messages.</p>
              </article>
              <article>
                <strong>Store relaunch update</strong>
                <p>Use only for contacts marked Marketing Eligible, or after legal review confirms a proper basis for outreach.</p>
              </article>
              <article>
                <strong>Support-only follow-up</strong>
                <p>Use for order/admin matters. Keep it transactional, specific, and non-promotional.</p>
              </article>
            </div>
          </section>
        )}

        {activeTab === 'export' && (
          <section className="content-card">
            <div className="section-heading">
              <h3>Export recovery package</h3>
              <p>When recovery finishes, choose where to save the Etsy Email Recovery Results folder.</p>
            </div>
            <div className="export-options">
              {exportOptions.map((option) => (
                <label key={option.key} className={`export-option ${selectedExports[option.key] ? 'selected' : ''}`}>
                  <input type="checkbox" checked={Boolean(selectedExports[option.key])} onChange={() => toggleExportOption(option.key)} />
                  <span>
                    <strong>{option.label}</strong>
                    <small>{option.detail}</small>
                  </span>
                  <em>{option.count}</em>
                </label>
              ))}
            </div>
            <div className="button-row">
              <button className="primary export-selected-button" onClick={exportSelectedFiles}>Open Selected Result Files</button>
              <button className="secondary" onClick={() => setSelectedExports({
                recoveredEmails: true,
                etsyLinks: true,
                missingEmail: true,
                audit: true,
                dnc: false,
                transactional: false,
                marketing: false,
                full: false
              })}>Recommended Set</button>
              <button className="secondary" onClick={() => setSelectedExports({
                recoveredEmails: true,
                etsyLinks: true,
                missingEmail: true,
                audit: true,
                dnc: true,
                transactional: true,
                marketing: false,
                full: true
              })}>Select All Safe Exports</button>
            </div>
            
          </section>
        )}
      </main>

      <nav className="mobile-nav" aria-label="Mobile navigation">
        {[
          ['home', 'Home'],
          ['import', 'Import'],
          ['review', 'Review'],
          ['export', 'Export']
        ].map(([id, label]) => (
          <button key={id} className={activeTab === id ? 'active' : ''} onClick={() => setActiveTab(id)}>{label}</button>
        ))}
      </nav>
    </div>
  );
}

createRoot(document.getElementById('root')).render(<App />);
