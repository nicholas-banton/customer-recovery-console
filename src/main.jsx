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

function inferEtsyOrderId(messageText, fallbackIndex) {
  const value = String(messageText || '');

  const urlOrderMatch = value.match(/[?&](?:order_id|receipt_id|transaction_id)=(\d+)/i);
  if (urlOrderMatch) return urlOrderMatch[1];

  const orderMatch = value.match(/(?:order|order_id|receipt|transaction)[^\d]{0,12}(\d{4,})/i);
  if (orderMatch) return orderMatch[1];

  return `etsy-message-${fallbackIndex}`;
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
    'Order Total': '0',
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
        orderIds: [orderId || `row-${index + 1}`],
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
      orders: current.orders + customer.orders,
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
  const [isDragging, setIsDragging] = useState(false);
  const [lastMessage, setLastMessage] = useState('Ready to import authorized customer files.');

  useEffect(() => {
    localStorage.setItem('crc.customers', JSON.stringify(customers));
  }, [customers]);

  useEffect(() => {
    localStorage.setItem('crc.importHistory', JSON.stringify(importHistory));
  }, [importHistory]);

  function resetProject() {
    if (!confirm('Reset this local project and remove imported records from this device?')) return;
    setCustomers([]);
    setImportHistory([]);
    localStorage.removeItem('crc.customers');
    localStorage.removeItem('crc.importHistory');
    setLastMessage('Local project reset. Ready for a new recovery import.');
    setActiveTab('home');
  }

  const stats = useMemo(() => {
    const totalSpend = customers.reduce((sum, c) => sum + c.totalSpend, 0);
    return {
      total: customers.length,
      marketing: customers.filter((c) => c.status === 'Marketing Eligible').length,
      review: customers.filter((c) => c.status === 'Needs Review').length,
      dnc: customers.filter((c) => c.status === 'Do Not Contact').length,
      transactional: customers.filter((c) => c.status === 'Transactional Only').length,
      duplicates: customers.reduce((sum, c) => sum + c.duplicates, 0),
      etsySales: customers.filter((c) => c.isEtsySale).length,
      linkReview: customers.filter((c) => c.isEtsySale && (!c.email || c.linkReviewStatus === 'Link Review Needed')).length,
      totalSpend: totalSpend.toLocaleString(undefined, { style: 'currency', currency: 'USD' })
    };
  }, [customers]);

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

  async function handleFiles(files) {
    const list = Array.from(files || []);
    if (!list.length) return;

    for (const file of list) {
      try {
        const rows = await parseFile(file);
        const { customers: imported } = buildCustomers(rows, file.name);
        setCustomers((prev) => mergeCustomers(prev, imported));
        setImportHistory((prev) => [...prev, { file: file.name, rows: rows.length, customers: imported.length }]);
        setLastMessage(`Imported ${rows.length} rows from ${file.name}. Found ${imported.length} customer records.`);
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
    setLastMessage('Loaded sample project. You can test review, filters, and exports safely.');
    setActiveTab('review');
  }

  function updateCustomer(id, updates) {
    setCustomers((prev) => prev.map((customer) => customer.id === id ? { ...customer, ...updates } : customer));
  }

  function exportCsv(type) {
    let rows = customers;
    let filename = 'customer-recovery-full.csv';
    if (type === 'marketing') {
      rows = customers.filter((c) => c.status === 'Marketing Eligible');
      filename = 'marketing-eligible-customers.csv';
    }
    if (type === 'transactional') {
      rows = customers.filter((c) => c.status === 'Transactional Only');
      filename = 'transactional-support-customers.csv';
    }
    if (type === 'dnc') {
      rows = customers.filter((c) => c.status === 'Do Not Contact');
      filename = 'do-not-contact-list.csv';
    }
    if (type === 'etsy-links') {
      rows = customers.filter((c) => c.isEtsySale && c.etsyLinks);
      filename = 'etsy-link-review-queue.csv';
    }
    if (type === 'recovered-emails') {
      rows = customers.filter((c) => c.isEtsySale && c.email);
      filename = 'etsy-recovered-emails.csv';
    }
    if (type === 'missing-email') {
      rows = customers.filter((c) => c.isEtsySale && !c.email);
      filename = 'etsy-missing-email-review.csv';
    }
    downloadText(filename, customersToCsv(rows), 'text/csv');
  }

  function exportAudit() {
    downloadText('customer-recovery-audit-report.md', auditMarkdown(customers, importHistory), 'text/markdown');
  }

  return (
    <div className="app-shell">
      <aside className="sidebar" aria-label="Primary navigation">
        <div className="brand-block">
          <div className="brand-mark">CRC</div>
          <div>
            <h1>Customer Recovery Console</h1>
            <p>Etsy/Gmail recovery desk</p>
          </div>
        </div>

        <nav className="nav-list">
          {[
            ['home', 'Home'],
            ['import', 'Import'],
            ['review', 'Review'],
            ['outreach', 'Outreach Prep'],
            ['export', 'Export Center']
          ].map(([id, label]) => (
            <button key={id} className={activeTab === id ? 'active' : ''} onClick={() => setActiveTab(id)}>{label}</button>
          ))}
        </nav>

        <div className="safety-card">
          <strong>Safety boundary</strong>
          <p>Only import files your client is legally allowed to provide. This MVP does not scrape Etsy or send bulk email.</p>
        </div>
      </aside>

      <main className="main-content">
        <header className="topbar">
          <div>
            <p className="eyebrow">MVP v0.3 Etsy mode</p>
            <h2>{activeTab === 'home' ? 'Recovery dashboard' : activeTab.replace('-', ' ')}</h2>
          </div>
          <div className="status-pill">{lastMessage}</div>
        </header>

        {activeTab === 'home' && (
          <section className="panel-grid">
            <div className="hero-panel">
              <p className="eyebrow">Guided local recovery</p>
              <h3>Recover Etsy sales contacts from Gmail exports without 22,000 copy/paste actions.</h3>
              <p>Import a Google Takeout MBOX, detect Etsy sales emails, extract available buyer emails and Etsy transaction links, then review and export clean records.</p>
              <div className="hero-actions">
                <button className="primary" onClick={() => setActiveTab('import')}>Start Recovery</button>
                <button className="secondary" onClick={loadSample}>Try Sample Project</button>
                <button className="secondary" onClick={resetProject}>Reset Local Project</button>
              </div>
            </div>
            <div className="stat-grid">
              <StatCard label="Unique records" value={stats.total} />
              <StatCard label="Marketing eligible" value={stats.marketing} tone="good" />
              <StatCard label="Needs review" value={stats.review} tone="warn" />
              <StatCard label="Do not contact" value={stats.dnc} tone="danger" />
              <StatCard label="Etsy sales emails" value={stats.etsySales} />
              <StatCard label="Link review needed" value={stats.linkReview} tone="warn" />
            </div>
          </section>
        )}

        {activeTab === 'import' && (
          <section className="content-card">
            <div className="section-heading">
              <h3>Import Gmail Takeout or Etsy files</h3>
              <p>Drag in a Gmail Takeout .mbox file, Etsy CSV export, Google Contacts CSV, or customer spreadsheet. Etsy sales emails are detected automatically.</p>
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
              <strong>Drop files here or tap to choose</strong>
              <small>Supported: CSV, TXT, XLS, XLSX, MBOX, EML</small>
            </label>

            <div className="help-grid">
              <article>
                <strong>Best first file</strong>
                <p>For this client: Gmail Takeout Mail .mbox filtered around Etsy sales emails.</p>
              </article>
              <article>
                <strong>Privacy posture</strong>
                <p>Files are processed locally in the app. Do not import data the client is not authorized to access.</p>
              </article>
              <article>
                <strong>Coming later</strong>
                <p>Now supports first-pass Gmail Takeout MBOX and EML import. SQLite project storage and PDF reports are planned next.</p>
              </article>
            </div>
          </section>
        )}

        {activeTab === 'review' && (
          <section className="content-card full-height">
            <div className="section-heading split">
              <div>
                <h3>Review customer records</h3>
                <p>Search, filter, and classify records before export.</p>
              </div>
              <div className="review-controls">
                <input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search customers..." />
                <select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
                  {['All', 'Etsy Sales Emails', 'Email Found', 'Email Missing', 'Link Review Needed', 'Marketing Eligible', 'Transactional Only', 'Needs Review', 'Do Not Contact'].map((option) => <option key={option}>{option}</option>)}
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
                    <span className="record-status">{customer.status}</span>
                  </div>
                  <div className="record-meta">
                    <span>Orders: {customer.orders}</span>
                    <span>Spend: {customer.totalSpend.toLocaleString(undefined, { style: 'currency', currency: 'USD' })}</span>
                    <span>Last purchase: {customer.lastPurchase || 'Unknown'}</span>
                    <span>Duplicates: {customer.duplicates}</span>
                  </div>
                  {customer.subject && <p className="source-line">Subject: {customer.subject}</p>}
                  {customer.etsyLinks && <p className="source-line">Etsy links: {customer.etsyLinks.split(' | ').slice(0, 3).map((link, index) => <a key={link} href={link} target="_blank" rel="noreferrer">Open link {index + 1}</a>)}</p>}
                  <p className="source-line">Source: {customer.sources}</p>
                  <div className="record-actions">
                    <button onClick={() => updateCustomer(customer.id, { status: 'Marketing Eligible', consent: 'Marketing Eligible', notes: 'Operator marked consent as confirmed.' })}>Mark Marketing Eligible</button>
                    <button onClick={() => updateCustomer(customer.id, { status: 'Transactional Only', consent: 'Unknown', notes: 'Limited to order/support/admin use unless consent is later confirmed.' })}>Transactional Only</button>
                    <button onClick={() => updateCustomer(customer.id, { status: 'Do Not Contact', consent: 'Do Not Contact', notes: 'Marked as suppressed/do-not-contact.' })}>Do Not Contact</button>
                    {customer.isEtsySale && <button onClick={() => updateCustomer(customer.id, { linkReviewStatus: 'Reviewed - No Email Available', notes: 'Operator reviewed Etsy link; no buyer email available in exported source.' })}>Reviewed / No Email</button>}
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
            <div className="warning-box">
              <strong>Operator warning</strong>
              <p>Do not export marketplace buyers into a marketing platform unless consent or another lawful basis has been confirmed.</p>
            </div>
          </section>
        )}

        {activeTab === 'export' && (
          <section className="content-card">
            <div className="section-heading">
              <h3>Export center</h3>
              <p>Create clean lists and an audit report for client handoff.</p>
            </div>
            <div className="export-grid">
              <button onClick={() => exportCsv('full')}>Export Full Clean CSV</button>
              <button onClick={() => exportCsv('marketing')}>Export Marketing Eligible Only</button>
              <button onClick={() => exportCsv('transactional')}>Export Transactional/Support List</button>
              <button onClick={() => exportCsv('dnc')}>Export Do-Not-Contact List</button>
              <button onClick={() => exportCsv('etsy-links')}>Export Etsy Link Review Queue</button>
              <button onClick={() => exportCsv('recovered-emails')}>Export Etsy Recovered Emails</button>
              <button onClick={() => exportCsv('missing-email')}>Export Etsy Missing Email Review</button>
              <button onClick={exportAudit}>Export Audit Report</button>
            </div>
            <div className="acknowledgment-box">
              <strong>Before client use</strong>
              <p>The operator should verify consent records, platform terms, and applicable email/privacy law before any marketing export is used.</p>
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
