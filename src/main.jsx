import React, { useEffect, useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import './styles.css';

const EMAIL_RE = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;

const FIELD_ALIASES = {
  email: ['email', 'e-mail', 'buyer email', 'customer email', 'ship email', 'contact email'],
  name: ['name', 'buyer', 'buyer name', 'customer', 'customer name', 'ship name', 'recipient'],
  orderId: ['order id', 'order', 'order number', 'receipt id', 'transaction id', 'sale id'],
  amount: ['amount', 'total', 'order total', 'price', 'gross', 'net', 'revenue'],
  date: ['date', 'order date', 'sale date', 'created', 'created at', 'purchase date'],
  consent: ['consent', 'marketing consent', 'opt in', 'opt-in', 'subscribed', 'newsletter']
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

    if (!email) invalid += 1;
    const key = email || `missing-email-${fileName}-${index}`;
    const existing = map.get(key);

    if (existing) {
      existing.orders += orderId ? 1 : 0;
      existing.totalSpend += amount;
      existing.duplicates += 1;
      existing.sources.add(fileName);
      existing.orderIds.push(orderId || `row-${index + 1}`);
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
        status: email ? 'Needs Review' : 'Do Not Contact',
        notes: email ? 'Imported from authorized file. Consent must be verified before marketing use.' : 'No valid email found.',
        duplicates: 0,
        sourceRows: 1,
        sources: new Set([fileName]),
        orderIds: [orderId || `row-${index + 1}`]
      });
    }
  });

  const customers = Array.from(map.values()).map((customer) => {
    const sources = Array.from(customer.sources).join(', ');
    let status = customer.status;
    if (customer.consent === 'Do Not Contact' || !customer.email) status = 'Do Not Contact';
    else if (customer.consent === 'Marketing Eligible') status = 'Marketing Eligible';
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
  const headers = ['name', 'email', 'orders', 'totalSpend', 'lastPurchase', 'consent', 'status', 'sources', 'notes'];
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
      totalSpend: totalSpend.toLocaleString(undefined, { style: 'currency', currency: 'USD' })
    };
  }, [customers]);

  const filtered = useMemo(() => {
    return customers.filter((customer) => {
      const matchesStatus = statusFilter === 'All' || customer.status === statusFilter;
      const text = `${customer.name} ${customer.email} ${customer.sources} ${customer.notes}`.toLowerCase();
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

    if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const firstSheet = workbook.SheetNames[0];
      return XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: '' });
    }

    throw new Error('Unsupported file type. Use CSV, TXT, XLS, or XLSX for this MVP.');
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
            <p>Local cleanup and export desk</p>
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
            <p className="eyebrow">MVP v0.1</p>
            <h2>{activeTab === 'home' ? 'Recovery dashboard' : activeTab.replace('-', ' ')}</h2>
          </div>
          <div className="status-pill">{lastMessage}</div>
        </header>

        {activeTab === 'home' && (
          <section className="panel-grid">
            <div className="hero-panel">
              <p className="eyebrow">Guided local recovery</p>
              <h3>Recover customer records without overwhelming a non-technical client.</h3>
              <p>Import authorized Etsy/customer files, dedupe records, classify contact risk, and export clean lists with an audit trail.</p>
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
              <StatCard label="Duplicates collapsed" value={stats.duplicates} />
              <StatCard label="Tracked spend" value={stats.totalSpend} />
            </div>
          </section>
        )}

        {activeTab === 'import' && (
          <section className="content-card">
            <div className="section-heading">
              <h3>Import customer files</h3>
              <p>Drag in CSV or Excel files. The app automatically looks for names, emails, order IDs, dates, totals, and consent signals.</p>
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
              <input type="file" multiple accept=".csv,.txt,.xlsx,.xls" onChange={(event) => handleFiles(event.target.files)} />
              <span className="upload-icon">⇪</span>
              <strong>Drop files here or tap to choose</strong>
              <small>Supported in v0.1: CSV, TXT, XLS, XLSX</small>
            </label>

            <div className="help-grid">
              <article>
                <strong>Best first file</strong>
                <p>Etsy order CSV or any spreadsheet with customer name and email columns.</p>
              </article>
              <article>
                <strong>Privacy posture</strong>
                <p>Files are processed locally in the app. Do not import data the client is not authorized to access.</p>
              </article>
              <article>
                <strong>Coming later</strong>
                <p>MBOX/EML mailbox parsing, SQLite project storage, and packaged Mac/Windows installers.</p>
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
                  {['All', 'Marketing Eligible', 'Transactional Only', 'Needs Review', 'Do Not Contact'].map((option) => <option key={option}>{option}</option>)}
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
                  <p className="source-line">Source: {customer.sources}</p>
                  <div className="record-actions">
                    <button onClick={() => updateCustomer(customer.id, { status: 'Marketing Eligible', consent: 'Marketing Eligible', notes: 'Operator marked consent as confirmed.' })}>Mark Marketing Eligible</button>
                    <button onClick={() => updateCustomer(customer.id, { status: 'Transactional Only', consent: 'Unknown', notes: 'Limited to order/support/admin use unless consent is later confirmed.' })}>Transactional Only</button>
                    <button onClick={() => updateCustomer(customer.id, { status: 'Do Not Contact', consent: 'Do Not Contact', notes: 'Marked as suppressed/do-not-contact.' })}>Do Not Contact</button>
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
