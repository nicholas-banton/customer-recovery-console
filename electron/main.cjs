const crcFsPromises = require('fs/promises');
const crcPath = require('path');

const { app, BrowserWindow, shell, protocol, dialog, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');

const isDev = Boolean(process.env.VITE_DEV_SERVER_URL);
const isDebug = process.env.CRC_DEBUG === '1';

const distDir = path.join(__dirname, '..', 'dist');

// Streaming importer settings.
// The app no longer limits total .mbox size. It processes one email message at a time.
// Per-message truncation protects the app from enormous attachment-heavy messages.
const MAX_MESSAGE_CHARS = Number(process.env.CRC_MAX_MESSAGE_CHARS || 512 * 1024);
const rawRecoveryRowLimit = process.env.CRC_MAX_RECOVERY_ROWS;
const parsedRecoveryRowLimit = Number(rawRecoveryRowLimit);
const MAX_RECOVERY_ROWS = Number.isFinite(parsedRecoveryRowLimit) && parsedRecoveryRowLimit > 0
  ? parsedRecoveryRowLimit
  : null;
const PROGRESS_INTERVAL_BYTES = Number(process.env.CRC_PROGRESS_INTERVAL_BYTES || 5 * 1024 * 1024);
const MAX_EMAILS_PER_MESSAGE = Number(process.env.CRC_MAX_EMAILS_PER_MESSAGE || 5);
const MAX_PREVIEW_ROWS = Number(process.env.CRC_MAX_PREVIEW_ROWS || 500);

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2'
};

protocol.registerSchemesAsPrivileged([
  {
    scheme: 'crc',
    privileges: {
      standard: true,
      secure: true,
      supportFetchAPI: true,
      corsEnabled: false
    }
  }
]);

function formatBytes(bytes) {
  if (!Number.isFinite(bytes)) return 'unknown size';
  if (bytes < 1024) return `${bytes} B`;
  const kb = bytes / 1024;
  if (kb < 1024) return `${kb.toFixed(1)} KB`;
  const mb = kb / 1024;
  if (mb < 1024) return `${mb.toFixed(1)} MB`;
  const gb = mb / 1024;
  return `${gb.toFixed(2)} GB`;
}


const RECOVERY_COLUMNS = [
  'Customer Name',
  'Customer Email',
  'Order ID',
  'Order Date',
  'Order Total',
  'Marketing Consent',
  'Source Type',
  'Subject',
  'Raw Source File',
  'Etsy Link',
  'Etsy Recovery Status',
  'Review Status'
];

function csvEscape(value) {
  const text = String(value ?? '');
  const escaped = text.replace(/"/g, '""');
  return /[",\n\r]/.test(escaped) ? `"${escaped}"` : escaped;
}

function writeRecoveryCsvRow(stream, row) {
  stream.write(`${RECOVERY_COLUMNS.map((column) => csvEscape(row[column] || '')).join(',')}\n`);
}

function safeExportBaseName(fileName) {
  return String(fileName || 'mailbox')
    .replace(/\.[^.]+$/, '')
    .replace(/[^a-z0-9._-]+/gi, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 80) || 'mailbox';
}

function timestampForFileName() {
  return new Date().toISOString().replace(/[:.]/g, '-');
}


function importError(message, details = {}) {
  return {
    ok: false,
    canceled: false,
    error: message,
    details
  };
}

function logImport(message, details = {}) {
  if (isDev || isDebug) {
    console.log('[CRC import]', message, details);
  }
}

function emitImportProgress(sender, payload) {
  try {
    if (sender && !sender.isDestroyed()) {
      sender.send('crc:mbox-import-progress', payload);
    }
  } catch {
    // Progress events should never break import.
  }
}

function cleanText(value) {
  return String(value || '')
    .replace(/=\r?\n/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ')
    .trim();
}

function clipText(value, max = 500) {
  const cleaned = cleanText(value);
  return cleaned.length > max ? `${cleaned.slice(0, max)}…` : cleaned;
}

function decodeQuotedPrintableFragments(value) {
  return String(value || '')
    .replace(/=0D=0A/gi, '\n')
    .replace(/=0A/gi, '\n')
    .replace(/=0D/gi, '\n')
    .replace(/=20/gi, ' ')
    .replace(/=09/gi, ' ')
    .replace(/=3D/gi, '=')
    .replace(/=2E/gi, '.')
    .replace(/=2C/gi, ',')
    .replace(/=27/gi, "'")
    .replace(/=22/gi, '"')
    .replace(/=\r?\n/g, '');
}

function normalizePossibleName(value) {
  return decodeQuotedPrintableFragments(value)
    .replace(/<[^>]*>/g, ' ')
    .replace(/https?:\/\/\S+/gi, ' ')
    .replace(/\s+/g, ' ')
    .replace(/^[:\-–—\s]+/, '')
    .replace(/[:\-–—\s]+$/, '')
    .trim();
}

function isLikelyPersonName(value) {
  const name = normalizePossibleName(value);
  const lower = name.toLowerCase();

  if (!name) return false;
  if (name.length < 2 || name.length > 70) return false;
  if (name.includes('@')) return false;
  if (/[=<>{}[\]|\\]/.test(name)) return false;
  if (/\d{4,}/.test(name)) return false;
  if (/[-_=]{4,}/.test(name)) return false;
  if (/\.(png|jpe?g|gif|webp|svg|ico|css|js|json|woff2?)\b/i.test(name)) return false;

  const blockedFragments = [
    'did not leave a note',
    'leave a note',
    'buyer note',
    'message from buyer',
    'order number',
    'receipt',
    'shipping label',
    'conversation with',
    'needs help',
    'they placed',
    'unsubscribe',
    'view this order',
    'download the app',
    'etsy_logo',
    'logo',
    'image',
    'tracking'
  ];

  if (blockedFragments.some((fragment) => lower.includes(fragment))) return false;

  const genericNameLabels = new Set([
    'transactional',
    'transaction',
    'etsy',
    'receipt',
    'order',
    'shipping',
    'notification',
    'support',
    'customer',
    'buyer',
    'seller',
    'shop',
    'message',
    'conversation',
    'unknown'
  ]);

  if (genericNameLabels.has(lower)) return false;

  const words = name.split(/\s+/).filter(Boolean);
  if (words.length > 5) return false;

  // Allow letters, apostrophes, periods, hyphens, and spaces.
  if (!/^[A-Za-zÀ-ÖØ-öø-ÿ.' -]+$/.test(name)) return false;

  return true;
}

function nameFromEmail(email) {
  const local = String(email || '').split('@')[0] || '';
  const cleaned = local
    .replace(/\+.*$/, '')
    .replace(/[._-]+/g, ' ')
    .replace(/\d{4,}/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  if (!cleaned || cleaned.length < 2) {
    return 'Recovered Email Contact';
  }

  const title = cleaned.replace(/\b[a-z]/g, (char) => char.toUpperCase());
  return clipText(title, 70) || 'Recovered Email Contact';
}


function decodeMimeHeader(value) {
  const raw = String(value || '');

  return raw.replace(/=\?([^?]+)\?([QB])\?([^?]+)\?=/gi, (_match, charset, encoding, encoded) => {
    try {
      const normalizedCharset = String(charset || '').toLowerCase();
      if (!normalizedCharset.includes('utf')) return _match;

      if (String(encoding).toUpperCase() === 'B') {
        return Buffer.from(encoded, 'base64').toString('utf8');
      }

      return Buffer.from(
        encoded.replace(/_/g, ' ').replace(/=([0-9A-F]{2})/gi, (_m, hex) => String.fromCharCode(parseInt(hex, 16))),
        'binary'
      ).toString('utf8');
    } catch {
      return _match;
    }
  });
}

function splitHeaderAndBody(message) {
  const normalized = String(message || '');
  const match = normalized.match(/\r?\n\r?\n/);
  if (!match || typeof match.index !== 'number') {
    return { headerText: normalized, bodyText: '' };
  }

  const breakEnd = match.index + match[0].length;
  return {
    headerText: normalized.slice(0, match.index),
    bodyText: normalized.slice(breakEnd)
  };
}

function parseEmailHeaderBlock(message) {
  const { headerText } = splitHeaderAndBody(message);
  const lines = headerText.split(/\r?\n/);
  const unfolded = [];

  for (const line of lines) {
    if (/^\s/.test(line) && unfolded.length) {
      unfolded[unfolded.length - 1] += ` ${line.trim()}`;
    } else {
      unfolded.push(line);
    }
  }

  const headers = {};

  for (const line of unfolded) {
    const index = line.indexOf(':');
    if (index <= 0) continue;

    const key = line.slice(0, index).trim().toLowerCase();
    const value = decodeMimeHeader(line.slice(index + 1).trim());
    headers[key] = value;
  }

  return headers;
}

function unique(values) {
  return Array.from(new Set(values.filter(Boolean)));
}

function isAssetLikeEmail(email) {
  const lower = String(email || '').trim().toLowerCase();
  const [local = '', domain = ''] = lower.split('@');

  if (!local || !domain) return true;

  const domainParts = domain.split('.');
  const tld = domainParts[domainParts.length - 1] || '';

  const blockedTlds = new Set([
    'png',
    'jpg',
    'jpeg',
    'gif',
    'webp',
    'svg',
    'ico',
    'css',
    'js',
    'woff',
    'woff2',
    'map',
    'json'
  ]);

  if (blockedTlds.has(tld)) return true;
  if (/\.(png|jpe?g|gif|webp|svg|ico|css|js|woff2?|map|json)$/i.test(lower)) return true;
  if (/^(1x|2x|3x|4x)\./i.test(domain)) return true;
  if (/(^|[_-])(logo|icon|sprite|avatar|image|photo|thumbnail|asset)([_-]|$)/i.test(local)) return true;

  return false;
}

function filterLikelyBuyerEmails(emails) {
  return unique(emails.map((email) => String(email || '').trim().toLowerCase()))
    .filter((email) => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))
    .filter((email) => !isAssetLikeEmail(email))
    .filter((email) => {
      const lower = email.toLowerCase();
      const domain = lower.split('@')[1] || '';
      const tld = domain.split('.').pop() || '';

      if (lower.includes('@etsy.com')) return false;
      if (lower.includes('example.com')) return false;
      if (!/^[a-z]{2,24}$/i.test(tld)) return false;
      if (/no-?reply|do-?not-?reply|donotreply|notification|transaction|receipt|support|help|mailer-daemon/.test(lower)) return false;

      return true;
    });
}

function extractBuyerEmailsFromBody(bodyText) {
  const body = String(bodyText || '');
  const priority = [];

  const priorityRegexes = [
    /(?:buyer|customer|contact|email|e-mail)[^\n\r<>]{0,120}?([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/gi,
    /mailto:([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/gi
  ];

  for (const regex of priorityRegexes) {
    let match;
    while ((match = regex.exec(body)) !== null) {
      priority.push(match[1]);
    }
  }

  const priorityEmails = filterLikelyBuyerEmails(priority);
  if (priorityEmails.length) return priorityEmails.slice(0, MAX_EMAILS_PER_MESSAGE);

  const allEmails = body.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || [];
  return filterLikelyBuyerEmails(allEmails).slice(0, MAX_EMAILS_PER_MESSAGE);
}

function extractBuyerName(message, email) {
  const text = decodeQuotedPrintableFragments(String(message || ''));

  // Only accept explicit name-style labels. Do not treat generic "Buyer ..."
  // sentence fragments as names.
  const patterns = [
    /(?:Buyer\s+name|Customer\s+name|Recipient\s+name|Ship\s+to|Deliver\s+to|Sold\s+to)\s*:?\s*([^\n\r<]{2,90})/i,
    /(?:Name)\s*:\s*([^\n\r<]{2,90})/i
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);

    if (match?.[1]) {
      const candidate = normalizePossibleName(match[1])
        .replace(/\b(?:buyer email|email|order|receipt|shipping|tracking)\b.*$/i, '')
        .trim();

      if (isLikelyPersonName(candidate)) {
        return clipText(candidate, 70);
      }
    }
  }

  // If no trusted name is found, use a readable email-based contact label.
  // This prevents body fragments like "did not leave a note.=0A..." from
  // appearing above recovered emails.
  if (email) {
    return nameFromEmail(email);
  }

  return 'Unknown Buyer';
}

function extractEtsyLink(message) {
  const text = String(message || '');
  const match = text.match(/https?:\/\/[^\s"'<>]+etsy\.com[^\s"'<>]*/i);
  if (!match) return '';

  return cleanText(match[0])
    .replace(/[),.;]+$/g, '')
    .replace(/&amp;/g, '&');
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

function extractOrderTotal(message) {
  const text = decodeQuotedPrintableFragments(String(message || ''))
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

function inferEtsyOrderId(message, messageIndex) {
  const text = String(message || '');

  const patterns = [
    /Order\s*#?\s*([0-9-]{5,})/i,
    /Order\s+number\s*:?\s*([0-9-]{5,})/i,
    /Receipt\s*#?\s*([0-9-]{5,})/i,
    /\/(?:orders|receipts|sold)\/([0-9-]{5,})/i
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match?.[1]) return match[1];
  }

  return '';
}

function buildEtsyRecoveryRow({ email, messageIndex, headers, message, fileName }) {
  const messageSample = String(message || '').slice(0, MAX_MESSAGE_CHARS);
  const etsyLink = extractEtsyLink(messageSample);
  const buyerName = extractBuyerName(messageSample, email);
  const hasBuyerEmail = Boolean(email);

  return {
    'Customer Name': clipText(buyerName, 120),
    'Customer Email': clipText(email || '', 160),
    'Order ID': clipText(inferEtsyOrderId(messageSample, messageIndex), 80),
    'Order Date': clipText(headers.date || '', 120),
    'Order Total': extractOrderTotal(messageSample),
    'Marketing Consent': 'Unknown',
    'Source Type': 'Etsy Gmail Takeout',
    'Subject': clipText(headers.subject || '', 220),
    'Raw Source File': clipText(fileName, 180),
    'Etsy Link': clipText(etsyLink, 500),
    'Etsy Recovery Status': hasBuyerEmail ? 'Email Found' : 'Link Review Needed',
    'Review Status': hasBuyerEmail ? 'Email Found' : 'Link Review Needed'
  };
}

function rowsFromMboxMessage(message, messageIndex, fileName) {
  const raw = String(message || '');
  if (!raw.trim()) return [];

  const sample = raw.slice(0, MAX_MESSAGE_CHARS);
  const { bodyText } = splitHeaderAndBody(sample);
  const headers = parseEmailHeaderBlock(sample);
  const subject = headers.subject || '';

  const isEtsySale =
    /etsy/i.test(sample) &&
    /(transaction|sale|sold|order|receipt|buyer|purchase|shop|customer)/i.test(`${sample} ${subject}`);

  if (!isEtsySale) return [];

  const buyerEmails = extractBuyerEmailsFromBody(bodyText);

  if (buyerEmails.length) {
    return buyerEmails.map((email) =>
      buildEtsyRecoveryRow({
        email,
        messageIndex,
        headers,
        message: sample,
        fileName
      })
    );
  }

  return [
    buildEtsyRecoveryRow({
      email: '',
      messageIndex,
      headers,
      message: sample,
      fileName
    })
  ];
}

async function streamMboxToRows(filePath, fileName, sizeBytes, sender) {
  return new Promise((resolve, reject) => {
    const previewRows = [];
    const seenEmailKeys = new Set();

    let buffer = '';
    let activeMessage = '';
    let messageIndex = 0;
    let etsyMessageCount = 0;
    let truncatedMessageCount = 0;
    let recoveredRowCount = 0;
    let rowLimitHit = false;
    let lastProgressBytes = 0;
    let messageWasTruncated = false;
    let settled = false;

    const BOUNDARY_TAIL_CHARS = 8192;

    const exportDir = path.join(app.getPath('desktop'), 'Customer Recovery Exports');
    fs.mkdirSync(exportDir, { recursive: true });

    const outputPath = path.join(
      exportDir,
      `${safeExportBaseName(fileName)}-${timestampForFileName()}-recovered-emails.csv`
    );

    const outputStream = fs.createWriteStream(outputPath, { encoding: 'utf8' });
    outputStream.write(`\ufeff${RECOVERY_COLUMNS.map(csvEscape).join(',')}\n`);

    const readStream = fs.createReadStream(filePath, {
      encoding: 'utf8',
      highWaterMark: 64 * 1024
    });

    function finish(payload) {
      if (settled) return;
      settled = true;

      outputStream.end(() => {
        resolve({
          rows: previewRows,
          exportPath: outputPath,
          stats: {
            ...payload,
            recoveredRows: recoveredRowCount,
            previewRows: previewRows.length,
            outputPath,
            rowLimitHit,
            readMethod: 'disk-backed-email-first-parser',
            maxMessageChars: MAX_MESSAGE_CHARS,
            maxRecoveryRows: MAX_RECOVERY_ROWS
          }
        });
      });
    }

    function fail(error) {
      if (settled) return;
      settled = true;

      try {
        outputStream.destroy();
      } catch {}

      reject(error);
    }

    function isLikelyBoundaryAt(source, index) {
      const lineStart = source.startsWith('From ', index)
        ? index
        : source.startsWith('\nFrom ', index)
          ? index + 1
          : -1;

      if (lineStart < 0) return false;

      const lineEnd = source.indexOf('\n', lineStart);
      if (lineEnd < 0) return false;

      const line = source.slice(lineStart, lineEnd);
      return /^From\s+\S+/.test(line) && /\d{4}/.test(line);
    }

    function findBoundary(source, startIndex = 0) {
      if (startIndex <= 0 && isLikelyBoundaryAt(source, 0)) {
        return 0;
      }

      let index = source.indexOf('\nFrom ', Math.max(0, startIndex));

      while (index !== -1) {
        if (isLikelyBoundaryAt(source, index)) {
          return index;
        }

        index = source.indexOf('\nFrom ', index + 1);
      }

      return -1;
    }

    function appendToActiveMessage(segment) {
      if (!segment || rowLimitHit) return;

      const remaining = MAX_MESSAGE_CHARS - activeMessage.length;

      if (remaining > 0) {
        activeMessage += segment.slice(0, remaining);
      }

      if (segment.length > remaining) {
        messageWasTruncated = true;
      }
    }

    function maybeStorePreview(row) {
      if (previewRows.length < MAX_PREVIEW_ROWS) {
        previewRows.push(row);
      }
    }

    function acceptRecoveredRow(row) {
      const email = String(row['Customer Email'] || '').trim().toLowerCase();

      // Email-first recovery mode: do not export or preview link-review/no-email rows.
      // This keeps the Review screen focused on recovered buyer/customer emails.
      if (!email) {
        return;
      }

      if (isAssetLikeEmail(email)) {
        return;
      }

      if (seenEmailKeys.has(email)) {
        return;
      }

      seenEmailKeys.add(email);

      if (MAX_RECOVERY_ROWS && recoveredRowCount >= MAX_RECOVERY_ROWS) {
        rowLimitHit = true;
        return;
      }

      writeRecoveryCsvRow(outputStream, row);
      recoveredRowCount += 1;
      maybeStorePreview(row);

      if (MAX_RECOVERY_ROWS && recoveredRowCount >= MAX_RECOVERY_ROWS) {
        rowLimitHit = true;
      }
    }

    function flushActiveMessage() {
      if (!activeMessage.trim()) {
        activeMessage = '';
        messageWasTruncated = false;
        return;
      }

      messageIndex += 1;

      const message = messageWasTruncated
        ? `${activeMessage}\n\n[CRC note: message body truncated during scan because it exceeded the per-message safety limit.]`
        : activeMessage;

      const extractedRows = rowsFromMboxMessage(message, messageIndex, fileName);

      if (messageWasTruncated) {
        truncatedMessageCount += 1;
      }

      if (extractedRows.length) {
        etsyMessageCount += 1;

        for (const row of extractedRows) {
          acceptRecoveredRow(row);

          if (rowLimitHit) {
            break;
          }
        }
      }

      activeMessage = '';
      messageWasTruncated = false;

      if (rowLimitHit) {
        readStream.destroy();
      }
    }

    function maybeEmitProgress(force = false) {
      const bytesRead = readStream.bytesRead || 0;
      if (!force && bytesRead - lastProgressBytes < PROGRESS_INTERVAL_BYTES) return;

      lastProgressBytes = bytesRead;
      const percent = sizeBytes > 0 ? Math.min(100, Math.round((bytesRead / sizeBytes) * 100)) : 0;
      const heapUsedMb = Math.round(process.memoryUsage().heapUsed / 1024 / 1024);

      emitImportProgress(sender, {
        phase: rowLimitHit ? 'limit_reached' : 'scanning',
        fileName,
        bytesRead,
        sizeBytes,
        percent,
        messageCount: messageIndex,
        recoveredRows: recoveredRowCount,
        heapUsedMb,
        outputPath,
        message: rowLimitHit
          ? `Recovered ${recoveredRowCount.toLocaleString()} unique email record(s). Safe target reached; writing export to disk.`
          : `Scanning ${fileName}: ${percent}% complete, ${messageIndex.toLocaleString()} messages checked, ${recoveredRowCount.toLocaleString()} unique email records found.`
      });

      logImport('Scan progress', {
        fileName,
        percent,
        messagesScanned: messageIndex,
        recoveredRows: recoveredRowCount,
        previewRows: previewRows.length,
        heapUsedMb,
        outputPath
      });
    }

    function consumeBuffer(force = false) {
      if (rowLimitHit) return;

      if (!activeMessage) {
        const firstBoundary = findBoundary(buffer, 0);

        if (firstBoundary === -1) {
          if (force) {
            buffer = '';
          } else if (buffer.length > BOUNDARY_TAIL_CHARS) {
            buffer = buffer.slice(-BOUNDARY_TAIL_CHARS);
          }
          return;
        }

        buffer = buffer.slice(firstBoundary === 0 ? 0 : firstBoundary + 1);
      }

      while (buffer.length) {
        const nextBoundary = findBoundary(buffer, 1);

        if (nextBoundary === -1) {
          if (force) {
            appendToActiveMessage(buffer);
            buffer = '';
          } else if (buffer.length > BOUNDARY_TAIL_CHARS) {
            const safeLength = buffer.length - BOUNDARY_TAIL_CHARS;
            appendToActiveMessage(buffer.slice(0, safeLength));
            buffer = buffer.slice(safeLength);
          }
          break;
        }

        appendToActiveMessage(buffer.slice(0, nextBoundary));
        flushActiveMessage();
        buffer = buffer.slice(nextBoundary + 1);

        if (rowLimitHit) {
          break;
        }
      }
    }

    outputStream.on('error', fail);
    readStream.on('error', fail);

    readStream.on('data', (chunk) => {
      buffer += chunk;
      consumeBuffer(false);
      maybeEmitProgress(false);
    });

    readStream.on('end', () => {
      consumeBuffer(true);
      flushActiveMessage();
      maybeEmitProgress(true);

      finish({
        messagesScanned: messageIndex,
        etsyMessages: etsyMessageCount,
        truncatedMessages: truncatedMessageCount
      });
    });

    readStream.on('close', () => {
      if (rowLimitHit && !settled) {
        consumeBuffer(true);
        flushActiveMessage();
        maybeEmitProgress(true);

        finish({
          messagesScanned: messageIndex,
          etsyMessages: etsyMessageCount,
          truncatedMessages: truncatedMessageCount
        });
      }
    });
  });
}

async function readSelectedMbox(filePath, sender) {
  if (!filePath || typeof filePath !== 'string') {
    return importError('No file was selected.');
  }

  const fileName = path.basename(filePath);
  const ext = path.extname(filePath).toLowerCase();

  if (ext !== '.mbox') {
    return importError('The selected file does not appear to be an .mbox file.', {
      fileName,
      extension: ext || '(none)'
    });
  }

  let stats;

  try {
    stats = await fs.promises.stat(filePath);
  } catch (error) {
    return importError('The selected file could not be found.', {
      fileName,
      code: error.code || 'STAT_FAILED'
    });
  }

  if (!stats.isFile()) {
    return importError('The selected item is not a readable file.', {
      fileName
    });
  }

  if (stats.size <= 0) {
    return importError('The selected .mbox file is empty.', {
      fileName,
      sizeBytes: stats.size
    });
  }

  try {
    await fs.promises.access(filePath, fs.constants.R_OK);
  } catch (error) {
    return importError('The selected .mbox file is not readable by this app.', {
      fileName,
      code: error.code || 'ACCESS_DENIED'
    });
  }

  try {
    logImport('Streaming selected .mbox file', {
      fileName,
      size: formatBytes(stats.size),
      path: filePath
    });

    emitImportProgress(sender, {
      phase: 'starting',
      fileName,
      bytesRead: 0,
      sizeBytes: stats.size,
      percent: 0,
      messageCount: 0,
      recoveredRows: 0,
      heapUsedMb: Math.round(process.memoryUsage().heapUsed / 1024 / 1024),
      message: `File selected: ${fileName} (${formatBytes(stats.size)}). Starting local scan now...`
    });

    const streamed = await streamMboxToRows(filePath, fileName, stats.size, sender);

    logImport('Streaming import complete', {
      fileName,
      size: formatBytes(stats.size),
      stats: streamed.stats
    });

    return {
      ok: true,
      canceled: false,
      fileName,
      sizeBytes: stats.size,
      sizeLabel: formatBytes(stats.size),
      rows: streamed.rows,
      stats: streamed.stats
    };
  } catch (error) {
    return importError('Import failed while streaming the .mbox file.', {
      fileName,
      code: error.code || 'STREAM_FAILED',
      message: error.message
    });
  }
}

function safeDistPath(requestUrl) {
  const url = new URL(requestUrl);
  let pathname = decodeURIComponent(url.pathname || '/index.html');

  if (pathname === '/' || pathname.trim() === '') {
    pathname = '/index.html';
  }

  const requestedPath = path.normalize(path.join(distDir, pathname));

  if (!requestedPath.startsWith(path.normalize(distDir))) {
    return null;
  }

  return requestedPath;
}

function registerLocalAssetProtocol() {
  protocol.handle('crc', async (request) => {
    const requestedPath = safeDistPath(request.url);

    if (!requestedPath) {
      return new Response('Blocked path', { status: 403 });
    }

    try {
      const data = await fs.promises.readFile(requestedPath);
      const ext = path.extname(requestedPath).toLowerCase();

      return new Response(data, {
        status: 200,
        headers: {
          'Content-Type': MIME_TYPES[ext] || 'application/octet-stream'
        }
      });
    } catch (error) {
      console.error('[CRC] Asset load failed:', requestedPath, error.message);
      return new Response('Not found', { status: 404 });
    }
  });
}

function buildMboxOpenDialogOptions() {
  return {
    title: 'Choose Gmail Takeout .mbox file',
    buttonLabel: 'Import .mbox',
    properties: ['openFile'],
    filters: [
      { name: 'Gmail Takeout Mailbox', extensions: ['mbox'] },
      { name: 'All Files', extensions: ['*'] }
    ]
  };
}

function registerImportHandlers() {
  
ipcMain.handle('crc:copy-full-export-file', async (_event, payload = {}) => {
  const sourcePath = String(payload.sourcePath || '').trim();
  const requestedName = String(payload.filename || 'customer-recovery-full-export.csv').trim();
  const count = payload.count || null;

  if (!sourcePath) {
    throw new Error('No source CSV path was provided for the full export copy.');
  }

  const sourceStats = await crcFsPromises.stat(sourcePath);
  if (!sourceStats.isFile()) {
    throw new Error(`Full export source is not a file: ${sourcePath}`);
  }

  const exportDir = crcPath.join(app.getPath('desktop'), 'Customer Recovery Exports');
  await crcFsPromises.mkdir(exportDir, { recursive: true });

  let safeName = requestedName
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, '-')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '')
    .slice(0, 160);

  if (!safeName.toLowerCase().endsWith('.csv')) {
    safeName += '.csv';
  }

  let destinationPath = crcPath.join(exportDir, safeName);

  if (crcPath.resolve(destinationPath) === crcPath.resolve(sourcePath)) {
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    destinationPath = crcPath.join(exportDir, safeName.replace(/\.csv$/i, `-${stamp}.csv`));
  } else {
    try {
      await crcFsPromises.access(destinationPath);
      const stamp = new Date().toISOString().replace(/[:.]/g, '-');
      destinationPath = crcPath.join(exportDir, safeName.replace(/\.csv$/i, `-${stamp}.csv`));
    } catch (_error) {
      // Destination does not exist yet. Safe to use.
    }
  }

  await crcFsPromises.copyFile(sourcePath, destinationPath);

  return {
    ok: true,
    path: destinationPath,
    count,
    sourcePath
  };
});



function crcResultsDir() {
  return crcPath.join(app.getPath('downloads'), 'Etsy Email Recovery Results');
}

function crcSafeCsvName(recordCount) {
  const countLabel = Number(recordCount || 0) > 0 ? Number(recordCount).toLocaleString('en-US').replace(/,/g, '') : 'full';
  return `Recovered-Etsy-Emails-${countLabel}-records.csv`;
}

ipcMain.handle('crc:prepare-results-package', async (_event, payload = {}) => {
  const sourcePath = String(payload.sourcePath || '').trim();
  const recordCount = Number(payload.recordCount || payload.count || 0);
  const messagesScanned = Number(payload.messagesScanned || 0);
  const previewRows = Number(payload.previewRows || 0);

  if (!sourcePath) {
    throw new Error('No recovered CSV path was provided.');
  }

  const sourceStats = await crcFsPromises.stat(sourcePath);
  if (!sourceStats.isFile()) {
    throw new Error(`Recovered CSV source is not a file: ${sourcePath}`);
  }

  const resultsDir = crcResultsDir();
  await crcFsPromises.mkdir(resultsDir, { recursive: true });

  const recoveredCsvName = crcSafeCsvName(recordCount);
  const recoveredCsvPath = crcPath.join(resultsDir, recoveredCsvName);

  await crcFsPromises.copyFile(sourcePath, recoveredCsvPath);

  const readmePath = crcPath.join(resultsDir, 'READ-ME.txt');
  const readme = [
    'Etsy Email Recovery Results',
    '',
    `Recovered email records: ${recordCount.toLocaleString()}`,
    `Messages scanned: ${messagesScanned.toLocaleString()}`,
    `Preview rows shown in app: ${previewRows.toLocaleString()}`,
    '',
    'Main file:',
    recoveredCsvName,
    '',
    'Notes:',
    '- The CSV file is the complete recovered record set.',
    '- The app may show only 500 preview rows for performance.',
    '- Use the CSV file in this folder for the client handoff.',
    '- Marketing consent is not automatically confirmed by this recovery tool.',
    '',
    `Original source export: ${sourcePath}`,
    `Generated: ${new Date().toISOString()}`,
    ''
  ].join('\n');

  await crcFsPromises.writeFile(readmePath, readme, 'utf8');

  const auditPath = crcPath.join(resultsDir, 'Recovery-Audit-Summary.txt');
  const audit = [
    'Recovery Audit Summary',
    '',
    `Recovered records: ${recordCount.toLocaleString()}`,
    `Messages scanned: ${messagesScanned.toLocaleString()}`,
    `Preview rows shown in app: ${previewRows.toLocaleString()}`,
    `Recovered CSV: ${recoveredCsvPath}`,
    `Generated: ${new Date().toISOString()}`,
    ''
  ].join('\n');

  await crcFsPromises.writeFile(auditPath, audit, 'utf8');

  return {
    ok: true,
    resultsDir,
    recoveredCsvPath,
    readmePath,
    auditPath,
    recordCount,
    messagesScanned,
    previewRows
  };
});

ipcMain.handle('crc:open-results-folder', async (_event, folderPath = '') => {
  const target = String(folderPath || crcResultsDir()).trim();
  await crcFsPromises.mkdir(target, { recursive: true });
  return shell.openPath(target);
});

ipcMain.handle('crc:reveal-path', async (_event, targetPath = '') => {
  const target = String(targetPath || '').trim();
  if (!target) {
    throw new Error('No file path was provided.');
  }
  shell.showItemInFolder(target);
  return { ok: true, path: target };
});


ipcMain.handle('crc:select-mbox-file', async (event) => {
    const parentWindow = BrowserWindow.fromWebContents(event.sender);

    const result = await dialog.showOpenDialog(parentWindow, buildMboxOpenDialogOptions());

    if (result.canceled || !result.filePaths || !result.filePaths[0]) {
      return {
        ok: false,
        canceled: true,
        error: 'Import canceled.'
      };
    }

    return readSelectedMbox(result.filePaths[0], event.sender);
  });

  ipcMain.on('crc:start-mbox-import', async (event) => {
    const sender = event.sender;
    const parentWindow = BrowserWindow.fromWebContents(sender);

    try {
      emitImportProgress(sender, {
        phase: 'selecting',
        percent: 0,
        messageCount: 0,
        recoveredRows: 0,
        message: 'Waiting for you to choose a Gmail Takeout .mbox file...'
      });

      const result = await dialog.showOpenDialog(parentWindow, buildMboxOpenDialogOptions());

      if (result.canceled || !result.filePaths || !result.filePaths[0]) {
        sender.send('crc:mbox-import-complete', {
          ok: false,
          canceled: true,
          error: 'Import canceled.'
        });
        return;
      }

      const filePath = result.filePaths[0];

      emitImportProgress(sender, {
        phase: 'selected',
        percent: 0,
        messageCount: 0,
        recoveredRows: 0,
        message: `File selected: ${path.basename(filePath)}. Starting local scan now...`
      });

      const importResult = await readSelectedMbox(filePath, sender);
      sender.send('crc:mbox-import-complete', importResult);
    } catch (error) {
      sender.send('crc:mbox-import-complete', {
        ok: false,
        canceled: false,
        error: error.message || 'Import failed unexpectedly.'
      });
    }
  });
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 320,
    minHeight: 560,
    title: 'Advanced Email Recovery Tool',
    backgroundColor: '#0f172a',
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false
    }
  });

  win.once('ready-to-show', () => {
    win.show();
  });

  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });

  win.webContents.on('did-fail-load', (_event, code, description, url) => {
    console.error('[CRC] Renderer failed to load:', { code, description, url });
  });

  win.webContents.on('render-process-gone', (_event, details) => {
    console.error('[CRC] Renderer process gone:', details);
  });

  win.webContents.on('console-message', (_event, level, message, line, sourceId) => {
    console.log('[CRC renderer]', { level, message, line, sourceId });
  });

  if (isDev) {
    win.loadURL(process.env.VITE_DEV_SERVER_URL);
  } else {
    win.loadURL('crc://app/index.html');
  }

  if (isDebug) {
    win.webContents.openDevTools({ mode: 'detach' });
  }
}

app.whenReady().then(() => {
  registerImportHandlers();

  if (!isDev) {
    registerLocalAssetProtocol();
  }

  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
