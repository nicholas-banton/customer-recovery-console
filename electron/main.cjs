const { app, BrowserWindow, shell, protocol } = require('electron');
const path = require('path');
const fs = require('fs');

const isDev = Boolean(process.env.VITE_DEV_SERVER_URL);
const isDebug = process.env.CRC_DEBUG === '1';

const distDir = path.join(__dirname, '..', 'dist');

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

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 320,
    minHeight: 560,
    title: 'Customer Recovery Console',
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
