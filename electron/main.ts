import { app, BrowserWindow, ipcMain } from 'electron';
import * as path from 'path';
import * as http from 'http';

let mainWindow: BrowserWindow | null = null;

function isValidTallyHost(host: string): boolean {
  if (!host) return false;
  const h = host.toLowerCase().trim();
  if (h === 'localhost') return true;
  if (h === '127.0.0.1') return true;
  
  // Regex to check IPv4
  const ipv4Regex = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/;
  const match = h.match(ipv4Regex);
  if (!match) return false;
  
  const o1 = parseInt(match[1], 10);
  const o2 = parseInt(match[2], 10);
  const o3 = parseInt(match[3], 10);
  const o4 = parseInt(match[4], 10);
  
  if (o1 > 255 || o2 > 255 || o3 > 255 || o4 > 255) return false;
  
  // Loopback
  if (o1 === 127) return true;
  
  // Private IPs:
  // 10.x.x.x
  if (o1 === 10) return true;
  // 192.168.x.x
  if (o1 === 192 && o2 === 168) return true;
  // 172.16.x.x to 172.31.x.x
  if (o1 === 172 && o2 >= 16 && o2 <= 31) return true;
  
  return false;
}

function postToTally(host: string, port: number, xml: string): Promise<string> {
  return new Promise((resolve, reject) => {
    if (!isValidTallyHost(host)) {
      return reject(new Error(`Security block: Forbidden host '${host}'. Only local or private network hosts are allowed.`));
    }
    
    const options = {
      hostname: host,
      port: port,
      path: '/',
      method: 'POST',
      headers: {
        'Content-Type': 'text/xml; charset=utf-8',
        'Content-Length': Buffer.byteLength(xml, 'utf8')
      },
      timeout: 15000 // 15 seconds timeout
    };
    
    const req = http.request(options, (res) => {
      let data = '';
      res.setEncoding('utf8');
      res.on('data', (chunk) => {
        data += chunk;
      });
      res.on('end', () => {
        resolve(data);
      });
    });
    
    req.on('error', (err) => {
      reject(err);
    });
    
    req.on('timeout', () => {
      req.destroy();
      reject(new Error('Request to Tally timed out.'));
    });
    
    req.write(xml);
    req.end();
  });
}

// Register secure IPC Handlers for Tally connection
ipcMain.handle('tally:testConnection', async (event, { config, xml }) => {
  const { host = 'localhost', port = 9000 } = config || {};
  const queryXml = xml || `
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
    <TYPE>Data</TYPE>
    <ID>Active Company</ID>
  </HEADER>
  <BODY>
    <DESC></DESC>
  </BODY>
</ENVELOPE>`.trim();
  try {
    const response = await postToTally(host, port, queryXml);
    return { success: true, response };
  } catch (error: any) {
    return { success: false, error: error.message || String(error) };
  }
});

ipcMain.handle('tally:fetchCompany', async (event, { config, xml }) => {
  const { host = 'localhost', port = 9000 } = config || {};
  const queryXml = xml || `
<ENVELOPE>
  <HEADER>
    <VERSION>1</VERSION>
    <TALLYREQUEST>Export Data</TALLYREQUEST>
    <TYPE>Data</TYPE>
    <ID>Active Company</ID>
  </HEADER>
  <BODY>
    <DESC></DESC>
  </BODY>
</ENVELOPE>`.trim();
  try {
    const response = await postToTally(host, port, queryXml);
    return { success: true, response };
  } catch (error: any) {
    return { success: false, error: error.message || String(error) };
  }
});

const tallyChannels = [
  'tally:fetchMasters',
  'tally:fetchLedgers',
  'tally:fetchGroups',
  'tally:fetchStockItems',
  'tally:fetchUnits',
  'tally:fetchDaybook',
  'tally:pushXml'
];

for (const channel of tallyChannels) {
  ipcMain.handle(channel, async (event, { config, xml }) => {
    const { host = 'localhost', port = 9000 } = config || {};
    try {
      const response = await postToTally(host, port, xml);
      return { success: true, response };
    } catch (error: any) {
      return { success: false, error: error.message || String(error) };
    }
  });
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    title: 'TallyGen Pro',
    webPreferences: {
      preload: path.join(__dirname, 'preload.cjs'),
      nodeIntegration: false,
      contextIsolation: true,
    },
  });

  const isDev = !app.isPackaged;
  if (isDev) {
    mainWindow.loadURL('http://localhost:3000');
  } else {
    mainWindow.loadFile(path.join(__dirname, '../dist/index.html'));
  }

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(() => {
  createWindow();

  // Safe check for auto-updater
  try {
    const { autoUpdater } = require('electron-updater');
    if (autoUpdater) {
      autoUpdater.checkForUpdatesAndNotify().catch((err: any) => {
        console.warn('Auto updater check failed:', err);
      });
    }
  } catch (error) {
    console.warn('Auto updater unavailable:', error);
  }

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
