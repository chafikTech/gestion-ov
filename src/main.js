const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { spawn } = require('node:child_process');
const { initDatabase } = require('./database/init');
const workersAPI = require('./backend/workers');
const presenceAPI = require('./backend/presence');
const databaseBackup = require('./backend/databaseBackup');
const settingsStore = require('./backend/settingsStore');
const { getBackendPath: resolveBackendPath } = require('./backend/backendPath');
const {
  getBackendWorkingDirectory,
  getBackendLogPath,
  appendBackendLog
} = require('./backend/backendRuntime');

function loadEnvFile() {
  const envPath = path.join(__dirname, '..', '.env');
  if (!fs.existsSync(envPath)) return;
  const contents = fs.readFileSync(envPath, 'utf8');
  contents.split('\n').forEach(line => {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) return;
    const idx = trimmed.indexOf('=');
    if (idx === -1) return;
    const key = trimmed.slice(0, idx).trim();
    const value = trimmed.slice(idx + 1).trim();
    if (!process.env[key]) {
      process.env[key] = value;
    }
  });
}

loadEnvFile();
const documentsAPI = require('./backend/documents');

let mainWindow;

function getBackendPath() {
  return resolveBackendPath();
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function runBackendPingOnce(backendPath) {
  return new Promise((resolve, reject) => {
    const backendCwd = getBackendWorkingDirectory(backendPath);
    appendBackendLog(`PING start exe="${backendPath}" cwd="${backendCwd}"`);

    const child = spawn(backendPath, ['--ping'], {
      stdio: ['ignore', 'pipe', 'pipe'],
      windowsHide: true,
      cwd: backendCwd
    });

    let stdout = '';
    let stderr = '';

    child.stdout.on('data', (chunk) => {
      const text = chunk.toString('utf8');
      stdout += text;
      const cleaned = text.trim();
      if (cleaned) {
        console.log(`[backend stdout] ${cleaned}`);
      }
    });

    child.stderr.on('data', (chunk) => {
      const text = chunk.toString('utf8');
      stderr += text;
      const cleaned = text.trim();
      if (cleaned) {
        console.error(`[backend stderr] ${cleaned}`);
        appendBackendLog(`PING stderr: ${cleaned}`);
      }
    });

    child.on('error', (error) => {
      appendBackendLog(
        `PING spawn error exe="${backendPath}" cwd="${backendCwd}" ` +
        `code=${error.code || 'UNKNOWN'} message="${error.message}"`
      );
      reject(
        new Error(
          `Failed to spawn backend at "${backendPath}". ` +
          `Code=${error.code || 'UNKNOWN'} Message=${error.message}`
        )
      );
    });

    child.on('close', (code) => {
      if (code !== 0) {
        appendBackendLog(
          `PING failed exe="${backendPath}" exit=${code} stderr="${stderr.trim().slice(0, 1000)}"`
        );
        reject(
          new Error(
            `Backend ping failed (exit ${code}) at "${backendPath}". ` +
            `stderr=${stderr.trim().slice(0, 500)}`
          )
        );
        return;
      }
      const response = stdout.trim();
      try {
        const parsed = response ? JSON.parse(response) : {};
        if (!parsed || parsed.success !== true) {
          throw new Error('Backend ping returned non-success payload');
        }
      } catch (error) {
        appendBackendLog(`PING invalid payload: "${response.slice(0, 500)}"`);
        reject(
          new Error(
            `Backend ping invalid response at "${backendPath}". ` +
            `stdout=${response.slice(0, 500)}`
          )
        );
        return;
      }
      appendBackendLog(`PING success exe="${backendPath}"`);
      resolve(stdout.trim());
    });
  });
}

async function ensureBackendReady() {
  const backendPath = getBackendPath();
  console.log(`[backend] resolved path: ${backendPath}`);
  appendBackendLog(`Backend path resolved: "${backendPath}"`);
  appendBackendLog(`Backend log file: "${getBackendLogPath()}"`);

  if (!fs.existsSync(backendPath)) {
    console.warn(
      `[backend] executable not found at "${backendPath}". ` +
      `Run "npm run build:all" to build/copy backend artifacts for this environment.`
    );
    appendBackendLog(`Backend executable not found: "${backendPath}"`);
    return false;
  }

  const timeoutMs = Number(process.env.GOV_BACKEND_READY_TIMEOUT_MS || 30000);
  const retryDelayMs = Number(process.env.GOV_BACKEND_READY_RETRY_MS || 1500);
  const deadline = Date.now() + timeoutMs;
  let attempt = 0;

  while (Date.now() <= deadline) {
    attempt += 1;
    try {
      await runBackendPingOnce(backendPath);
      console.log(`[backend] ping success (attempt ${attempt})`);
      appendBackendLog(`Backend ready after ${attempt} attempt(s)`);
      return true;
    } catch (error) {
      console.error(`[backend] ping failed (attempt ${attempt}): ${error.message}`);
      appendBackendLog(`Backend not ready on attempt ${attempt}: ${error.message}`);
      if (Date.now() + retryDelayMs > deadline) break;
      await wait(retryDelayMs);
    }
  }

  console.error('[backend] backend check failed. Python-powered features may be unavailable.');
  appendBackendLog('Backend readiness check failed after timeout');
  return false;
}

function createWindow() {
  const iconPath =
    process.platform === 'win32'
      ? path.join(__dirname, '../icon/app.ico')
      : path.join(__dirname, '../icon/11965535.png');

  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    },
    icon: iconPath
  });

  mainWindow.loadFile(path.join(__dirname, 'frontend/index.html'));

  // Open DevTools in development mode
  if (process.argv.includes('--dev')) {
    mainWindow.webContents.openDevTools();
  }

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// Initialize application
app.whenReady().then(async () => {
  try {
    // Validate backend binary (non-blocking for core app startup)
    await ensureBackendReady();

    // Initialize database
    await initDatabase();
    
    // Register IPC handlers
    registerIPCHandlers();
    
    // Create main window
    createWindow();
  } catch (error) {
    console.error('Failed to initialize application:', error);
    app.quit();
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

// Register all IPC handlers
function registerIPCHandlers() {
  // Settings API
  ipcMain.handle('settings:get', async () => {
    return settingsStore.getAppSettings();
  });

  ipcMain.handle('settings:save', async (event, settings) => {
    return settingsStore.saveAppSettings(settings);
  });

  // Workers API
  ipcMain.handle('workers:getAll', async () => {
    return await workersAPI.getAllWorkers();
  });

  ipcMain.handle('workers:getById', async (event, id) => {
    return await workersAPI.getWorkerById(id);
  });

  ipcMain.handle('workers:search', async (event, query) => {
    return await workersAPI.searchWorkers(query);
  });

  ipcMain.handle('workers:create', async (event, workerData) => {
    return await workersAPI.createWorker(workerData);
  });

  ipcMain.handle('workers:update', async (event, id, workerData) => {
    return await workersAPI.updateWorker(id, workerData);
  });

  ipcMain.handle('workers:delete', async (event, id) => {
    return await workersAPI.deleteWorker(id);
  });

  // Presence API
  ipcMain.handle('presence:get', async (event, workerId, year, month) => {
    return await presenceAPI.getPresence(workerId, year, month);
  });

  ipcMain.handle('presence:save', async (event, workerId, year, month, days, attachmentNumber, monthlyAttachmentEntries) => {
    return await presenceAPI.savePresence(workerId, year, month, days, attachmentNumber, monthlyAttachmentEntries);
  });

  ipcMain.handle('presence:calculate', async (event, workerId, year, month) => {
    return await presenceAPI.calculateMonthlyStats(workerId, year, month);
  });

  ipcMain.handle('presence:getQuarterly', async (event, workerId, year, quarter) => {
    return await presenceAPI.getQuarterlyPresence(workerId, year, quarter);
  });

  ipcMain.handle('presence:getAttachmentNumber', async (event, workerId, year, month) => {
    return await presenceAPI.getAttachmentNumber(workerId, year, month);
  });

  ipcMain.handle('presence:saveAttachmentNumber', async (event, workerId, year, month, attachmentNumber) => {
    return await presenceAPI.saveAttachmentNumber(workerId, year, month, attachmentNumber);
  });

  ipcMain.handle('presence:getMonthlyAttachmentNumbers', async (event, year, month) => {
    return await presenceAPI.getMonthlyAttachmentNumbers(year, month);
  });

  ipcMain.handle('presence:saveMonthlyAttachmentNumbers', async (event, year, month, entries) => {
    return await presenceAPI.saveMonthlyAttachmentNumbers(year, month, entries);
  });

  ipcMain.handle('presence:getWorkersForMonth', async (event, year, month) => {
    return await presenceAPI.getWorkersWithAttachmentOrderForMonth(year, month);
  });

  // Documents API
  ipcMain.handle('documents:generateMonthly', async (event, documentType, workerId, year, month) => {
    return await documentsAPI.generateMonthlyDocument(documentType, workerId, year, month);
  });

  ipcMain.handle('documents:generateQuarterly', async (event, documentType, workerId, year, quarter) => {
    return await documentsAPI.generateQuarterlyDocument(documentType, workerId, year, quarter);
  });

  ipcMain.handle('documents:getAllWorkers', async (event, year, month) => {
    return await documentsAPI.getAllWorkersForMonth(year, month);
  });

  ipcMain.handle('documents:getAllQuarterlyWorkers', async (event, year, quarter) => {
    return await documentsAPI.getAllWorkersForQuarter(year, quarter);
  });

  ipcMain.handle('documents:generateMonthlyBatch', async (event, documentType, year, month, outputDir, options) => {
    return await documentsAPI.generateMonthlyDocumentBatch(documentType, year, month, outputDir, options);
  });

  ipcMain.handle('documents:generateQuarterlyBatch', async (event, documentType, year, quarter, outputDir, options) => {
    return await documentsAPI.generateQuarterlyDocumentBatch(documentType, year, quarter, outputDir, options);
  });

  ipcMain.handle('documents:generateCombinedMonthly', async (event, documentType, year, month, outputDir, options) => {
    return await documentsAPI.generateCombinedMonthlyDocument(documentType, year, month, outputDir, options);
  });

  ipcMain.handle('documents:generateCombinedQuarterly', async (event, documentType, year, quarter, outputDir, options) => {
    return await documentsAPI.generateCombinedQuarterlyDocument(documentType, year, quarter, outputDir, options);
  });

  // File Dialog API
  ipcMain.handle('dialog:selectDirectory', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openDirectory', 'createDirectory'],
      title: 'Sélectionner le dossier de destination'
    });
    return result.canceled ? null : result.filePaths[0];
  });

  ipcMain.handle('dialog:saveFile', async (event, options = {}) => {
    const result = await dialog.showSaveDialog(mainWindow, {
      title: options.title || 'Enregistrer le fichier',
      defaultPath: options.defaultPath,
      filters: options.filters || []
    });
    return result.canceled ? null : result.filePath;
  });

  ipcMain.handle('dialog:openFile', async (event, options = {}) => {
    const result = await dialog.showOpenDialog(mainWindow, {
      title: options.title || 'Ouvrir le fichier',
      filters: options.filters || [],
      properties: ['openFile']
    });
    return result.canceled ? null : result.filePaths[0];
  });

  ipcMain.handle('database:export', async (event, filePath) => {
    return await databaseBackup.exportDatabase(filePath);
  });

  ipcMain.handle('database:import', async (event, filePath) => {
    const result = await databaseBackup.importDatabase(filePath);
    setTimeout(() => {
      app.relaunch();
      app.exit(0);
    }, 300);
    return result;
  });
}
