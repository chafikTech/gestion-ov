const fs = require('node:fs');
const path = require('node:path');
const { app } = require('electron');

function getBackendWorkingDirectory(executablePath) {
  return path.dirname(executablePath);
}

function getBackendLogPath() {
  try {
    return path.join(app.getPath('userData'), 'backend.log');
  } catch (error) {
    return path.join(process.cwd(), 'backend.log');
  }
}

function appendBackendLog(message) {
  const logPath = getBackendLogPath();
  try {
    fs.mkdirSync(path.dirname(logPath), { recursive: true });
    fs.appendFileSync(logPath, `[${new Date().toISOString()}] ${message}\n`, 'utf8');
  } catch (error) {
    // Intentionally swallow logging failures to avoid blocking runtime features.
    console.error(`[backend log] append failed: ${error.message}`);
  }
}

module.exports = {
  getBackendWorkingDirectory,
  getBackendLogPath,
  appendBackendLog
};
