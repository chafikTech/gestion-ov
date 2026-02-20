#!/usr/bin/env node

const fs = require('node:fs');
const path = require('node:path');

const projectRoot = path.resolve(__dirname, '..');
const backendExe = path.join(projectRoot, 'resources', 'bin', 'mybackend', 'mybackend.exe');

function fail(message) {
  console.error(`[assert-windows-backend] ${message}`);
  process.exit(1);
}

function main() {
  if (!fs.existsSync(backendExe)) {
    fail(
      `Missing Windows backend executable: ${backendExe}\n` +
      `Run a Windows backend build first (CI windows-latest or Docker+Wine script).`
    );
  }

  const fd = fs.openSync(backendExe, 'r');
  const header = Buffer.alloc(4);
  fs.readSync(fd, header, 0, 4, 0);
  fs.closeSync(fd);

  // PE executables start with "MZ"
  if (header[0] !== 0x4d || header[1] !== 0x5a) {
    // ELF starts with 0x7F 45 4C 46
    if (header[0] === 0x7f && header[1] === 0x45 && header[2] === 0x4c && header[3] === 0x46) {
      fail(
        `Detected ELF backend instead of Windows PE: ${backendExe}\n` +
        `Do not package Windows installer with Linux backend artifacts.`
      );
    }
    fail(`Unexpected backend executable format at ${backendExe}`);
  }

  console.log(`[assert-windows-backend] OK: ${backendExe}`);
}

main();
