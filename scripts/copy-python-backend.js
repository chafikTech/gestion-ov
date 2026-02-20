#!/usr/bin/env node

const fs = require('node:fs');
const path = require('node:path');

const projectRoot = path.resolve(__dirname, '..');
const targetDir = path.join(projectRoot, 'resources', 'bin');
const sourceDir = path.join(projectRoot, 'dist', 'mybackend');
const sourceExe = path.join(projectRoot, 'dist', 'mybackend.exe');
const sourceBinaryNoExt = path.join(projectRoot, 'dist', 'mybackend');
const targetDirBackend = path.join(targetDir, 'mybackend');
const targetExe = path.join(targetDir, 'mybackend.exe');
const targetBinaryNoExt = path.join(targetDir, 'mybackend');

function fail(message) {
  console.error(`[copy-python-backend] ${message}`);
  process.exit(1);
}

function info(message) {
  console.log(`[copy-python-backend] ${message}`);
}

function main() {
  fs.mkdirSync(targetDir, { recursive: true });

  const hasSourceDir = fs.existsSync(sourceDir) && fs.statSync(sourceDir).isDirectory();
  if (hasSourceDir) {
    fs.rmSync(targetDirBackend, { recursive: true, force: true });
    fs.rmSync(targetExe, { force: true });
    fs.rmSync(targetBinaryNoExt, { force: true });
    fs.cpSync(sourceDir, targetDirBackend, { recursive: true });
    info(
      `Copied ${path.relative(projectRoot, sourceDir)} -> ` +
      `${path.relative(projectRoot, targetDirBackend)}`
    );
    return;
  }

  if (fs.existsSync(sourceExe)) {
    fs.rmSync(targetDirBackend, { recursive: true, force: true });
    fs.copyFileSync(sourceExe, targetExe);
    info(`Copied ${path.relative(projectRoot, sourceExe)} -> ${path.relative(projectRoot, targetExe)}`);
    return;
  }

  if (fs.existsSync(sourceBinaryNoExt) && fs.statSync(sourceBinaryNoExt).isFile()) {
    fs.rmSync(targetDirBackend, { recursive: true, force: true });
    fs.rmSync(targetExe, { force: true });
    fs.copyFileSync(sourceBinaryNoExt, targetBinaryNoExt);
    info(
      `Copied ${path.relative(projectRoot, sourceBinaryNoExt)} -> ` +
      `${path.relative(projectRoot, targetBinaryNoExt)}`
    );
    return;
  }

  fail(
    `Missing ${path.relative(projectRoot, sourceDir)}, ` +
    `${path.relative(projectRoot, sourceExe)} and ${path.relative(projectRoot, sourceBinaryNoExt)}. ` +
    `Run npm run build:py first.`
  );
}

main();
