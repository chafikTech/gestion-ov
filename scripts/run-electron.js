#!/usr/bin/env node

const { spawn } = require('node:child_process');

const electronBinaryPath = require('electron');
const args = process.argv.slice(2);

const env = { ...process.env };
delete env.ELECTRON_RUN_AS_NODE;

const child = spawn(electronBinaryPath, args, {
  stdio: 'inherit',
  env
});

child.on('error', (error) => {
  console.error('Failed to launch Electron:', error);
  process.exit(1);
});

child.on('exit', (code, signal) => {
  if (typeof code === 'number') process.exit(code);
  if (signal) process.exit(1);
  process.exit(1);
});
