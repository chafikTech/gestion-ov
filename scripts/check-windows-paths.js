#!/usr/bin/env node

const { spawnSync } = require('node:child_process');

const FORBIDDEN_CHARS_RE = /[<>:"\\|?*\x00-\x1F]/;
const RESERVED_WINDOWS_NAMES_RE = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$/i;

function fail(message) {
  console.error(`[check-windows-paths] ${message}`);
  process.exit(1);
}

function formatForbiddenCharacter(char) {
  if (char === '"') return '\\"';
  if (char === '\\') return '\\\\';
  const code = char.codePointAt(0);
  if (typeof code === 'number' && code <= 0x1f) {
    return `\\x${code.toString(16).padStart(2, '0')}`;
  }
  return char;
}

function getTrackedPaths() {
  const result = spawnSync('git', ['ls-files', '-z'], {
    encoding: 'utf8',
    stdio: ['ignore', 'pipe', 'pipe']
  });

  if (result.error) {
    fail(`Unable to execute "git ls-files": ${result.error.message}`);
  }
  if (result.status !== 0) {
    fail(
      `Failed to list tracked files (exit ${result.status}). ` +
      `${(result.stderr || '').trim()}`
    );
  }

  return (result.stdout || '').split('\0').filter(Boolean);
}

function getPathIssues(repoPath) {
  const issues = [];
  const segments = repoPath.split('/');

  for (const segment of segments) {
    if (segment.length === 0) {
      issues.push('Contains an empty path segment');
      continue;
    }

    if (/[ .]$/.test(segment)) {
      issues.push(`Segment "${segment}" ends with a space or dot`);
    }

    const forbidden = segment.match(FORBIDDEN_CHARS_RE);
    if (forbidden) {
      issues.push(
        `Segment "${segment}" contains forbidden character ` +
        `"${formatForbiddenCharacter(forbidden[0])}"`
      );
    }

    if (RESERVED_WINDOWS_NAMES_RE.test(segment)) {
      issues.push(`Segment "${segment}" uses a reserved Windows name`);
    }
  }

  return issues;
}

function main() {
  const paths = getTrackedPaths();
  const failures = [];

  for (const repoPath of paths) {
    const issues = getPathIssues(repoPath);
    if (issues.length > 0) {
      failures.push({ repoPath, issues });
    }
  }

  if (failures.length > 0) {
    console.error(
      `[check-windows-paths] Found ${failures.length} Windows-incompatible path(s):`
    );
    for (const failure of failures) {
      console.error(`- ${failure.repoPath}`);
      for (const issue of failure.issues) {
        console.error(`  * ${issue}`);
      }
    }
    process.exit(1);
  }

  console.log(
    `[check-windows-paths] OK: ${paths.length} tracked path(s) are Windows-compatible`
  );
}

main();
