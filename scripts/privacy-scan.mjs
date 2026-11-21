import { execFileSync } from 'node:child_process';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';

export function inspectText(text, personalRoots = []) {
  const findings = [];
  const rules = [
    ['absolute-windows-path', /\b[A-Za-z]:[\\/](?![\\/])/g],
    ['personal-home-path', /\/(?:Users|home)\/[A-Za-z0-9_.-]+\//g],
    ['private-key', /-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----/g],
    ['github-token', /\b(?:gh[pousr]_[A-Za-z0-9]{20,}|github_pat_[A-Za-z0-9_]{30,})\b/g],
    ['credential-value', /(?:api[_-]?key|access[_-]?token|password)\s*[:=]\s*["'](?!["']|example|placeholder|test|process\.|env\.|\$\{)[A-Za-z0-9_+\/-]{24,}["']/gi],
    ['raw-chat-envelope', new RegExp('<send_user_message_' + 'question_reply>|<environment_' + 'context>|"role"\\s*:\\s*"(?:assistant|user)"\\s*,\\s*"content"', 'g')]
  ];
  for (const [rule, pattern] of rules) {
    for (const match of text.matchAll(pattern)) findings.push({ rule, line: text.slice(0, match.index).split('\n').length });
  }
  for (const root of personalRoots.filter(value => value && value.length > 3)) {
    if (text.toLowerCase().includes(root.toLowerCase())) findings.push({ rule: 'current-machine-location', line: 1 });
  }
  return findings;
}

export async function scanRepository(root = process.cwd()) {
  const names = execFileSync('git', ['ls-files', '--cached', '--others', '--exclude-standard', '-z'], { cwd: root, encoding: 'utf8', windowsHide: true })
    .split('\0').filter(Boolean);
  const failures = [];
  let fileCount = 0;
  const personalRoots = [root, os.homedir(), process.env.USERPROFILE, process.env.LOCALAPPDATA].filter(Boolean);
  for (const name of new Set(names)) {
    let bytes;
    try { bytes = await readFile(path.join(root, name)); }
    catch (error) {
      // Git's index still lists tracked files deleted from the working tree.
      // Missing files contain nothing to scan; other read failures stay fatal.
      if (error?.code === 'ENOENT') continue;
      throw error;
    }
    fileCount++;
    if (bytes.includes(0)) {
      if (!/\.(?:gif|png|jpe?g|webp|ico)$/i.test(name)) failures.push({ file: name, rule: 'unexpected-tracked-binary', line: 1 });
      continue;
    }
    for (const finding of inspectText(bytes.toString('utf8'), personalRoots)) failures.push({ file: name, ...finding });
  }
  return { fileCount, failures };
}

if (process.argv[1] && path.resolve(process.argv[1]) === path.resolve(fileURLToPath(import.meta.url))) {
  const result = await scanRepository();
  // Report only filenames/rule IDs, never the potentially sensitive matching text.
  for (const failure of result.failures) console.error(`${failure.file}:${failure.line}: ${failure.rule}`);
  console.log(`Privacy scan: ${result.fileCount} source files; ${result.failures.length} finding(s).`);
  if (result.failures.length) process.exitCode = 1;
}
