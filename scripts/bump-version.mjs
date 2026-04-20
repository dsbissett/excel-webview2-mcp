#!/usr/bin/env node
// Bump the package version across every file that tracks it.
// Usage: npm run bump-version -- <new-version>
//   e.g. npm run bump-version -- 0.0.4

import {readFileSync, writeFileSync} from 'node:fs';
import {fileURLToPath} from 'node:url';
import {dirname, join} from 'node:path';

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..');

const newVersion = process.argv[2];
if (!newVersion || !/^\d+\.\d+\.\d+(-[\w.]+)?$/.test(newVersion)) {
  console.error(
    'Usage: npm run bump-version -- <new-version>  (e.g. 0.0.4 or 1.2.3-beta.1)',
  );
  process.exit(1);
}

const pkgPath = join(repoRoot, 'package.json');
const currentVersion = JSON.parse(readFileSync(pkgPath, 'utf8')).version;

if (currentVersion === newVersion) {
  console.error(`Version is already ${newVersion}; nothing to do.`);
  process.exit(1);
}

/**
 * Each target describes a file + the replacements to apply.
 * `occurrences` is the exact number of `from` strings expected — we fail loud
 * if the file drifts so this script never silently misses a spot.
 */
const targets = [
  {
    file: 'package.json',
    replacements: [
      {
        from: `"version": "${currentVersion}"`,
        to: `"version": "${newVersion}"`,
        occurrences: 1,
      },
    ],
  },
  {
    file: 'package-lock.json',
    replacements: [
      {
        from: `"version": "${currentVersion}"`,
        to: `"version": "${newVersion}"`,
        occurrences: 2,
      },
    ],
  },
  {
    file: 'server.json',
    replacements: [
      {
        from: `"version": "${currentVersion}"`,
        to: `"version": "${newVersion}"`,
        occurrences: 2,
      },
    ],
  },
  {
    file: '.release-please-manifest.json',
    replacements: [
      {
        from: `".": "${currentVersion}"`,
        to: `".": "${newVersion}"`,
        occurrences: 1,
      },
    ],
  },
  {
    file: '.claude-plugin/plugin.json',
    replacements: [
      {
        from: `"version": "${currentVersion}"`,
        to: `"version": "${newVersion}"`,
        occurrences: 1,
      },
    ],
  },
  {
    file: '.claude-plugin/marketplace.json',
    replacements: [
      {
        from: `"version": "${currentVersion}"`,
        to: `"version": "${newVersion}"`,
        occurrences: 1,
      },
    ],
  },
  {
    file: 'src/version.ts',
    replacements: [
      {
        from: `export const VERSION = '${currentVersion}';`,
        to: `export const VERSION = '${newVersion}';`,
        occurrences: 1,
      },
    ],
  },
];

const failures = [];
for (const {file, replacements} of targets) {
  const fullPath = join(repoRoot, file);
  let contents = readFileSync(fullPath, 'utf8');
  for (const {from, to, occurrences} of replacements) {
    const found = contents.split(from).length - 1;
    if (found !== occurrences) {
      failures.push(
        `${file}: expected ${occurrences} occurrence(s) of ${JSON.stringify(from)}, found ${found}`,
      );
      continue;
    }
    contents = contents.split(from).join(to);
  }
  writeFileSync(fullPath, contents);
}

if (failures.length > 0) {
  console.error('Version bump failed — no files written cleanly:');
  for (const f of failures) console.error(`  - ${f}`);
  console.error(
    '\nInvestigate and either fix the target file or update scripts/bump-version.mjs.',
  );
  process.exit(1);
}

console.log(
  `Bumped version ${currentVersion} -> ${newVersion} across ${targets.length} files.`,
);
console.log('Remember to:');
console.log('  1. Add a CHANGELOG.md entry for the new version.');
console.log('  2. Commit the changes.');
console.log(
  `  3. Tag the release commit (e.g. git tag release-${newVersion}).`,
);
