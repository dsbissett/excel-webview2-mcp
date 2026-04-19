import assert from 'node:assert';
import {mkdtemp, rm, writeFile, mkdir} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import path from 'node:path';
import {describe, it} from 'node:test';

import {detectExcelAddin} from '../../src/launch/detectAddin.js';

describe('detectExcelAddin', () => {
  it('detects a classic XML workbook add-in repo', async () => {
    const root = await mkdtemp(path.join(tmpdir(), 'detect-addin-xml-'));

    try {
      await writeFixture(root, {
        'package.json': JSON.stringify({name: 'excel-addin'}, null, 2),
        'manifest.xml': `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp>
  <Hosts>
    <Host Name="Workbook" />
  </Hosts>
</OfficeApp>`,
        'package-lock.json': '{}',
      });

      const detected = await detectExcelAddin(root);

      assert.deepStrictEqual(detected, {
        root,
        manifestPath: path.join(root, 'manifest.xml'),
        manifestKind: 'xml',
        packageManager: 'npm',
        existingCdpScript: undefined,
      });
    } finally {
      await rm(root, {recursive: true, force: true});
    }
  });

  it('detects a unified JSON workbook add-in repo from a nested cwd', async () => {
    const root = await mkdtemp(path.join(tmpdir(), 'detect-addin-json-'));
    const nestedCwd = path.join(root, 'src', 'client');

    try {
      await writeFixture(root, {
        'package.json': JSON.stringify({name: 'excel-addin'}, null, 2),
        'manifest.json': JSON.stringify(
          {
            extensions: [
              {
                requirements: {
                  scopes: ['workbook'],
                },
              },
            ],
          },
          null,
          2,
        ),
        'pnpm-lock.yaml': 'lockfileVersion: 9.0',
      });
      await mkdir(nestedCwd, {recursive: true});

      const detected = await detectExcelAddin(nestedCwd);

      assert.deepStrictEqual(detected, {
        root,
        manifestPath: path.join(root, 'manifest.json'),
        manifestKind: 'json',
        packageManager: 'pnpm',
        existingCdpScript: undefined,
      });
    } finally {
      await rm(root, {recursive: true, force: true});
    }
  });

  it('returns null for a repo that is not an Excel add-in project', async () => {
    const root = await mkdtemp(path.join(tmpdir(), 'detect-addin-none-'));

    try {
      await writeFixture(root, {
        'package.json': JSON.stringify(
          {
            name: 'not-an-addin',
            devDependencies: {
              'office-addin-debugging': '^0.0.0',
            },
          },
          null,
          2,
        ),
        'manifest.xml': `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
</OfficeApp>`,
      });

      const detected = await detectExcelAddin(root);

      assert.strictEqual(detected, null);
    } finally {
      await rm(root, {recursive: true, force: true});
    }
  });

  it('reports an existing CDP launch script when one is already defined', async () => {
    const root = await mkdtemp(path.join(tmpdir(), 'detect-addin-script-'));

    try {
      await writeFixture(root, {
        'package.json': JSON.stringify(
          {
            name: 'excel-addin',
            scripts: {
              start: 'vite',
              'start:cdp':
                'cross-env WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS="--remote-debugging-port=9222" office-addin-debugging start manifest.xml',
            },
          },
          null,
          2,
        ),
        'manifest.xml': `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp>
  <Hosts>
    <Host Name="Workbook" />
  </Hosts>
</OfficeApp>`,
        'yarn.lock': '# yarn lockfile v1',
      });

      const detected = await detectExcelAddin(root);

      assert.deepStrictEqual(detected, {
        root,
        manifestPath: path.join(root, 'manifest.xml'),
        manifestKind: 'xml',
        packageManager: 'yarn',
        existingCdpScript: 'start:cdp',
      });
    } finally {
      await rm(root, {recursive: true, force: true});
    }
  });
});

async function writeFixture(
  root: string,
  files: Record<string, string>,
): Promise<void> {
  for (const [relativePath, content] of Object.entries(files)) {
    const filePath = path.join(root, relativePath);
    await mkdir(path.dirname(filePath), {recursive: true});
    await writeFile(filePath, content, 'utf8');
  }
}
