/**
 * @license
 * Copyright 2026 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import assert from 'node:assert';
import {describe, it} from 'node:test';

import {excelContextInfo} from '../../src/tools/excel.js';
import {html, withMcpContext} from '../utils.js';

describe('excel', () => {
  describe('excel_context_info', () => {
    it('returns false when Office.js is not available on the target', async () => {
      await withMcpContext(async (response, context) => {
        const page = context.getSelectedPptrPage();
        await page.setContent(html`<main>No Office.js here</main>`);

        await excelContextInfo.handler(
          {
            params: {},
            page: context.getSelectedMcpPage(),
          },
          response,
          context,
        );

        assert.strictEqual(response.snapshotParams, undefined);

        const payload = JSON.parse(response.responseLines[0] ?? 'null') as {
          hasOfficeGlobal: boolean;
          hasExcelGlobal: boolean;
          hostInfo?: unknown;
          contentLanguage?: unknown;
          displayLanguage?: unknown;
          requirementSets: string[];
        };

        assert.strictEqual(payload.hasOfficeGlobal, false);
        assert.strictEqual(payload.hasExcelGlobal, false);
        assert.ok(!('hostInfo' in payload));
        assert.ok(!('contentLanguage' in payload));
        assert.ok(!('displayLanguage' in payload));
        assert.deepStrictEqual(payload.requirementSets, []);
      });
    });

    it('returns Office host details and supported requirement sets', async () => {
      await withMcpContext(async (response, context) => {
        const page = context.getSelectedPptrPage();
        await page.setContent(html`<main>Excel add-in fixture</main>`);
        await page.evaluate(() => {
          const globalObject = globalThis as typeof globalThis & {
            Office?: unknown;
            Excel?: unknown;
          };

          globalObject.Office = {
            onReady: () =>
              Promise.resolve({
                host: 'Excel',
                platform: 'PC',
              }),
            context: {
              diagnostics: {
                host: 'Excel',
                platform: 'PC',
                version: '16.0.12345.1000',
              },
              contentLanguage: 'en-US',
              displayLanguage: 'en-US',
              requirements: {
                isSetSupported: (name: string, version?: string) => {
                  return (
                    (name === 'ExcelApi' && version === '1.10') ||
                    (name === 'DialogApi' && version === '1.2') ||
                    (name === 'SharedRuntime' && version === '1.1')
                  );
                },
              },
            },
          };
          globalObject.Excel = {};
        });

        await excelContextInfo.handler(
          {
            params: {},
            page: context.getSelectedMcpPage(),
          },
          response,
          context,
        );

        const payload = JSON.parse(response.responseLines[0] ?? 'null') as {
          hasOfficeGlobal: boolean;
          hasExcelGlobal: boolean;
          hostInfo?: {
            host?: string;
            platform?: string;
            version?: string;
          };
          contentLanguage?: string;
          displayLanguage?: string;
          requirementSets: string[];
        };

        assert.strictEqual(payload.hasOfficeGlobal, true);
        assert.strictEqual(payload.hasExcelGlobal, true);
        assert.deepStrictEqual(payload.hostInfo, {
          host: 'Excel',
          platform: 'PC',
          version: '16.0.12345.1000',
        });
        assert.strictEqual(payload.contentLanguage, 'en-US');
        assert.strictEqual(payload.displayLanguage, 'en-US');
        assert.deepStrictEqual(payload.requirementSets, [
          'ExcelApi 1.10',
          'SharedRuntime 1.1',
          'DialogApi 1.2',
        ]);
      });
    });
  });
});
