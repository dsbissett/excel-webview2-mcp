import assert from 'node:assert';
import {describe, it} from 'node:test';

import {excelActiveRange, excelContextInfo} from '../../src/tools/excel.js';
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

  describe('excel_active_range', () => {
    it('returns an error when Excel is not on the target', async () => {
      await withMcpContext(async (response, context) => {
        const page = context.getSelectedPptrPage();
        await page.setContent(html`<main>no excel</main>`);

        await excelActiveRange.handler(
          {
            params: {},
            page: context.getSelectedMcpPage(),
          },
          response,
          context,
        );

        assert.ok(
          response.responseLines.some(line =>
            line.includes('Excel API not available'),
          ),
          `expected error message, got ${JSON.stringify(response.responseLines)}`,
        );
      });
    });

    it('returns address, dimensions, and values from Excel.run', async () => {
      await withMcpContext(async (response, context) => {
        const page = context.getSelectedPptrPage();
        await page.setContent(html`<main>excel fixture</main>`);
        await page.evaluate(() => {
          const globalObject = globalThis as typeof globalThis & {
            Excel?: unknown;
          };
          globalObject.Excel = {
            run: async (
              batch: (ctx: {
                workbook: {getSelectedRange: () => unknown};
                sync: () => Promise<void>;
              }) => Promise<unknown>,
            ) => {
              const range = {
                load: () => undefined,
                address: 'Sheet1!A1:B2',
                values: [
                  [1, 2],
                  [3, 4],
                ],
                formulas: [
                  ['=1', '=2'],
                  ['=3', '=4'],
                ],
                numberFormat: [
                  ['General', 'General'],
                  ['General', 'General'],
                ],
                rowCount: 2,
                columnCount: 2,
              };
              const ctx = {
                workbook: {getSelectedRange: () => range},
                sync: async () => undefined,
              };
              return batch(ctx);
            },
          };
        });

        await excelActiveRange.handler(
          {
            params: {includeFormulas: true, includeNumberFormat: true},
            page: context.getSelectedMcpPage(),
          },
          response,
          context,
        );

        const jsonLine = response.responseLines.find(line =>
          line.startsWith('{'),
        );
        const payload = JSON.parse(jsonLine ?? 'null') as {
          address: string;
          rowCount: number;
          columnCount: number;
          values: unknown[][];
          formulas: string[][];
          numberFormat: string[][];
          truncated: boolean;
        };

        assert.strictEqual(payload.address, 'Sheet1!A1:B2');
        assert.strictEqual(payload.rowCount, 2);
        assert.strictEqual(payload.columnCount, 2);
        assert.deepStrictEqual(payload.values, [
          [1, 2],
          [3, 4],
        ]);
        assert.deepStrictEqual(payload.formulas, [
          ['=1', '=2'],
          ['=3', '=4'],
        ]);
        assert.strictEqual(payload.truncated, false);
      });
    });

    it('truncates values when the range exceeds the cell cap', async () => {
      await withMcpContext(async (response, context) => {
        const page = context.getSelectedPptrPage();
        await page.setContent(html`<main>excel fixture</main>`);
        await page.evaluate(() => {
          const globalObject = globalThis as typeof globalThis & {
            Excel?: unknown;
          };
          globalObject.Excel = {
            run: async (
              batch: (ctx: {
                workbook: {getSelectedRange: () => unknown};
                sync: () => Promise<void>;
              }) => Promise<unknown>,
            ) => {
              const bigValues = Array.from({length: 200}, (_, r) =>
                Array.from({length: 200}, (_, c) => r * 200 + c),
              );
              const range = {
                load: () => undefined,
                address: 'Sheet1!A1:GR200',
                values: bigValues,
                formulas: bigValues,
                numberFormat: bigValues,
                rowCount: 200,
                columnCount: 200,
              };
              const ctx = {
                workbook: {getSelectedRange: () => range},
                sync: async () => undefined,
              };
              return batch(ctx);
            },
          };
        });

        await excelActiveRange.handler(
          {
            params: {},
            page: context.getSelectedMcpPage(),
          },
          response,
          context,
        );

        assert.ok(
          response.responseLines.some(
            line => line.includes('truncated') && line.includes('200x200'),
          ),
          `expected truncation notice, got ${JSON.stringify(response.responseLines)}`,
        );

        const jsonLine = response.responseLines.find(line =>
          line.startsWith('{'),
        );
        const payload = JSON.parse(jsonLine ?? 'null') as {
          truncated: boolean;
          values: unknown[][];
        };
        assert.strictEqual(payload.truncated, true);
        assert.strictEqual(payload.values.length, 1);
        assert.strictEqual(payload.values[0]?.length, 1);
      });
    });
  });
});
