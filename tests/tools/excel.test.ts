import assert from 'node:assert';
import {describe, it} from 'node:test';

import {
  excelActiveRange,
  excelContextInfo,
  excelListNamedItems,
  excelListTables,
  excelListWorksheets,
  excelReadRange,
  excelUsedRange,
  excelWorkbookInfo,
  excelWorksheetInfo,
} from '../../src/tools/excel.js';
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

  describe('excel_workbook_info', () => {
    it('returns workbook metadata from Excel.run', async () => {
      await withMcpContext(async (response, context) => {
        const page = context.getSelectedPptrPage();
        await page.setContent(html`<main>excel fixture</main>`);
        await page.evaluate(() => {
          const g = globalThis as typeof globalThis & {Excel?: unknown};
          g.Excel = {
            run: async (
              batch: (ctx: {
                workbook: unknown;
                sync: () => Promise<void>;
              }) => Promise<unknown>,
            ) => {
              const ctx = {
                workbook: {
                  load: () => undefined,
                  name: 'Book1.xlsx',
                  isDirty: true,
                  readOnly: false,
                  application: {
                    load: () => undefined,
                    calculationMode: 'Automatic',
                    calculationState: 'Done',
                  },
                  protection: {load: () => undefined, protected: false},
                },
                sync: async () => undefined,
              };
              return batch(ctx);
            },
          };
        });

        await excelWorkbookInfo.handler(
          {params: {}, page: context.getSelectedMcpPage()},
          response,
          context,
        );

        const jsonLine = response.responseLines.find(l => l.startsWith('{'));
        const payload = JSON.parse(jsonLine ?? 'null') as {
          name: string;
          isDirty: boolean;
          readOnly: boolean;
          protected: boolean;
          calculationMode: string;
          calculationState: string;
        };
        assert.strictEqual(payload.name, 'Book1.xlsx');
        assert.strictEqual(payload.isDirty, true);
        assert.strictEqual(payload.readOnly, false);
        assert.strictEqual(payload.protected, false);
        assert.strictEqual(payload.calculationMode, 'Automatic');
        assert.strictEqual(payload.calculationState, 'Done');
      });
    });
  });

  describe('excel_list_worksheets', () => {
    it('lists sheets with active flag', async () => {
      await withMcpContext(async (response, context) => {
        const page = context.getSelectedPptrPage();
        await page.setContent(html`<main>excel fixture</main>`);
        await page.evaluate(() => {
          const g = globalThis as typeof globalThis & {Excel?: unknown};
          g.Excel = {
            run: async (batch: (ctx: unknown) => Promise<unknown>) => {
              const sheetsCollection = {
                load: () => undefined,
                items: [
                  {
                    name: 'Sheet1',
                    id: 's1',
                    position: 0,
                    visibility: 'Visible',
                    tabColor: '',
                  },
                  {
                    name: 'Sheet2',
                    id: 's2',
                    position: 1,
                    visibility: 'Hidden',
                    tabColor: '#FF0000',
                  },
                ],
                getActiveWorksheet: () => ({
                  load: () => undefined,
                  id: 's2',
                }),
              };
              const ctx = {
                workbook: {worksheets: sheetsCollection},
                sync: async () => undefined,
              };
              return batch(ctx);
            },
          };
        });

        await excelListWorksheets.handler(
          {params: {}, page: context.getSelectedMcpPage()},
          response,
          context,
        );

        const jsonLine = response.responseLines.find(l => l.startsWith('{'));
        const payload = JSON.parse(jsonLine ?? 'null') as {
          worksheets: Array<{
            name: string;
            active: boolean;
            visibility: string;
          }>;
        };
        assert.strictEqual(payload.worksheets.length, 2);
        assert.strictEqual(payload.worksheets[0]?.name, 'Sheet1');
        assert.strictEqual(payload.worksheets[0]?.active, false);
        assert.strictEqual(payload.worksheets[1]?.name, 'Sheet2');
        assert.strictEqual(payload.worksheets[1]?.active, true);
        assert.strictEqual(payload.worksheets[1]?.visibility, 'Hidden');
      });
    });
  });

  describe('excel_worksheet_info', () => {
    it('returns sheet metadata including used range', async () => {
      await withMcpContext(async (response, context) => {
        const page = context.getSelectedPptrPage();
        await page.setContent(html`<main>excel fixture</main>`);
        await page.evaluate(() => {
          const g = globalThis as typeof globalThis & {Excel?: unknown};
          g.Excel = {
            run: async (batch: (ctx: unknown) => Promise<unknown>) => {
              const ws = {
                load: () => undefined,
                name: 'Sheet1',
                id: 's1',
                position: 0,
                visibility: 'Visible',
                tabColor: '',
                showGridlines: true,
                showHeadings: true,
                standardHeight: 15,
                standardWidth: 8.43,
                getUsedRangeOrNullObject: () => ({
                  load: () => undefined,
                  isNullObject: false,
                  address: 'Sheet1!A1:C3',
                }),
                protection: {load: () => undefined, protected: false},
              };
              const ctx = {
                workbook: {
                  worksheets: {
                    getItem: () => ws,
                    getActiveWorksheet: () => ws,
                  },
                },
                sync: async () => undefined,
              };
              return batch(ctx);
            },
          };
        });

        await excelWorksheetInfo.handler(
          {params: {}, page: context.getSelectedMcpPage()},
          response,
          context,
        );

        const jsonLine = response.responseLines.find(l => l.startsWith('{'));
        const payload = JSON.parse(jsonLine ?? 'null') as {
          name: string;
          usedRangeAddress?: string;
          showGridlines: boolean;
          protected: boolean;
        };
        assert.strictEqual(payload.name, 'Sheet1');
        assert.strictEqual(payload.usedRangeAddress, 'Sheet1!A1:C3');
        assert.strictEqual(payload.showGridlines, true);
        assert.strictEqual(payload.protected, false);
      });
    });
  });

  describe('excel_used_range', () => {
    it('returns used range values', async () => {
      await withMcpContext(async (response, context) => {
        const page = context.getSelectedPptrPage();
        await page.setContent(html`<main>excel fixture</main>`);
        await page.evaluate(() => {
          const g = globalThis as typeof globalThis & {Excel?: unknown};
          g.Excel = {
            run: async (batch: (ctx: unknown) => Promise<unknown>) => {
              const used = {
                load: () => undefined,
                isNullObject: false,
                address: 'Sheet1!A1:B2',
                values: [
                  [1, 2],
                  [3, 4],
                ],
                rowCount: 2,
                columnCount: 2,
              };
              const ws = {
                getUsedRangeOrNullObject: () => used,
              };
              const ctx = {
                workbook: {
                  worksheets: {
                    getActiveWorksheet: () => ws,
                    getItem: () => ws,
                  },
                },
                sync: async () => undefined,
              };
              return batch(ctx);
            },
          };
        });

        await excelUsedRange.handler(
          {params: {}, page: context.getSelectedMcpPage()},
          response,
          context,
        );

        const jsonLine = response.responseLines.find(l => l.startsWith('{'));
        const payload = JSON.parse(jsonLine ?? 'null') as {
          address: string;
          values: unknown[][];
          truncated: boolean;
        };
        assert.strictEqual(payload.address, 'Sheet1!A1:B2');
        assert.deepStrictEqual(payload.values, [
          [1, 2],
          [3, 4],
        ]);
        assert.strictEqual(payload.truncated, false);
      });
    });

    it('reports empty when the worksheet has no used range', async () => {
      await withMcpContext(async (response, context) => {
        const page = context.getSelectedPptrPage();
        await page.setContent(html`<main>excel fixture</main>`);
        await page.evaluate(() => {
          const g = globalThis as typeof globalThis & {Excel?: unknown};
          g.Excel = {
            run: async (batch: (ctx: unknown) => Promise<unknown>) => {
              const used = {load: () => undefined, isNullObject: true};
              const ws = {getUsedRangeOrNullObject: () => used};
              const ctx = {
                workbook: {
                  worksheets: {
                    getActiveWorksheet: () => ws,
                    getItem: () => ws,
                  },
                },
                sync: async () => undefined,
              };
              return batch(ctx);
            },
          };
        });

        await excelUsedRange.handler(
          {params: {}, page: context.getSelectedMcpPage()},
          response,
          context,
        );

        assert.ok(
          response.responseLines.some(l => l.includes('no used range')),
          `expected empty notice, got ${JSON.stringify(response.responseLines)}`,
        );
      });
    });
  });

  describe('excel_read_range', () => {
    it("parses 'Sheet!A1:B2' and returns values", async () => {
      await withMcpContext(async (response, context) => {
        const page = context.getSelectedPptrPage();
        await page.setContent(html`<main>excel fixture</main>`);
        await page.evaluate(() => {
          const g = globalThis as typeof globalThis & {
            Excel?: unknown;
            __excelLastArgs?: unknown;
          };
          g.Excel = {
            run: async (batch: (ctx: unknown) => Promise<unknown>) => {
              const range = {
                load: () => undefined,
                address: 'Sheet2!A1:B2',
                values: [
                  ['a', 'b'],
                  ['c', 'd'],
                ],
                rowCount: 2,
                columnCount: 2,
              };
              const ws = {
                getRange: (a1: string) => {
                  g.__excelLastArgs = {
                    ...((g.__excelLastArgs as object | undefined) ?? {}),
                    a1,
                  };
                  return range;
                },
              };
              const ctx = {
                workbook: {
                  worksheets: {
                    getItem: (name: string) => {
                      g.__excelLastArgs = {
                        ...(g.__excelLastArgs as object),
                        sheet: name,
                      };
                      return ws;
                    },
                    getActiveWorksheet: () => ws,
                  },
                  getSelectedRange: () => range,
                },
                sync: async () => undefined,
              };
              return batch(ctx);
            },
          };
        });

        await excelReadRange.handler(
          {
            params: {address: 'Sheet2!A1:B2'},
            page: context.getSelectedMcpPage(),
          },
          response,
          context,
        );

        const jsonLine = response.responseLines.find(l => l.startsWith('{'));
        const payload = JSON.parse(jsonLine ?? 'null') as {
          address: string;
          values: unknown[][];
        };
        assert.strictEqual(payload.address, 'Sheet2!A1:B2');
        assert.deepStrictEqual(payload.values, [
          ['a', 'b'],
          ['c', 'd'],
        ]);

        const lastArgs = await page.evaluate(
          () =>
            (globalThis as typeof globalThis & {__excelLastArgs?: unknown})
              .__excelLastArgs,
        );
        assert.deepStrictEqual(lastArgs, {sheet: 'Sheet2', a1: 'A1:B2'});
      });
    });
  });

  describe('excel_list_tables', () => {
    it('lists tables with resolved range addresses', async () => {
      await withMcpContext(async (response, context) => {
        const page = context.getSelectedPptrPage();
        await page.setContent(html`<main>excel fixture</main>`);
        await page.evaluate(() => {
          const g = globalThis as typeof globalThis & {Excel?: unknown};
          g.Excel = {
            run: async (batch: (ctx: unknown) => Promise<unknown>) => {
              const tables = {
                load: () => undefined,
                items: [
                  {
                    name: 'Table1',
                    id: 't1',
                    showHeaders: true,
                    showTotals: false,
                    style: 'TableStyleMedium2',
                    worksheet: {name: 'Sheet1'},
                    getRange: () => ({
                      load: function () {
                        return this;
                      },
                      address: 'Sheet1!A1:C4',
                    }),
                    getDataBodyRange: () => ({
                      load: function () {
                        return this;
                      },
                      rowCount: 3,
                    }),
                  },
                ],
              };
              const ctx = {
                workbook: {tables},
                sync: async () => undefined,
              };
              return batch(ctx);
            },
          };
        });

        await excelListTables.handler(
          {params: {}, page: context.getSelectedMcpPage()},
          response,
          context,
        );

        const jsonLine = response.responseLines.find(l => l.startsWith('{'));
        const payload = JSON.parse(jsonLine ?? 'null') as {
          tables: Array<{
            name: string;
            worksheet: string;
            address: string;
            rowCount: number;
          }>;
        };
        assert.strictEqual(payload.tables.length, 1);
        assert.strictEqual(payload.tables[0]?.name, 'Table1');
        assert.strictEqual(payload.tables[0]?.worksheet, 'Sheet1');
        assert.strictEqual(payload.tables[0]?.address, 'Sheet1!A1:C4');
        assert.strictEqual(payload.tables[0]?.rowCount, 3);
      });
    });
  });

  describe('excel_list_named_items', () => {
    it('lists workbook-scoped names', async () => {
      await withMcpContext(async (response, context) => {
        const page = context.getSelectedPptrPage();
        await page.setContent(html`<main>excel fixture</main>`);
        await page.evaluate(() => {
          const g = globalThis as typeof globalThis & {Excel?: unknown};
          g.Excel = {
            run: async (batch: (ctx: unknown) => Promise<unknown>) => {
              const names = {
                load: () => undefined,
                items: [
                  {
                    name: 'TaxRate',
                    type: 'Double',
                    value: 0.08,
                    formula: '=0.08',
                    visible: true,
                    comment: 'VAT',
                  },
                ],
              };
              const ctx = {
                workbook: {names},
                sync: async () => undefined,
              };
              return batch(ctx);
            },
          };
        });

        await excelListNamedItems.handler(
          {params: {}, page: context.getSelectedMcpPage()},
          response,
          context,
        );

        const jsonLine = response.responseLines.find(l => l.startsWith('{'));
        const payload = JSON.parse(jsonLine ?? 'null') as {
          names: Array<{name: string; value: unknown; comment: string}>;
        };
        assert.strictEqual(payload.names.length, 1);
        assert.strictEqual(payload.names[0]?.name, 'TaxRate');
        assert.strictEqual(payload.names[0]?.value, 0.08);
        assert.strictEqual(payload.names[0]?.comment, 'VAT');
      });
    });
  });
});
