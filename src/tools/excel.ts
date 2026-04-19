/* eslint-disable @typescript-eslint/no-explicit-any -- Office.js/Excel.js surface is intentionally loose in these read tools. */
import {zod} from '../third_party/index.js';

import {ToolCategory} from './categories.js';
import {definePageTool} from './ToolDefinition.js';

const MAX_CELLS = 1000;

const rangeTargetSchema = {
  address: zod
    .string()
    .optional()
    .describe(
      "A1 reference such as 'Sheet1!A1:C10' or 'A1:C10'. If omitted, the active selection is used.",
    ),
  sheet: zod
    .string()
    .optional()
    .describe(
      'Worksheet name. Used when address omits the sheet prefix; ignored if address includes one.',
    ),
};

const REQUIREMENT_SET_PROBES = [
  ['ExcelApi', '1.1'],
  ['ExcelApi', '1.7'],
  ['ExcelApi', '1.10'],
  ['ExcelApi', '1.12'],
  ['ExcelApi', '1.14'],
  ['ExcelApiOnline', '1.1'],
  ['SharedRuntime', '1.1'],
  ['DialogApi', '1.1'],
  ['DialogApi', '1.2'],
  ['RibbonApi', '1.1'],
  ['IdentityAPI', '1.3'],
] as const;

export const excelContextInfo = definePageTool({
  name: 'excel_context_info',
  description:
    'Returns Office.js and Excel host information for the selected page, including supported requirement sets when available.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {},
  handler: async (request, response) => {
    const result = await request.page.pptrPage.evaluate(async probes => {
      const globalObject = globalThis as typeof globalThis & {
        Office?: {
          onReady?: () => Promise<unknown>;
          context?: {
            diagnostics?: {
              host?: string;
              platform?: string;
              version?: string;
            };
            contentLanguage?: string;
            displayLanguage?: string;
            requirements?: {
              isSetSupported?: (name: string, version?: string) => boolean;
            };
          };
        };
        Excel?: unknown;
      };

      const hasOfficeGlobal = typeof globalObject.Office !== 'undefined';
      const hasExcelGlobal = typeof globalObject.Excel !== 'undefined';

      if (!hasOfficeGlobal) {
        return {
          hasOfficeGlobal,
          hasExcelGlobal,
          hostInfo: undefined,
          contentLanguage: undefined,
          displayLanguage: undefined,
          requirementSets: [],
        };
      }

      try {
        if (typeof globalObject.Office?.onReady === 'function') {
          await Promise.race([
            globalObject.Office.onReady().catch(() => undefined),
            new Promise(resolve => globalThis.setTimeout(resolve, 1000)),
          ]);
        }
      } catch {
        // Treat readiness failures as "best effort" and continue probing.
      }

      const officeContext = globalObject.Office?.context;
      const diagnostics = officeContext?.diagnostics;
      const requirementSets: string[] = [];

      for (const [name, version] of probes) {
        try {
          if (officeContext?.requirements?.isSetSupported?.(name, version)) {
            requirementSets.push(`${name} ${version}`);
          }
        } catch {
          // Unknown or unsupported requirement-set probes should not fail the tool.
        }
      }

      return {
        hasOfficeGlobal,
        hasExcelGlobal,
        hostInfo: diagnostics
          ? {
              host: diagnostics.host,
              platform: diagnostics.platform,
              version: diagnostics.version,
            }
          : undefined,
        contentLanguage: officeContext?.contentLanguage,
        displayLanguage: officeContext?.displayLanguage,
        requirementSets,
      };
    }, REQUIREMENT_SET_PROBES);

    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelActiveRange = definePageTool({
  name: 'excel_active_range',
  description:
    'Returns the currently selected Excel range (address, dimensions, and values). Optionally includes formulas and number formats. Requires an Excel add-in target with Excel.run available.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    includeFormulas: zod
      .boolean()
      .optional()
      .describe('If true, also return the A1-style formulas for each cell.'),
    includeNumberFormat: zod
      .boolean()
      .optional()
      .describe('If true, also return the Excel number-format code per cell.'),
  },
  handler: async (request, response) => {
    const includeFormulas = request.params.includeFormulas ?? false;
    const includeNumberFormat = request.params.includeNumberFormat ?? false;

    const result = await request.page.pptrPage.evaluate(
      async args => {
        const globalObject = globalThis as typeof globalThis & {
          Excel?: {
            run: <T>(
              batch: (ctx: {
                workbook: {
                  getSelectedRange: () => {
                    load: (props: string | string[]) => void;
                    address: string;
                    values: unknown[][];
                    formulas: string[][];
                    numberFormat: string[][];
                    rowCount: number;
                    columnCount: number;
                  };
                };
                sync: () => Promise<void>;
              }) => Promise<T>,
            ) => Promise<T>;
          };
        };

        if (typeof globalObject.Excel === 'undefined') {
          return {error: 'Excel API not available on this target'} as const;
        }

        try {
          return await globalObject.Excel.run(async ctx => {
            const range = ctx.workbook.getSelectedRange();
            range.load(['address', 'values', 'rowCount', 'columnCount']);
            if (args.includeFormulas) {
              range.load('formulas');
            }
            if (args.includeNumberFormat) {
              range.load('numberFormat');
            }
            await ctx.sync();

            const totalCells = range.rowCount * range.columnCount;
            const truncated = totalCells > args.maxCells;
            const values = truncated
              ? range.values.slice(0, 1).map(row => row.slice(0, 1))
              : range.values;
            const formulas = args.includeFormulas
              ? truncated
                ? range.formulas.slice(0, 1).map(row => row.slice(0, 1))
                : range.formulas
              : undefined;
            const numberFormat = args.includeNumberFormat
              ? truncated
                ? range.numberFormat.slice(0, 1).map(row => row.slice(0, 1))
                : range.numberFormat
              : undefined;

            return {
              address: range.address,
              rowCount: range.rowCount,
              columnCount: range.columnCount,
              values,
              formulas,
              numberFormat,
              truncated,
            };
          });
        } catch (error) {
          return {
            error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
          } as const;
        }
      },
      {includeFormulas, includeNumberFormat, maxCells: MAX_CELLS},
    );

    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }

    if (result.truncated) {
      response.appendResponseLine(
        `Range ${result.address} has ${result.rowCount}x${result.columnCount} cells (> ${MAX_CELLS}); values truncated to the first cell.`,
      );
    }

    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelWorkbookInfo = definePageTool({
  name: 'excel_workbook_info',
  description:
    'Returns workbook-level metadata: name, save state, calculation mode and state, and protection state.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {},
  handler: async (request, response) => {
    const result = await request.page.pptrPage.evaluate(async () => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const wb = ctx.workbook;
          wb.load(['name', 'isDirty', 'readOnly']);
          const app = ctx.workbook.application;
          app.load(['calculationMode', 'calculationState']);
          const protection = wb.protection;
          protection.load('protected');
          await ctx.sync();
          return {
            name: wb.name,
            isDirty: wb.isDirty,
            readOnly: wb.readOnly,
            protected: protection.protected,
            calculationMode: app.calculationMode,
            calculationState: app.calculationState,
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    });
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelListWorksheets = definePageTool({
  name: 'excel_list_worksheets',
  description:
    'Lists all worksheets in the workbook with name, id, position, visibility, and tab color.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {},
  handler: async (request, response) => {
    const result = await request.page.pptrPage.evaluate(async () => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const sheets = ctx.workbook.worksheets;
          sheets.load(
            'items/name,items/id,items/position,items/visibility,items/tabColor',
          );
          const active = ctx.workbook.worksheets.getActiveWorksheet();
          active.load('id');
          await ctx.sync();
          const activeId = active.id;
          return {
            worksheets: sheets.items.map((s: any) => ({
              name: s.name,
              id: s.id,
              position: s.position,
              visibility: s.visibility,
              tabColor: s.tabColor,
              active: s.id === activeId,
            })),
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    });
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelWorksheetInfo = definePageTool({
  name: 'excel_worksheet_info',
  description:
    'Returns metadata for a single worksheet: used range address, visibility, protection, gridlines, tab color, and dimensions.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    sheet: zod
      .string()
      .optional()
      .describe('Worksheet name. Omit to use the active worksheet.'),
  },
  handler: async (request, response) => {
    const sheet = request.params.sheet;
    const result = await request.page.pptrPage.evaluate(
      async (args: {sheet?: string}) => {
        const g = globalThis as typeof globalThis & {Excel?: any};
        if (typeof g.Excel === 'undefined') {
          return {error: 'Excel API not available on this target'} as const;
        }
        try {
          return await g.Excel.run(async (ctx: any) => {
            const ws = args.sheet
              ? ctx.workbook.worksheets.getItem(args.sheet)
              : ctx.workbook.worksheets.getActiveWorksheet();
            ws.load([
              'name',
              'id',
              'position',
              'visibility',
              'tabColor',
              'showGridlines',
              'showHeadings',
              'standardHeight',
              'standardWidth',
            ]);
            const used = ws.getUsedRangeOrNullObject(true);
            used.load('address');
            const protection = ws.protection;
            protection.load('protected');
            await ctx.sync();
            return {
              name: ws.name,
              id: ws.id,
              position: ws.position,
              visibility: ws.visibility,
              tabColor: ws.tabColor,
              showGridlines: ws.showGridlines,
              showHeadings: ws.showHeadings,
              standardHeight: ws.standardHeight,
              standardWidth: ws.standardWidth,
              protected: protection.protected,
              usedRangeAddress: used.isNullObject ? undefined : used.address,
            };
          });
        } catch (error) {
          return {
            error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
          } as const;
        }
      },
      {sheet},
    );
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelUsedRange = definePageTool({
  name: 'excel_used_range',
  description:
    'Returns values (and optionally formulas / number formats) for a worksheet’s used range, with truncation when the range exceeds the cell cap.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    sheet: zod
      .string()
      .optional()
      .describe('Worksheet name. Omit to use the active worksheet.'),
    valuesOnly: zod
      .boolean()
      .optional()
      .describe(
        'If true (default), only cells with values count toward the used range.',
      ),
    includeFormulas: zod
      .boolean()
      .optional()
      .describe('If true, also return the A1-style formulas for each cell.'),
    includeNumberFormat: zod
      .boolean()
      .optional()
      .describe('If true, also return the Excel number-format code per cell.'),
  },
  handler: async (request, response) => {
    const args = {
      sheet: request.params.sheet,
      valuesOnly: request.params.valuesOnly ?? true,
      includeFormulas: request.params.includeFormulas ?? false,
      includeNumberFormat: request.params.includeNumberFormat ?? false,
      maxCells: MAX_CELLS,
    };
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const ws = args.sheet
            ? ctx.workbook.worksheets.getItem(args.sheet)
            : ctx.workbook.worksheets.getActiveWorksheet();
          const used = ws.getUsedRangeOrNullObject(args.valuesOnly);
          const props = ['address', 'values', 'rowCount', 'columnCount'];
          if (args.includeFormulas) {
            props.push('formulas');
          }
          if (args.includeNumberFormat) {
            props.push('numberFormat');
          }
          used.load(props);
          await ctx.sync();
          if (used.isNullObject) {
            return {empty: true} as const;
          }
          const total = used.rowCount * used.columnCount;
          const truncated = total > args.maxCells;
          const slice = <T>(v: T[][]): T[][] =>
            truncated ? v.slice(0, 1).map(r => r.slice(0, 1)) : v;
          return {
            address: used.address,
            rowCount: used.rowCount,
            columnCount: used.columnCount,
            values: slice(used.values),
            formulas: args.includeFormulas ? slice(used.formulas) : undefined,
            numberFormat: args.includeNumberFormat
              ? slice(used.numberFormat)
              : undefined,
            truncated,
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    if ('empty' in result) {
      response.appendResponseLine('Worksheet has no used range.');
      return;
    }
    if (result.truncated) {
      response.appendResponseLine(
        `Range ${result.address} has ${result.rowCount}x${result.columnCount} cells (> ${MAX_CELLS}); values truncated to the first cell.`,
      );
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelReadRange = definePageTool({
  name: 'excel_read_range',
  description:
    "Reads a range by address (e.g. 'Sheet1!A1:C10' or 'A1:C10' with a sheet param). Omit address to read the active selection. Returns values and optionally formulas / number formats.",
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    ...rangeTargetSchema,
    includeFormulas: zod
      .boolean()
      .optional()
      .describe('If true, also return the A1-style formulas for each cell.'),
    includeNumberFormat: zod
      .boolean()
      .optional()
      .describe('If true, also return the Excel number-format code per cell.'),
  },
  handler: async (request, response) => {
    const args = {
      address: request.params.address,
      sheet: request.params.sheet,
      includeFormulas: request.params.includeFormulas ?? false,
      includeNumberFormat: request.params.includeNumberFormat ?? false,
      maxCells: MAX_CELLS,
    };
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          let range: any;
          if (!args.address) {
            range = ctx.workbook.getSelectedRange();
          } else {
            const bangIdx = args.address.indexOf('!');
            let sheetName = args.sheet;
            let a1 = args.address;
            if (bangIdx >= 0) {
              sheetName = args.address.slice(0, bangIdx).replace(/^'|'$/g, '');
              a1 = args.address.slice(bangIdx + 1);
            }
            const ws = sheetName
              ? ctx.workbook.worksheets.getItem(sheetName)
              : ctx.workbook.worksheets.getActiveWorksheet();
            range = ws.getRange(a1);
          }
          const props = ['address', 'values', 'rowCount', 'columnCount'];
          if (args.includeFormulas) {
            props.push('formulas');
          }
          if (args.includeNumberFormat) {
            props.push('numberFormat');
          }
          range.load(props);
          await ctx.sync();
          const total = range.rowCount * range.columnCount;
          const truncated = total > args.maxCells;
          const slice = <T>(v: T[][]): T[][] =>
            truncated ? v.slice(0, 1).map(r => r.slice(0, 1)) : v;
          return {
            address: range.address,
            rowCount: range.rowCount,
            columnCount: range.columnCount,
            values: slice(range.values),
            formulas: args.includeFormulas ? slice(range.formulas) : undefined,
            numberFormat: args.includeNumberFormat
              ? slice(range.numberFormat)
              : undefined,
            truncated,
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    if (result.truncated) {
      response.appendResponseLine(
        `Range ${result.address} has ${result.rowCount}x${result.columnCount} cells (> ${MAX_CELLS}); values truncated to the first cell.`,
      );
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelRangeProperties = definePageTool({
  name: 'excel_range_properties',
  description:
    'Returns rich properties for a range: value types, hasSpill, row/column hidden flags, and selected format details (font, fill, alignment, borders). Use include flags to bound payload.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    ...rangeTargetSchema,
    includeFormat: zod
      .boolean()
      .optional()
      .describe(
        'If true, include font, fill, alignment, and border summary per cell.',
      ),
    includeStyle: zod
      .boolean()
      .optional()
      .describe('If true, include the named style of each cell.'),
  },
  handler: async (request, response) => {
    const args = {
      address: request.params.address,
      sheet: request.params.sheet,
      includeFormat: request.params.includeFormat ?? false,
      includeStyle: request.params.includeStyle ?? false,
      maxCells: MAX_CELLS,
    };
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          let range: any;
          if (!args.address) {
            range = ctx.workbook.getSelectedRange();
          } else {
            const bangIdx = args.address.indexOf('!');
            let sheetName = args.sheet;
            let a1 = args.address;
            if (bangIdx >= 0) {
              sheetName = args.address.slice(0, bangIdx).replace(/^'|'$/g, '');
              a1 = args.address.slice(bangIdx + 1);
            }
            const ws = sheetName
              ? ctx.workbook.worksheets.getItem(sheetName)
              : ctx.workbook.worksheets.getActiveWorksheet();
            range = ws.getRange(a1);
          }
          const props = [
            'address',
            'rowCount',
            'columnCount',
            'valueTypes',
            'hasSpill',
            'rowHidden',
            'columnHidden',
          ];
          if (args.includeStyle) {
            props.push('style');
          }
          range.load(props);
          if (args.includeFormat) {
            range.format.load([
              'horizontalAlignment',
              'verticalAlignment',
              'wrapText',
            ]);
            range.format.font.load(['name', 'size', 'bold', 'italic', 'color']);
            range.format.fill.load('color');
          }
          await ctx.sync();
          const total = range.rowCount * range.columnCount;
          const truncated = total > args.maxCells;
          const slice = <T>(v: T[][]): T[][] =>
            truncated ? v.slice(0, 1).map(r => r.slice(0, 1)) : v;
          return {
            address: range.address,
            rowCount: range.rowCount,
            columnCount: range.columnCount,
            valueTypes: slice(range.valueTypes),
            hasSpill: range.hasSpill,
            rowHidden: range.rowHidden,
            columnHidden: range.columnHidden,
            style: args.includeStyle ? range.style : undefined,
            format: args.includeFormat
              ? {
                  horizontalAlignment: range.format.horizontalAlignment,
                  verticalAlignment: range.format.verticalAlignment,
                  wrapText: range.format.wrapText,
                  font: {
                    name: range.format.font.name,
                    size: range.format.font.size,
                    bold: range.format.font.bold,
                    italic: range.format.font.italic,
                    color: range.format.font.color,
                  },
                  fill: {color: range.format.fill.color},
                }
              : undefined,
            truncated,
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    if (result.truncated) {
      response.appendResponseLine(
        `Range ${result.address} has ${result.rowCount}x${result.columnCount} cells (> ${MAX_CELLS}); grid payloads truncated to the first cell.`,
      );
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelRangeFormulas = definePageTool({
  name: 'excel_range_formulas',
  description:
    'Returns formulas (A1 and R1C1) alongside resolved values for a range. Useful for verifying formula edits.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    ...rangeTargetSchema,
  },
  handler: async (request, response) => {
    const args = {
      address: request.params.address,
      sheet: request.params.sheet,
      maxCells: MAX_CELLS,
    };
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          let range: any;
          if (!args.address) {
            range = ctx.workbook.getSelectedRange();
          } else {
            const bangIdx = args.address.indexOf('!');
            let sheetName = args.sheet;
            let a1 = args.address;
            if (bangIdx >= 0) {
              sheetName = args.address.slice(0, bangIdx).replace(/^'|'$/g, '');
              a1 = args.address.slice(bangIdx + 1);
            }
            const ws = sheetName
              ? ctx.workbook.worksheets.getItem(sheetName)
              : ctx.workbook.worksheets.getActiveWorksheet();
            range = ws.getRange(a1);
          }
          range.load([
            'address',
            'rowCount',
            'columnCount',
            'values',
            'formulas',
            'formulasR1C1',
          ]);
          await ctx.sync();
          const total = range.rowCount * range.columnCount;
          const truncated = total > args.maxCells;
          const slice = <T>(v: T[][]): T[][] =>
            truncated ? v.slice(0, 1).map(r => r.slice(0, 1)) : v;
          return {
            address: range.address,
            rowCount: range.rowCount,
            columnCount: range.columnCount,
            values: slice(range.values),
            formulas: slice(range.formulas),
            formulasR1C1: slice(range.formulasR1C1),
            truncated,
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    if (result.truncated) {
      response.appendResponseLine(
        `Range ${result.address} has ${result.rowCount}x${result.columnCount} cells (> ${MAX_CELLS}); values truncated to the first cell.`,
      );
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelRangeSpecialCells = definePageTool({
  name: 'excel_range_special_cells',
  description:
    "Finds cells within a range matching a category: 'constants', 'formulas', 'blanks', or 'visible'. Optionally filter by value type. Returns the resulting address and cell count.",
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    ...rangeTargetSchema,
    cellType: zod
      .enum(['constants', 'formulas', 'blanks', 'visible'])
      .describe('Category of special cells to locate.'),
    valueType: zod
      .enum(['all', 'errors', 'logical', 'numbers', 'text'])
      .optional()
      .describe(
        "For 'constants' or 'formulas', filter by value type. Defaults to 'all'.",
      ),
  },
  handler: async (request, response) => {
    const args = {
      address: request.params.address,
      sheet: request.params.sheet,
      cellType: request.params.cellType,
      valueType: request.params.valueType ?? 'all',
    };
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          let range: any;
          if (!args.address) {
            range = ctx.workbook.getSelectedRange();
          } else {
            const bangIdx = args.address.indexOf('!');
            let sheetName = args.sheet;
            let a1 = args.address;
            if (bangIdx >= 0) {
              sheetName = args.address.slice(0, bangIdx).replace(/^'|'$/g, '');
              a1 = args.address.slice(bangIdx + 1);
            }
            const ws = sheetName
              ? ctx.workbook.worksheets.getItem(sheetName)
              : ctx.workbook.worksheets.getActiveWorksheet();
            range = ws.getRange(a1);
          }
          const special =
            args.cellType === 'blanks' || args.cellType === 'visible'
              ? range.getSpecialCellsOrNullObject(args.cellType)
              : range.getSpecialCellsOrNullObject(
                  args.cellType,
                  args.valueType,
                );
          special.load(['address', 'cellCount', 'isNullObject']);
          await ctx.sync();
          if (special.isNullObject) {
            return {found: false} as const;
          }
          return {
            found: true,
            address: special.address,
            cellCount: special.cellCount,
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelFindInRange = definePageTool({
  name: 'excel_find_in_range',
  description:
    'Finds all matches of a text string within a range. Returns the combined match address and cell count.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    ...rangeTargetSchema,
    text: zod.string().describe('Text to search for.'),
    completeMatch: zod
      .boolean()
      .optional()
      .describe('If true, require a whole-cell match. Defaults to false.'),
    matchCase: zod
      .boolean()
      .optional()
      .describe('If true, the search is case-sensitive. Defaults to false.'),
  },
  handler: async (request, response) => {
    const args = {
      address: request.params.address,
      sheet: request.params.sheet,
      text: request.params.text,
      completeMatch: request.params.completeMatch ?? false,
      matchCase: request.params.matchCase ?? false,
    };
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          let range: any;
          if (!args.address) {
            range = ctx.workbook.getSelectedRange();
          } else {
            const bangIdx = args.address.indexOf('!');
            let sheetName = args.sheet;
            let a1 = args.address;
            if (bangIdx >= 0) {
              sheetName = args.address.slice(0, bangIdx).replace(/^'|'$/g, '');
              a1 = args.address.slice(bangIdx + 1);
            }
            const ws = sheetName
              ? ctx.workbook.worksheets.getItem(sheetName)
              : ctx.workbook.worksheets.getActiveWorksheet();
            range = ws.getRange(a1);
          }
          const found = range.findAllOrNullObject(args.text, {
            completeMatch: args.completeMatch,
            matchCase: args.matchCase,
          });
          found.load(['address', 'cellCount', 'isNullObject']);
          await ctx.sync();
          if (found.isNullObject) {
            return {found: false} as const;
          }
          return {
            found: true,
            address: found.address,
            cellCount: found.cellCount,
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelListConditionalFormats = definePageTool({
  name: 'excel_list_conditional_formats',
  description:
    'Lists conditional-format rules on a range: id, type, priority, stopIfTrue. Omit address to use the active worksheet’s used range.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    ...rangeTargetSchema,
  },
  handler: async (request, response) => {
    const args = {
      address: request.params.address,
      sheet: request.params.sheet,
    };
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          let range: any;
          if (!args.address) {
            const ws = args.sheet
              ? ctx.workbook.worksheets.getItem(args.sheet)
              : ctx.workbook.worksheets.getActiveWorksheet();
            range = ws.getUsedRangeOrNullObject(true);
          } else {
            const bangIdx = args.address.indexOf('!');
            let sheetName = args.sheet;
            let a1 = args.address;
            if (bangIdx >= 0) {
              sheetName = args.address.slice(0, bangIdx).replace(/^'|'$/g, '');
              a1 = args.address.slice(bangIdx + 1);
            }
            const ws = sheetName
              ? ctx.workbook.worksheets.getItem(sheetName)
              : ctx.workbook.worksheets.getActiveWorksheet();
            range = ws.getRange(a1);
          }
          const cfs = range.conditionalFormats;
          cfs.load('items/id,items/type,items/priority,items/stopIfTrue');
          range.load(['address', 'isNullObject']);
          await ctx.sync();
          if (range.isNullObject) {
            return {empty: true} as const;
          }
          return {
            address: range.address,
            conditionalFormats: cfs.items.map((c: any) => ({
              id: c.id,
              type: c.type,
              priority: c.priority,
              stopIfTrue: c.stopIfTrue,
            })),
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    if ('empty' in result) {
      response.appendResponseLine('Worksheet has no used range.');
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelListDataValidations = definePageTool({
  name: 'excel_list_data_validations',
  description:
    'Returns data-validation configuration on a range: type, rule, error alert, and prompt. Omit address to use the active selection.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    ...rangeTargetSchema,
  },
  handler: async (request, response) => {
    const args = {
      address: request.params.address,
      sheet: request.params.sheet,
    };
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          let range: any;
          if (!args.address) {
            range = ctx.workbook.getSelectedRange();
          } else {
            const bangIdx = args.address.indexOf('!');
            let sheetName = args.sheet;
            let a1 = args.address;
            if (bangIdx >= 0) {
              sheetName = args.address.slice(0, bangIdx).replace(/^'|'$/g, '');
              a1 = args.address.slice(bangIdx + 1);
            }
            const ws = sheetName
              ? ctx.workbook.worksheets.getItem(sheetName)
              : ctx.workbook.worksheets.getActiveWorksheet();
            range = ws.getRange(a1);
          }
          range.load('address');
          const dv = range.dataValidation;
          dv.load([
            'type',
            'rule',
            'errorAlert',
            'prompt',
            'ignoreBlanks',
            'valid',
          ]);
          await ctx.sync();
          return {
            address: range.address,
            dataValidation: {
              type: dv.type,
              rule: dv.rule,
              errorAlert: dv.errorAlert,
              prompt: dv.prompt,
              ignoreBlanks: dv.ignoreBlanks,
              valid: dv.valid,
            },
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelListTables = definePageTool({
  name: 'excel_list_tables',
  description:
    'Lists all tables (ListObjects) in the workbook with name, worksheet, address, header/total row flags, row count, and style.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {},
  handler: async (request, response) => {
    const result = await request.page.pptrPage.evaluate(async () => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const tables = ctx.workbook.tables;
          tables.load(
            'items/name,items/id,items/showHeaders,items/showTotals,items/style,items/worksheet/name',
          );
          await ctx.sync();
          const ranges = tables.items.map((t: any) =>
            t.getRange().load('address'),
          );
          const rowRanges = tables.items.map((t: any) =>
            t.getDataBodyRange().load('rowCount'),
          );
          await ctx.sync();
          return {
            tables: tables.items.map((t: any, i: number) => ({
              name: t.name,
              id: t.id,
              worksheet: t.worksheet.name,
              address: ranges[i].address,
              rowCount: rowRanges[i].rowCount,
              showHeaders: t.showHeaders,
              showTotals: t.showTotals,
              style: t.style,
            })),
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    });
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelListNamedItems = definePageTool({
  name: 'excel_list_named_items',
  description:
    'Lists workbook-scoped named items (named ranges and formulas) with name, type, value, formula, visibility, and comment.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {},
  handler: async (request, response) => {
    const result = await request.page.pptrPage.evaluate(async () => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const names = ctx.workbook.names;
          names.load(
            'items/name,items/type,items/value,items/visible,items/comment,items/formula',
          );
          await ctx.sync();
          return {
            names: names.items.map((n: any) => ({
              name: n.name,
              type: n.type,
              value: n.value,
              formula: n.formula,
              visible: n.visible,
              comment: n.comment,
            })),
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    });
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelTableInfo = definePageTool({
  name: 'excel_table_info',
  description:
    'Returns detail for a single table: name, worksheet, address, row count, columns (name + filter criteria), header/total row flags, and style.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    name: zod.string().describe('Table name (ListObject name).'),
  },
  handler: async (request, response) => {
    const args = {name: request.params.name};
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const table = ctx.workbook.tables.getItem(args.name);
          table.load([
            'name',
            'id',
            'showHeaders',
            'showTotals',
            'style',
            'worksheet/name',
          ]);
          const range = table.getRange().load('address');
          const body = table.getDataBodyRange().load('rowCount,columnCount');
          const cols = table.columns;
          cols.load('items/name,items/id,items/index');
          await ctx.sync();
          const columnFilters = cols.items.map((c: any) => {
            const f = c.filter;
            f.load('criteria');
            return f;
          });
          await ctx.sync();
          return {
            name: table.name,
            id: table.id,
            worksheet: table.worksheet.name,
            address: range.address,
            rowCount: body.rowCount,
            columnCount: body.columnCount,
            showHeaders: table.showHeaders,
            showTotals: table.showTotals,
            style: table.style,
            columns: cols.items.map((c: any, i: number) => ({
              name: c.name,
              id: c.id,
              index: c.index,
              filterCriteria: columnFilters[i].criteria,
            })),
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelTableRows = definePageTool({
  name: 'excel_table_rows',
  description:
    'Returns the data-body values of a table with truncation when row*column count exceeds the cell cap.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    name: zod.string().describe('Table name (ListObject name).'),
    includeHeaders: zod
      .boolean()
      .optional()
      .describe('If true, include the header row names.'),
  },
  handler: async (request, response) => {
    const args = {
      name: request.params.name,
      includeHeaders: request.params.includeHeaders ?? false,
      maxCells: MAX_CELLS,
    };
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const table = ctx.workbook.tables.getItem(args.name);
          const body = table.getDataBodyRange();
          body.load(['address', 'values', 'rowCount', 'columnCount']);
          let headers: any;
          if (args.includeHeaders) {
            headers = table.getHeaderRowRange().load('values');
          }
          await ctx.sync();
          const total = body.rowCount * body.columnCount;
          const truncated = total > args.maxCells;
          const values = truncated
            ? body.values.slice(0, 1).map((r: any[]) => r.slice(0, 1))
            : body.values;
          return {
            address: body.address,
            rowCount: body.rowCount,
            columnCount: body.columnCount,
            headers: args.includeHeaders ? headers.values?.[0] : undefined,
            values,
            truncated,
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    if (result.truncated) {
      response.appendResponseLine(
        `Table body has ${result.rowCount}x${result.columnCount} cells (> ${MAX_CELLS}); values truncated to the first cell.`,
      );
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelTableFilters = definePageTool({
  name: 'excel_table_filters',
  description:
    'Returns the active filter criteria per column for a table. Columns without an active filter have null criteria.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    name: zod.string().describe('Table name (ListObject name).'),
  },
  handler: async (request, response) => {
    const args = {name: request.params.name};
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const table = ctx.workbook.tables.getItem(args.name);
          const cols = table.columns;
          cols.load('items/name,items/index');
          await ctx.sync();
          const filters = cols.items.map((c: any) => {
            const f = c.filter;
            f.load('criteria');
            return f;
          });
          await ctx.sync();
          return {
            table: args.name,
            columns: cols.items.map((c: any, i: number) => ({
              name: c.name,
              index: c.index,
              criteria: filters[i].criteria,
            })),
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelListComments = definePageTool({
  name: 'excel_list_comments',
  description:
    'Lists comments and replies on a worksheet: author, content, timestamp, and cell address.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    sheet: zod
      .string()
      .optional()
      .describe('Worksheet name. Omit to use the active worksheet.'),
  },
  handler: async (request, response) => {
    const args = {sheet: request.params.sheet};
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const ws = args.sheet
            ? ctx.workbook.worksheets.getItem(args.sheet)
            : ctx.workbook.worksheets.getActiveWorksheet();
          ws.load('name');
          const comments = ws.comments;
          comments.load(
            'items/id,items/authorName,items/authorEmail,items/content,items/creationDate,items/resolved',
          );
          await ctx.sync();
          const cellRanges = comments.items.map((c: any) =>
            c.getLocation().load('address'),
          );
          const replyLists = comments.items.map((c: any) => {
            const r = c.replies;
            r.load(
              'items/id,items/authorName,items/content,items/creationDate',
            );
            return r;
          });
          await ctx.sync();
          return {
            worksheet: ws.name,
            comments: comments.items.map((c: any, i: number) => ({
              id: c.id,
              author: c.authorName,
              authorEmail: c.authorEmail,
              content: c.content,
              creationDate: c.creationDate,
              resolved: c.resolved,
              address: cellRanges[i].address,
              replies: replyLists[i].items.map((r: any) => ({
                id: r.id,
                author: r.authorName,
                content: r.content,
                creationDate: r.creationDate,
              })),
            })),
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelListShapes = definePageTool({
  name: 'excel_list_shapes',
  description:
    'Lists shapes (including images) on a worksheet: name, id, type, position, size, visibility.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    sheet: zod
      .string()
      .optional()
      .describe('Worksheet name. Omit to use the active worksheet.'),
  },
  handler: async (request, response) => {
    const args = {sheet: request.params.sheet};
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const ws = args.sheet
            ? ctx.workbook.worksheets.getItem(args.sheet)
            : ctx.workbook.worksheets.getActiveWorksheet();
          ws.load('name');
          const shapes = ws.shapes;
          shapes.load(
            'items/id,items/name,items/type,items/left,items/top,items/width,items/height,items/visible,items/altTextDescription',
          );
          await ctx.sync();
          return {
            worksheet: ws.name,
            shapes: shapes.items.map((s: any) => ({
              id: s.id,
              name: s.name,
              type: s.type,
              left: s.left,
              top: s.top,
              width: s.width,
              height: s.height,
              visible: s.visible,
              altTextDescription: s.altTextDescription,
            })),
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelCalculationState = definePageTool({
  name: 'excel_calculation_state',
  description:
    'Returns the workbook calculation mode (automatic/manual/etc.) and current calculation state (done/calculating/pending).',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {},
  handler: async (request, response) => {
    const result = await request.page.pptrPage.evaluate(async () => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const app = ctx.workbook.application;
          app.load([
            'calculationMode',
            'calculationState',
            'calculationEngineVersion',
            'iterativeCalculation',
          ]);
          await ctx.sync();
          return {
            calculationMode: app.calculationMode,
            calculationState: app.calculationState,
            calculationEngineVersion: app.calculationEngineVersion,
            iterativeCalculation: app.iterativeCalculation,
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    });
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelListPivotTables = definePageTool({
  name: 'excel_list_pivot_tables',
  description:
    'Lists all PivotTables in the workbook with name, worksheet, layout address, and enabled flags.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {},
  handler: async (request, response) => {
    const result = await request.page.pptrPage.evaluate(async () => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const pivots = ctx.workbook.pivotTables;
          pivots.load(
            'items/id,items/name,items/enableDataValueEditing,items/useCustomSortLists,items/worksheet/name',
          );
          await ctx.sync();
          const ranges = pivots.items.map((p: any) =>
            p.layout.getRange().load('address'),
          );
          await ctx.sync();
          return {
            pivotTables: pivots.items.map((p: any, i: number) => ({
              id: p.id,
              name: p.name,
              worksheet: p.worksheet.name,
              address: ranges[i].address,
              enableDataValueEditing: p.enableDataValueEditing,
              useCustomSortLists: p.useCustomSortLists,
            })),
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    });
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelPivotTableInfo = definePageTool({
  name: 'excel_pivot_table_info',
  description:
    'Returns the structure of a PivotTable: row, column, data, and filter hierarchies with their source field names.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    name: zod.string().describe('PivotTable name.'),
  },
  handler: async (request, response) => {
    const args = {name: request.params.name};
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const pt = ctx.workbook.pivotTables.getItem(args.name);
          pt.load(['id', 'name', 'worksheet/name']);
          const rows = pt.rowHierarchies;
          const cols = pt.columnHierarchies;
          const data = pt.dataHierarchies;
          const filters = pt.filterHierarchies;
          rows.load('items/id,items/name');
          cols.load('items/id,items/name');
          data.load(
            'items/id,items/name,items/summarizeBy,items/showAs,items/numberFormat',
          );
          filters.load('items/id,items/name');
          const layoutRange = pt.layout.getRange().load('address');
          await ctx.sync();
          return {
            id: pt.id,
            name: pt.name,
            worksheet: pt.worksheet.name,
            address: layoutRange.address,
            rowHierarchies: rows.items.map((h: any) => ({
              id: h.id,
              name: h.name,
            })),
            columnHierarchies: cols.items.map((h: any) => ({
              id: h.id,
              name: h.name,
            })),
            dataHierarchies: data.items.map((h: any) => ({
              id: h.id,
              name: h.name,
              summarizeBy: h.summarizeBy,
              showAs: h.showAs,
              numberFormat: h.numberFormat,
            })),
            filterHierarchies: filters.items.map((h: any) => ({
              id: h.id,
              name: h.name,
            })),
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelPivotTableValues = definePageTool({
  name: 'excel_pivot_table_values',
  description:
    'Returns the rendered values of a PivotTable layout range with truncation when it exceeds the cell cap.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    name: zod.string().describe('PivotTable name.'),
  },
  handler: async (request, response) => {
    const args = {name: request.params.name, maxCells: MAX_CELLS};
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const pt = ctx.workbook.pivotTables.getItem(args.name);
          const range = pt.layout.getRange();
          range.load(['address', 'values', 'rowCount', 'columnCount']);
          await ctx.sync();
          const total = range.rowCount * range.columnCount;
          const truncated = total > args.maxCells;
          const values = truncated
            ? range.values.slice(0, 1).map((r: any[]) => r.slice(0, 1))
            : range.values;
          return {
            address: range.address,
            rowCount: range.rowCount,
            columnCount: range.columnCount,
            values,
            truncated,
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    if (result.truncated) {
      response.appendResponseLine(
        `PivotTable layout has ${result.rowCount}x${result.columnCount} cells (> ${MAX_CELLS}); values truncated to the first cell.`,
      );
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelListCharts = definePageTool({
  name: 'excel_list_charts',
  description:
    'Lists all charts across worksheets: name, id, worksheet, type, title, position, and size.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    sheet: zod
      .string()
      .optional()
      .describe('Worksheet name. Omit to list charts on all worksheets.'),
  },
  handler: async (request, response) => {
    const args = {sheet: request.params.sheet};
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const worksheets: any[] = [];
          if (args.sheet) {
            const ws = ctx.workbook.worksheets.getItem(args.sheet);
            ws.load('name');
            worksheets.push(ws);
          } else {
            const list = ctx.workbook.worksheets;
            list.load('items/name');
            await ctx.sync();
            for (const ws of list.items) {
              worksheets.push(ws);
            }
          }
          const chartLists = worksheets.map(ws => {
            const charts = ws.charts;
            charts.load(
              'items/id,items/name,items/chartType,items/title/text,items/left,items/top,items/width,items/height',
            );
            return {ws, charts};
          });
          await ctx.sync();
          const charts: any[] = [];
          for (const {ws, charts: list} of chartLists) {
            for (const c of list.items) {
              charts.push({
                id: c.id,
                name: c.name,
                worksheet: ws.name,
                chartType: c.chartType,
                title: c.title?.text,
                left: c.left,
                top: c.top,
                width: c.width,
                height: c.height,
              });
            }
          }
          return {charts};
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelChartInfo = definePageTool({
  name: 'excel_chart_info',
  description:
    'Returns detailed information about a chart: type, title, series names, axis titles, and source data address.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    sheet: zod.string().describe('Worksheet name containing the chart.'),
    name: zod.string().describe('Chart name.'),
  },
  handler: async (request, response) => {
    const args = {sheet: request.params.sheet, name: request.params.name};
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const ws = ctx.workbook.worksheets.getItem(args.sheet);
          const chart = ws.charts.getItem(args.name);
          chart.load([
            'id',
            'name',
            'chartType',
            'title/text',
            'left',
            'top',
            'width',
            'height',
            'axes/categoryAxis/title/text',
            'axes/valueAxis/title/text',
          ]);
          const series = chart.series;
          series.load('items/name,items/chartType');
          await ctx.sync();
          return {
            id: chart.id,
            name: chart.name,
            chartType: chart.chartType,
            title: chart.title?.text,
            left: chart.left,
            top: chart.top,
            width: chart.width,
            height: chart.height,
            categoryAxisTitle: chart.axes?.categoryAxis?.title?.text,
            valueAxisTitle: chart.axes?.valueAxis?.title?.text,
            series: series.items.map((s: any) => ({
              name: s.name,
              chartType: s.chartType,
            })),
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelChartImage = definePageTool({
  name: 'excel_chart_image',
  description:
    'Returns a chart rendered as a PNG image, encoded as base64. Useful for visual verification.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    sheet: zod.string().describe('Worksheet name containing the chart.'),
    name: zod.string().describe('Chart name.'),
    width: zod
      .number()
      .optional()
      .describe('Image width in pixels. Defaults to the chart’s natural size.'),
    height: zod
      .number()
      .optional()
      .describe(
        'Image height in pixels. Defaults to the chart’s natural size.',
      ),
  },
  handler: async (request, response) => {
    const args = {
      sheet: request.params.sheet,
      name: request.params.name,
      width: request.params.width,
      height: request.params.height,
    };
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const ws = ctx.workbook.worksheets.getItem(args.sheet);
          const chart = ws.charts.getItem(args.name);
          const image =
            args.width && args.height
              ? chart.getImage(args.width, args.height)
              : chart.getImage();
          await ctx.sync();
          return {base64: image.value as string};
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(
      JSON.stringify(
        {
          sheet: args.sheet,
          name: args.name,
          mimeType: 'image/png',
          base64Length: result.base64.length,
          base64: result.base64,
        },
        null,
        2,
      ),
    );
  },
});

export const excelCustomXmlParts = definePageTool({
  name: 'excel_custom_xml_parts',
  description:
    'Lists custom XML parts stored in the workbook: id and namespace URI.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {},
  handler: async (request, response) => {
    const result = await request.page.pptrPage.evaluate(async () => {
      const g = globalThis as typeof globalThis & {Excel?: any};
      if (typeof g.Excel === 'undefined') {
        return {error: 'Excel API not available on this target'} as const;
      }
      try {
        return await g.Excel.run(async (ctx: any) => {
          const parts = ctx.workbook.customXmlParts;
          parts.load('items/id,items/namespaceUri');
          await ctx.sync();
          return {
            parts: parts.items.map((p: any) => ({
              id: p.id,
              namespaceUri: p.namespaceUri,
            })),
          };
        });
      } catch (error) {
        return {
          error: `Excel.run failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    });
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});

export const excelSettingsGet = definePageTool({
  name: 'excel_settings_get',
  description:
    'Reads add-in document settings from Office.context.document.settings. Returns all keys or a single key’s value.',
  annotations: {
    category: ToolCategory.EXCEL,
    readOnlyHint: true,
  },
  schema: {
    key: zod
      .string()
      .optional()
      .describe(
        'If provided, return only this setting’s value. Otherwise, return all settings.',
      ),
  },
  handler: async (request, response) => {
    const args = {key: request.params.key};
    const result = await request.page.pptrPage.evaluate(async args => {
      const g = globalThis as typeof globalThis & {
        Office?: {
          context?: {
            document?: {
              settings?: {
                get: (key: string) => unknown;
                getAll?: () => Record<string, unknown>;
              };
            };
          };
        };
      };
      const settings = g.Office?.context?.document?.settings;
      if (!settings) {
        return {
          error: 'Office.context.document.settings not available',
        } as const;
      }
      try {
        if (args.key) {
          return {key: args.key, value: settings.get(args.key)};
        }
        if (typeof settings.getAll === 'function') {
          return {settings: settings.getAll()};
        }
        return {
          error: 'settings.getAll is not available on this host',
        } as const;
      } catch (error) {
        return {
          error: `settings read failed: ${(error as Error)?.message ?? String(error)}`,
        } as const;
      }
    }, args);
    if ('error' in result) {
      response.appendResponseLine(`ERROR: ${result.error}`);
      return;
    }
    response.appendResponseLine(JSON.stringify(result, null, 2));
  },
});
