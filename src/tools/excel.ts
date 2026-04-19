/**
 * @license
 * Copyright 2026 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import {zod} from '../third_party/index.js';

import {ToolCategory} from './categories.js';
import {definePageTool} from './ToolDefinition.js';

const MAX_CELLS = 1000;

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
