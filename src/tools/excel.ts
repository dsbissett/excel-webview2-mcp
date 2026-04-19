/**
 * @license
 * Copyright 2026 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import {ToolCategory} from './categories.js';
import {definePageTool} from './ToolDefinition.js';

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
