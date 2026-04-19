/**
 * @license
 * Copyright 2026 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import assert from 'node:assert';

import type {TestScenario} from '../eval_gemini.ts';

const EXCEL_FIXTURE = `
  <script>
    window.Office = {
      HostType: { Excel: 'Excel' },
      PlatformType: { PC: 'PC' },
      context: {
        host: 'Excel',
        platform: 'PC',
        contentLanguage: 'en-US',
        displayLanguage: 'en-US',
        diagnostics: { host: 'Excel', platform: 'PC', version: '16.0.17928' },
        requirements: { isSetSupported: function () { return true; } },
        ui: {},
      },
    };
    window.Excel = {
      run: async function (batch) {
        const ctx = {
          workbook: {
            getSelectedRange: function () {
              return {
                address: 'Sheet1!B2:C3',
                values: [[1, 2], [3, 4]],
                rowCount: 2,
                columnCount: 2,
                formulas: [['=A1', '=A2'], ['=A3', '=A4']],
                numberFormat: [['General', 'General'], ['General', 'General']],
                load: function () {},
              };
            },
          },
          sync: async function () {},
        };
        return await batch(ctx);
      },
    };
  </script>
`;

export const scenario: TestScenario = {
  prompt:
    'Open <TEST_URL> and then read the currently selected Excel range. Include formulas in the result.',
  maxTurns: 3,
  htmlRoute: {
    path: '/excel_active_range_test.html',
    htmlContent: `<!doctype html><title>Excel Range Fixture</title>${EXCEL_FIXTURE}<h1>Excel Range Fixture</h1>`,
  },
  expectations: calls => {
    const toolNames = calls.map(c => c.name);
    assert.ok(
      toolNames.includes('excel_active_range'),
      `Expected excel_active_range to be called, got: ${toolNames.join(', ')}`,
    );
    const rangeCall = calls.find(c => c.name === 'excel_active_range');
    assert.ok(rangeCall, 'excel_active_range call should be captured');
    assert.strictEqual(
      rangeCall!.args.includeFormulas,
      true,
      'Scenario asked for formulas — tool should be invoked with includeFormulas: true',
    );
  },
};
