/**
 * @license
 * Copyright 2026 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import assert from 'node:assert';

import type {TestScenario} from '../eval_gemini.ts';

const OFFICE_FIXTURE = `
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
        requirements: {
          isSetSupported: function (name, version) {
            if (name === 'ExcelApi') return true;
            if (name === 'DialogApi') return true;
            if (name === 'SharedRuntime') return true;
            return false;
          },
        },
        ui: {},
      },
    };
    window.Excel = { run: async function () {} };
  </script>
`;

export const scenario: TestScenario = {
  prompt:
    'Open <TEST_URL> and then report the Excel add-in context info — host, platform, and which requirement sets the page supports.',
  maxTurns: 3,
  htmlRoute: {
    path: '/excel_context_test.html',
    htmlContent: `<!doctype html><title>Excel Fixture</title>${OFFICE_FIXTURE}<h1>Excel Add-in Fixture</h1>`,
  },
  expectations: calls => {
    const toolNames = calls.map(c => c.name);
    assert.ok(
      toolNames.includes('excel_context_info'),
      `Expected excel_context_info to be called, got: ${toolNames.join(', ')}`,
    );
  },
};
