/**
 * @license
 * Copyright 2026 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import assert from 'node:assert';
import path from 'node:path';
import {describe, it} from 'node:test';

import {Client} from '@modelcontextprotocol/sdk/client/index.js';
import {StdioClientTransport} from '@modelcontextprotocol/sdk/client/stdio.js';

import {getTextContent} from './utils.js';

describe('connection error response', () => {
  it('returns the formatted ConnectionError to the MCP client without cause text', async () => {
    const transport = new StdioClientTransport({
      command: 'node',
      args: [
        path.resolve('build/src/bin/excel-webview2-mcp.js'),
        '--browserUrl',
        'http://127.0.0.1:9',
        '--connectRetryBudget',
        '0',
        '--connectTimeout',
        '200',
      ],
    });
    const client = new Client(
      {
        name: 'connection-error-response-test',
        version: '1.0.0',
      },
      {
        capabilities: {},
      },
    );

    try {
      await client.connect(transport);
      const result = await client.callTool({
        name: 'list_pages',
        arguments: {},
      });
      const text = getTextContent(
        (result.content as Array<{type: 'text'; text: string}>)[0],
      );

      assert.strictEqual(result.isError, true);
      assert.match(
        text,
        /^Excel WebView2 debug endpoint is not reachable\.\n {2}Endpoint: http:\/\/127\.0\.0\.1:9\n {2}Reason:\s+unreachable\n {2}Verify:\s+curl http:\/\/127\.0\.0\.1:9\/json\/version\n {2}Docs:\s+https:\/\/github\.com\/dsbissett\/excel-webview2-mcp#launching-excel-with-debug-port$/,
      );
      assert.ok(!text.includes('\nCause:'), text);
    } finally {
      await client.close();
    }
  });
});
