import assert from 'node:assert';
import path from 'node:path';
import {describe, it} from 'node:test';

import {Client} from '@modelcontextprotocol/sdk/client/index.js';
import {StdioClientTransport} from '@modelcontextprotocol/sdk/client/stdio.js';

import {getTextContent} from './utils.js';

describe('connection_status tool', () => {
  it('returns status without forcing browser launch', async () => {
    const transport = new StdioClientTransport({
      command: 'node',
      args: [
        'build/src/bin/excel-webview2-mcp.js',
        '--headless',
        '--isolated',
        '--executable-path',
        path.resolve('tests/does-not-exist-browser'),
      ],
    });
    const client = new Client(
      {
        name: 'connection-status-test',
        version: '1.0.0',
      },
      {
        capabilities: {},
      },
    );

    try {
      await client.connect(transport);
      const result = await client.callTool({
        name: 'connection_status',
        arguments: {},
      });

      assert.notStrictEqual(result.isError, true);
      const payload = JSON.parse(
        getTextContent(
          (result.content as Array<{type: 'text'; text: string}>)[0],
        ),
      ) as {
        attached: boolean;
      };
      assert.strictEqual(payload.attached, false);
    } finally {
      await client.close();
    }
  });
});
