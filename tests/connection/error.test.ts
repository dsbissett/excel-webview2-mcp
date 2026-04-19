/**
 * @license
 * Copyright 2026 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import assert from 'node:assert';
import {describe, it} from 'node:test';

import {ConnectionError} from '../../src/connection/error.js';

describe('ConnectionError', () => {
  it('formats the canonical multi-line message for an http endpoint', () => {
    const err = new ConnectionError({
      url: 'http://127.0.0.1:9222',
      reason: 'unreachable',
      hint:
        "Run: curl http://127.0.0.1:9222/json/version — if this fails, your Excel add-in isn't exposing the debug port.",
    });

    assert.strictEqual(
      err.format(),
      `Excel WebView2 debug endpoint is not reachable.
  Endpoint: http://127.0.0.1:9222
  Reason:   unreachable
  Verify:   curl http://127.0.0.1:9222/json/version
  Docs:     https://github.com/dsbissett/excel-webview2-mcp#launching-excel-with-debug-port`,
    );
    assert.strictEqual(err.message, err.format());
  });

  it('normalizes a websocket endpoint to an http verify command', () => {
    const err = new ConnectionError({
      url: 'ws://127.0.0.1:9222/devtools/browser/test-browser-id',
      reason: 'connect-failed',
      hint:
        "Run: curl http://127.0.0.1:9222/json/version — if this fails, your Excel add-in isn't exposing the debug port.",
    });

    assert.match(
      err.format(),
      /Verify:\s+curl http:\/\/127\.0\.0\.1:9222\/json\/version/,
    );
  });
});
