/**
 * @license
 * Copyright 2026 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import assert from 'node:assert';
import {afterEach, beforeEach, describe, it} from 'node:test';

import {
  markDisconnected,
  resetSessionStateForTests,
} from '../../src/connection/session.js';
import {
  getConnectionStatusSnapshot,
  markConnectionAttached,
  markConnectionDetached,
  recordProbeResult,
  refreshConnectionStatusProbe,
  resetConnectionStatusForTests,
  setConnectionEndpoint,
} from '../../src/connection/status.js';

type FetchImpl = typeof fetch;

function okProbeResponse(): Response {
  return new Response(JSON.stringify({Browser: 'Edge/135.0.0.0'}), {
    status: 200,
    headers: {'content-type': 'application/json'},
  });
}

describe('connection status', () => {
  let originalFetch: FetchImpl;

  beforeEach(() => {
    originalFetch = globalThis.fetch;
    resetConnectionStatusForTests();
    resetSessionStateForTests();
  });

  afterEach(() => {
    globalThis.fetch = originalFetch;
    resetConnectionStatusForTests();
    resetSessionStateForTests();
  });

  it('returns cached endpoint, probe, and stale-session metadata', () => {
    setConnectionEndpoint('http://localhost:9222', 'default');
    recordProbeResult({ok: true, version: 'Edge/135.0.0.0'});
    markConnectionAttached();
    markDisconnected(() => 0);
    markConnectionDetached();

    const snapshot = getConnectionStatusSnapshot();
    assert.deepStrictEqual(snapshot, {
      attached: false,
      endpointUrl: 'http://localhost:9222',
      endpointSource: 'default',
      cdpVersion: 'Edge/135.0.0.0',
      lastProbeResult: 'ok',
      reconnectCount: 0,
      staleSince: '1970-01-01T00:00:00.000Z',
    });
  });

  it('re-probes a tracked ws endpoint via its HTTP origin', async () => {
    let requestedUrl = '';
    globalThis.fetch = (async (input: string | URL | Request) => {
      requestedUrl = String(input);
      return okProbeResponse();
    }) as FetchImpl;
    setConnectionEndpoint(
      'ws://127.0.0.1:9222/devtools/browser/test-browser-id',
      'wsEndpoint',
    );

    const snapshot = await refreshConnectionStatusProbe(1000);

    assert.strictEqual(requestedUrl, 'http://127.0.0.1:9222/json/version');
    assert.strictEqual(snapshot.lastProbeResult, 'ok');
    assert.strictEqual(snapshot.cdpVersion, 'Edge/135.0.0.0');
    assert.strictEqual(snapshot.endpointSource, 'wsEndpoint');
  });
});
