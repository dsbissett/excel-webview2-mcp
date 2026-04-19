import assert from 'node:assert';
import {afterEach, beforeEach, describe, it} from 'node:test';

import {probeCdpEndpoint} from '../../src/connection/probe.js';

type FetchImpl = typeof fetch;

describe('probeCdpEndpoint', () => {
  let originalFetch: FetchImpl;

  beforeEach(() => {
    originalFetch = globalThis.fetch;
  });

  afterEach(() => {
    globalThis.fetch = originalFetch;
  });

  it('returns ok with version on 2xx + valid JSON', async () => {
    globalThis.fetch = (async () => {
      return new Response(
        JSON.stringify({
          Browser: 'Edge/135.0.0.0',
          'Protocol-Version': '1.3',
        }),
        {status: 200, headers: {'content-type': 'application/json'}},
      );
    }) as FetchImpl;

    const result = await probeCdpEndpoint('http://localhost:9222', 1000);
    assert.deepStrictEqual(result, {ok: true, version: 'Edge/135.0.0.0'});
  });

  it('returns timeout when fetch aborts via TimeoutError', async () => {
    globalThis.fetch = (async () => {
      const err = new Error('timed out');
      err.name = 'TimeoutError';
      throw err;
    }) as FetchImpl;

    const result = await probeCdpEndpoint('http://localhost:9222', 50);
    assert.deepStrictEqual(result, {ok: false, reason: 'timeout'});
  });

  it('returns unreachable on network error', async () => {
    globalThis.fetch = (async () => {
      throw new TypeError('fetch failed');
    }) as FetchImpl;

    const result = await probeCdpEndpoint('http://localhost:9222', 1000);
    assert.deepStrictEqual(result, {ok: false, reason: 'unreachable'});
  });

  it('returns http-error:<status> on non-2xx', async () => {
    globalThis.fetch = (async () => {
      return new Response('forbidden', {status: 403});
    }) as FetchImpl;

    const result = await probeCdpEndpoint('http://localhost:9222', 1000);
    assert.deepStrictEqual(result, {ok: false, reason: 'http-error:403'});
  });

  it('returns invalid-response when JSON is unparseable', async () => {
    globalThis.fetch = (async () => {
      return new Response('<html>not json</html>', {
        status: 200,
        headers: {'content-type': 'text/html'},
      });
    }) as FetchImpl;

    const result = await probeCdpEndpoint('http://localhost:9222', 1000);
    assert.deepStrictEqual(result, {ok: false, reason: 'invalid-response'});
  });

  it('returns invalid-response when Browser field is missing', async () => {
    globalThis.fetch = (async () => {
      return new Response(JSON.stringify({'Protocol-Version': '1.3'}), {
        status: 200,
        headers: {'content-type': 'application/json'},
      });
    }) as FetchImpl;

    const result = await probeCdpEndpoint('http://localhost:9222', 1000);
    assert.deepStrictEqual(result, {ok: false, reason: 'invalid-response'});
  });

  it('strips trailing slash from url before appending /json/version', async () => {
    let observedUrl = '';
    globalThis.fetch = (async (input: RequestInfo | URL) => {
      observedUrl = typeof input === 'string' ? input : input.toString();
      return new Response(JSON.stringify({Browser: 'x'}), {status: 200});
    }) as FetchImpl;

    await probeCdpEndpoint('http://localhost:9222/', 1000);
    assert.strictEqual(observedUrl, 'http://localhost:9222/json/version');
  });
});
