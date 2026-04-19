/**
 * @license
 * Copyright 2026 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import assert from 'node:assert';
import {afterEach, beforeEach, describe, it} from 'node:test';

import {ConnectionError} from '../../src/connection/error.js';
import {connectWithRetry} from '../../src/connection/retry.js';
import type {Browser} from '../../src/third_party/index.js';

type FetchImpl = typeof fetch;

function okProbeResponse(): Response {
  return new Response(JSON.stringify({Browser: 'Edge/135.0.0.0'}), {
    status: 200,
    headers: {'content-type': 'application/json'},
  });
}

function unreachableFetch(): FetchImpl {
  return (async () => {
    throw new TypeError('fetch failed');
  }) as FetchImpl;
}

function httpErrorFetch(status: number): FetchImpl {
  return (async () => new Response('nope', {status})) as FetchImpl;
}

function makeSleepSpy() {
  const delays: number[] = [];
  const sleep = async (ms: number) => {
    delays.push(ms);
  };
  return {delays, sleep};
}

function makeClock(start = 1_000_000) {
  let t = start;
  return {
    now: () => t,
    advance: (ms: number) => {
      t += ms;
    },
  };
}

const fakeBrowser = {connected: true} as unknown as Browser;

describe('connectWithRetry', () => {
  let originalFetch: FetchImpl;

  beforeEach(() => {
    originalFetch = globalThis.fetch;
  });

  afterEach(() => {
    globalThis.fetch = originalFetch;
  });

  it('short-circuits on first success without sleeping', async () => {
    globalThis.fetch = (async () => okProbeResponse()) as FetchImpl;
    const {delays, sleep} = makeSleepSpy();
    let connectCalls = 0;

    const browser = await connectWithRetry({
      browserURL: 'http://localhost:9222',
      probeTimeoutMs: 1000,
      retryBudgetMs: 15000,
      connect: async () => {
        connectCalls++;
        return fakeBrowser;
      },
      sleep,
      now: () => 0,
    });

    assert.strictEqual(browser, fakeBrowser);
    assert.strictEqual(connectCalls, 1);
    assert.deepStrictEqual(delays, []);
  });

  it('retries on unreachable and succeeds within the budget', async () => {
    let probeCalls = 0;
    globalThis.fetch = (async () => {
      probeCalls++;
      if (probeCalls < 3) {
        throw new TypeError('fetch failed');
      }
      return okProbeResponse();
    }) as FetchImpl;
    const {delays, sleep} = makeSleepSpy();
    const clock = makeClock();

    const browser = await connectWithRetry({
      browserURL: 'http://localhost:9222',
      probeTimeoutMs: 1000,
      retryBudgetMs: 15000,
      connect: async () => fakeBrowser,
      sleep: async ms => {
        await sleep(ms);
        clock.advance(ms);
      },
      now: clock.now,
    });

    assert.strictEqual(browser, fakeBrowser);
    assert.strictEqual(probeCalls, 3);
    assert.deepStrictEqual(delays, [500, 1000]);
  });

  it('bounds backoff to the cap and stops at the budget', async () => {
    globalThis.fetch = unreachableFetch();
    const {delays, sleep} = makeSleepSpy();
    const clock = makeClock();

    await assert.rejects(
      () =>
        connectWithRetry({
          browserURL: 'http://localhost:9222',
          probeTimeoutMs: 1000,
          retryBudgetMs: 15000,
          connect: async () => fakeBrowser,
          sleep: async ms => {
            await sleep(ms);
            clock.advance(ms);
          },
          now: clock.now,
        }),
      err => {
        assert.ok(err instanceof ConnectionError);
        assert.strictEqual(err.reason, 'unreachable');
        return true;
      },
    );

    // 500 + 1000 + 2000 + 4000 + 5000 = 12500 elapsed; remaining budget is
    // 2500ms which is used for one final capped sleep, then the loop exits.
    assert.deepStrictEqual(delays, [500, 1000, 2000, 4000, 5000, 2500]);
    for (const d of delays) {
      assert.ok(d <= 5000, `delay ${d} exceeded cap`);
    }
  });

  it('fails fast with retryBudgetMs=0 and still makes one attempt', async () => {
    let probeCalls = 0;
    globalThis.fetch = (async () => {
      probeCalls++;
      throw new TypeError('fetch failed');
    }) as FetchImpl;
    const {delays, sleep} = makeSleepSpy();

    await assert.rejects(
      () =>
        connectWithRetry({
          browserURL: 'http://localhost:9222',
          probeTimeoutMs: 1000,
          retryBudgetMs: 0,
          connect: async () => fakeBrowser,
          sleep,
          now: () => 0,
        }),
      err => {
        assert.ok(err instanceof ConnectionError);
        assert.strictEqual(err.reason, 'unreachable');
        return true;
      },
    );

    assert.strictEqual(probeCalls, 1);
    assert.deepStrictEqual(delays, []);
  });

  it('does not retry on terminal http-error probe result', async () => {
    let probeCalls = 0;
    globalThis.fetch = (async () => {
      probeCalls++;
      return httpErrorFetch(403)(
        'http://localhost:9222/json/version',
      ) as unknown as Response;
    }) as FetchImpl;
    const {delays, sleep} = makeSleepSpy();

    await assert.rejects(
      () =>
        connectWithRetry({
          browserURL: 'http://localhost:9222',
          probeTimeoutMs: 1000,
          retryBudgetMs: 15000,
          connect: async () => fakeBrowser,
          sleep,
          now: () => 0,
        }),
      err => {
        assert.ok(err instanceof ConnectionError);
        assert.strictEqual(err.reason, 'http-error:403');
        return true;
      },
    );

    assert.strictEqual(probeCalls, 1);
    assert.deepStrictEqual(delays, []);
  });

  it('wraps a connect failure in a ConnectionError with reason connect-failed', async () => {
    globalThis.fetch = (async () => okProbeResponse()) as FetchImpl;
    const {sleep} = makeSleepSpy();

    await assert.rejects(
      () =>
        connectWithRetry({
          browserURL: 'http://localhost:9222',
          probeTimeoutMs: 1000,
          retryBudgetMs: 15000,
          connect: async () => {
            throw new Error('puppeteer exploded');
          },
          sleep,
          now: () => 0,
        }),
      err => {
        assert.ok(err instanceof ConnectionError);
        assert.strictEqual(err.reason, 'connect-failed');
        assert.ok(err.cause instanceof Error);
        return true;
      },
    );
  });
});
