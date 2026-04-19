import assert from 'node:assert';
import {EventEmitter} from 'node:events';
import {afterEach, beforeEach, describe, it} from 'node:test';

import {ConnectionError} from '../../src/connection/error.js';
import {
  beginReconnect,
  getSessionState,
  getStickyError,
  isSessionStale,
  markDisconnected,
  resetSessionStateForTests,
} from '../../src/connection/session.js';

const URL = 'http://localhost:9222';

describe('session state', () => {
  beforeEach(() => {
    resetSessionStateForTests();
  });

  afterEach(() => {
    resetSessionStateForTests();
  });

  it('starts not-stale with no sticky error', () => {
    assert.strictEqual(isSessionStale(), false);
    assert.strictEqual(getStickyError(), undefined);
    assert.strictEqual(getSessionState().reconnectCount, 0);
  });

  it('markDisconnected flips stale and records staleSince', () => {
    markDisconnected(() => 42_000);
    assert.strictEqual(isSessionStale(), true);
    assert.strictEqual(getSessionState().staleSince, 42_000);
  });

  it('markDisconnected is idempotent within a stale window', () => {
    markDisconnected(() => 1000);
    markDisconnected(() => 2000);
    assert.strictEqual(getSessionState().staleSince, 1000);
  });

  it('beginReconnect allows up to 3 reconnects in the 60s window', () => {
    let t = 0;
    const now = () => t;
    for (let i = 0; i < 3; i++) {
      markDisconnected(now);
      t += 1000;
      const blocker = beginReconnect(URL, now);
      assert.strictEqual(blocker, null, `attempt ${i + 1} should succeed`);
      assert.strictEqual(isSessionStale(), false);
      t += 1000;
    }
    assert.strictEqual(getSessionState().reconnectCount, 3);
    assert.strictEqual(getStickyError(), undefined);
  });

  it('beginReconnect latches a sticky error on the 4th reconnect in 60s', () => {
    let t = 0;
    const now = () => t;
    for (let i = 0; i < 3; i++) {
      markDisconnected(now);
      t += 500;
      assert.strictEqual(beginReconnect(URL, now), null);
      t += 500;
    }
    markDisconnected(now);
    t += 500;
    const blocker = beginReconnect(URL, now);
    assert.ok(blocker instanceof ConnectionError);
    assert.strictEqual(blocker.reason, 'unstable');
    assert.match(blocker.message, /Connection is unstable/);

    // Subsequent calls keep returning the same sticky error.
    const again = beginReconnect(URL, now);
    assert.strictEqual(again, blocker);
    assert.strictEqual(getStickyError(), blocker);
  });

  it('drops reconnect timestamps older than 60s', () => {
    let t = 0;
    const now = () => t;
    for (let i = 0; i < 3; i++) {
      markDisconnected(now);
      t += 100;
      assert.strictEqual(beginReconnect(URL, now), null);
      t += 100;
    }
    // Jump 61s into the future — prior timestamps expire.
    t += 61_000;
    markDisconnected(now);
    const blocker = beginReconnect(URL, now);
    assert.strictEqual(blocker, null);
    assert.strictEqual(getStickyError(), undefined);
  });
});

describe('disconnect event integration', () => {
  beforeEach(() => {
    resetSessionStateForTests();
  });

  afterEach(() => {
    resetSessionStateForTests();
  });

  it('firing disconnected on an EventEmitter-backed browser flips the stale flag', () => {
    const emitter = new EventEmitter();
    emitter.on('disconnected', () => markDisconnected(() => 123));
    assert.strictEqual(isSessionStale(), false);
    emitter.emit('disconnected');
    assert.strictEqual(isSessionStale(), true);
    assert.strictEqual(getSessionState().staleSince, 123);
  });
});
