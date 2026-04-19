/**
 * @license
 * Copyright 2026 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import {logger} from '../logger.js';

import {ConnectionError} from './error.js';

const RECONNECT_WINDOW_MS = 60_000;
const RECONNECT_MAX_IN_WINDOW = 3;

interface SessionState {
  stale: boolean;
  staleSince: number | null;
  reconnectCount: number;
  stickyError: ConnectionError | undefined;
  reconnectTimestamps: number[];
}

const state: SessionState = {
  stale: false,
  staleSince: null,
  reconnectCount: 0,
  stickyError: undefined,
  reconnectTimestamps: [],
};

export function getSessionState(): Readonly<SessionState> {
  return state;
}

export function isSessionStale(): boolean {
  return state.stale;
}

export function getStickyError(): ConnectionError | undefined {
  return state.stickyError;
}

/**
 * Called from a `browser.on('disconnected')` listener. Marks the session as
 * stale so the next tool call triggers a reconnect via
 * {@link ensureBrowserConnected}.
 */
export function markDisconnected(now: () => number = Date.now): void {
  if (state.stale) {
    return;
  }
  state.stale = true;
  state.staleSince = now();
  logger('CDP connection lost; next tool call will attempt reconnect.');
}

/**
 * Called at the start of a reconnect attempt. Enforces the session-level
 * reconnect cap: more than {@link RECONNECT_MAX_IN_WINDOW} reconnects within
 * {@link RECONNECT_WINDOW_MS} returns (and latches) a sticky ConnectionError
 * that callers must throw. Otherwise clears the stale flag so the caller can
 * proceed with a fresh connect.
 */
export function beginReconnect(
  url: string,
  now: () => number = Date.now,
): ConnectionError | null {
  if (state.stickyError) {
    return state.stickyError;
  }
  const t = now();
  state.reconnectTimestamps = state.reconnectTimestamps.filter(
    ts => t - ts < RECONNECT_WINDOW_MS,
  );
  state.reconnectTimestamps.push(t);
  state.reconnectCount++;
  if (state.reconnectTimestamps.length > RECONNECT_MAX_IN_WINDOW) {
    state.stickyError = new ConnectionError({
      url,
      reason: 'unstable',
      hint: 'Connection is unstable — restart the MCP server.',
    });
    return state.stickyError;
  }
  state.stale = false;
  state.staleSince = null;
  return null;
}

/** Test-only: wipe session state between cases. */
export function resetSessionStateForTests(): void {
  state.stale = false;
  state.staleSince = null;
  state.reconnectCount = 0;
  state.stickyError = undefined;
  state.reconnectTimestamps = [];
}
