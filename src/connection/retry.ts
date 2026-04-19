/**
 * @license
 * Copyright 2026 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import {logger} from '../logger.js';
import type {Browser} from '../third_party/index.js';

import {ConnectionError} from './error.js';
import {probeCdpEndpoint, type ProbeResult} from './probe.js';

export interface ConnectWithRetryOptions {
  browserURL: string;
  probeTimeoutMs: number;
  retryBudgetMs: number;
  verbose?: boolean;
  connect: () => Promise<Browser>;
  onProbeResult?: (result: ProbeResult) => void;
  sleep?: (ms: number) => Promise<void>;
  now?: () => number;
}

const BACKOFF_SEQUENCE_MS = [500, 1000, 2000, 4000, 5000];
const BACKOFF_CAP_MS = 5000;

function defaultSleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function hintFor(url: string): string {
  return `Run: curl ${url}/json/version — if this fails, your Excel add-in isn't exposing the debug port.`;
}

function logAttempt(
  verbose: boolean,
  attempt: number,
  reason: string,
  delayMs: number | null,
): void {
  const suffix = delayMs === null ? 'giving up.' : `retrying in ${delayMs}ms.`;
  const message = `Connect attempt ${attempt} failed (${reason}); ${suffix}`;
  logger(message);
  if (verbose) {
    console.error(`[excel-webview2-mcp] ${message}`);
  }
}

/**
 * Attempt to connect via CDP with an exponential-backoff retry loop bounded
 * by `retryBudgetMs`. Retries only on `unreachable` / `timeout` probe
 * failures; other probe failures and connect failures are terminal.
 */
export async function connectWithRetry(
  options: ConnectWithRetryOptions,
): Promise<Browser> {
  const {
    browserURL,
    probeTimeoutMs,
    retryBudgetMs,
    verbose = false,
    connect,
    onProbeResult,
    sleep = defaultSleep,
    now = Date.now,
  } = options;

  const start = now();
  let attempt = 0;

  // Always run at least one attempt, even when retryBudgetMs === 0.
  while (true) {
    attempt++;
    const probe = await probeCdpEndpoint(browserURL, probeTimeoutMs);
    onProbeResult?.(probe);

    if (probe.ok) {
      try {
        return await connect();
      } catch (err) {
        throw new ConnectionError({
          url: browserURL,
          reason: 'connect-failed',
          hint: hintFor(browserURL),
          cause: err,
        });
      }
    }

    const reason = probe.reason ?? 'unreachable';
    const retryable = reason === 'unreachable' || reason === 'timeout';
    const elapsed = now() - start;

    if (!retryable || retryBudgetMs <= 0 || elapsed >= retryBudgetMs) {
      if (retryable && attempt > 1) {
        logAttempt(verbose, attempt, reason, null);
      }
      throw new ConnectionError({
        url: browserURL,
        reason,
        hint: hintFor(browserURL),
      });
    }

    const baseDelay =
      BACKOFF_SEQUENCE_MS[
        Math.min(attempt - 1, BACKOFF_SEQUENCE_MS.length - 1)
      ] ?? BACKOFF_CAP_MS;
    const remaining = retryBudgetMs - elapsed;
    const delay = Math.min(baseDelay, remaining);
    logAttempt(verbose, attempt, reason, delay);
    await sleep(delay);
  }
}
