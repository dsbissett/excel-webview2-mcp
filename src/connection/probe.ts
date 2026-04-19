/**
 * @license
 * Copyright 2026 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import type {ConnectionErrorReason} from './error.js';

export type ProbeFailureReason = Exclude<
  ConnectionErrorReason,
  'connect-failed'
>;

export interface ProbeResult {
  ok: boolean;
  version?: string;
  reason?: ProbeFailureReason;
}

/**
 * GET {url}/json/version with a bounded timeout. Never throws — always
 * resolves to a {@link ProbeResult}. Used as a fail-fast liveness check
 * before handing off to puppeteer.connect.
 */
export async function probeCdpEndpoint(
  url: string,
  timeoutMs: number,
): Promise<ProbeResult> {
  const endpoint = `${url.replace(/\/+$/, '')}/json/version`;
  let response: Response;
  try {
    response = await fetch(endpoint, {
      signal: AbortSignal.timeout(timeoutMs),
    });
  } catch (err) {
    const name = (err as {name?: string} | null)?.name;
    if (name === 'TimeoutError' || name === 'AbortError') {
      return {ok: false, reason: 'timeout'};
    }
    return {ok: false, reason: 'unreachable'};
  }

  if (!response.ok) {
    return {
      ok: false,
      reason: `http-error:${response.status}` as ProbeFailureReason,
    };
  }

  let body: unknown;
  try {
    body = await response.json();
  } catch {
    return {ok: false, reason: 'invalid-response'};
  }

  const version =
    typeof body === 'object' && body !== null && 'Browser' in body
      ? (body as {Browser?: unknown}).Browser
      : undefined;
  if (typeof version !== 'string') {
    return {ok: false, reason: 'invalid-response'};
  }

  return {ok: true, version};
}
