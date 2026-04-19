/**
 * @license
 * Copyright 2026 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import type {ProbeFailureReason, ProbeResult} from './probe.js';
import {probeCdpEndpoint} from './probe.js';
import {getSessionState, isSessionStale} from './session.js';

export type ConnectionEndpointSource =
  | 'browserUrl'
  | 'wsEndpoint'
  | 'autoDetect'
  | 'default';

export type ConnectionProbeResult = 'ok' | ProbeFailureReason | null;

export interface ConnectionStatusSnapshot {
  attached: boolean;
  endpointUrl: string;
  endpointSource: ConnectionEndpointSource;
  cdpVersion: string | null;
  lastProbeResult: ConnectionProbeResult;
  reconnectCount: number;
  staleSince: string | null;
}

interface ConnectionStatusState {
  attached: boolean;
  endpointUrl: string;
  endpointSource: ConnectionEndpointSource;
  cdpVersion: string | null;
  lastProbeResult: ConnectionProbeResult;
}

const state: ConnectionStatusState = {
  attached: false,
  endpointUrl: '',
  endpointSource: 'default',
  cdpVersion: null,
  lastProbeResult: null,
};

function deriveProbeUrl(endpointUrl: string): string | null {
  if (!endpointUrl) {
    return null;
  }

  let parsed: URL;
  try {
    parsed = new URL(endpointUrl);
  } catch {
    return null;
  }

  if (parsed.protocol === 'http:' || parsed.protocol === 'https:') {
    return endpointUrl;
  }

  if (parsed.protocol === 'ws:' || parsed.protocol === 'wss:') {
    const protocol = parsed.protocol === 'ws:' ? 'http:' : 'https:';
    return `${protocol}//${parsed.host}`;
  }

  return null;
}

export function setConnectionEndpoint(
  endpointUrl: string,
  endpointSource: ConnectionEndpointSource,
): void {
  const changed =
    state.endpointUrl !== endpointUrl ||
    state.endpointSource !== endpointSource;
  state.endpointUrl = endpointUrl;
  state.endpointSource = endpointSource;
  if (changed) {
    state.cdpVersion = null;
    state.lastProbeResult = null;
  }
}

export function markConnectionAttached(endpointUrl?: string): void {
  state.attached = true;
  if (endpointUrl) {
    state.endpointUrl = endpointUrl;
  }
}

export function markConnectionDetached(): void {
  state.attached = false;
}

export function recordProbeResult(result: ProbeResult): void {
  state.lastProbeResult = result.ok ? 'ok' : (result.reason ?? 'unreachable');
  state.cdpVersion = result.ok ? (result.version ?? null) : null;
}

export function getConnectionStatusSnapshot(): ConnectionStatusSnapshot {
  const sessionState = getSessionState();
  return {
    attached: state.attached && !isSessionStale(),
    endpointUrl: state.endpointUrl,
    endpointSource: state.endpointSource,
    cdpVersion: state.cdpVersion,
    lastProbeResult: state.lastProbeResult,
    reconnectCount: sessionState.reconnectCount,
    staleSince:
      sessionState.staleSince === null
        ? null
        : new Date(sessionState.staleSince).toISOString(),
  };
}

export async function refreshConnectionStatusProbe(
  timeoutMs: number,
): Promise<ConnectionStatusSnapshot> {
  const probeUrl = deriveProbeUrl(state.endpointUrl);
  if (!probeUrl) {
    return getConnectionStatusSnapshot();
  }

  recordProbeResult(await probeCdpEndpoint(probeUrl, timeoutMs));
  return getConnectionStatusSnapshot();
}

export function resetConnectionStatusForTests(): void {
  state.attached = false;
  state.endpointUrl = '';
  state.endpointSource = 'default';
  state.cdpVersion = null;
  state.lastProbeResult = null;
}
