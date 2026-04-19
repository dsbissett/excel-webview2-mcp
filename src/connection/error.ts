/**
 * @license
 * Copyright 2026 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

export type ConnectionErrorReason =
  | 'unreachable'
  | 'timeout'
  | 'invalid-response'
  | `http-error:${number}`
  | 'connect-failed'
  | 'unstable';

export interface ConnectionErrorOptions {
  url: string;
  reason: ConnectionErrorReason;
  hint: string;
  cause?: unknown;
}

export const CONNECTION_DOCS_URL =
  'https://github.com/dsbissett/excel-webview2-mcp#launching-excel-with-debug-port';

function trimTrailingSlash(value: string): string {
  return value.replace(/\/+$/, '');
}

function toVerifyBaseUrl(url: string): string {
  try {
    const parsed = new URL(url);
    if (parsed.protocol === 'ws:' || parsed.protocol === 'wss:') {
      const protocol = parsed.protocol === 'ws:' ? 'http:' : 'https:';
      return `${protocol}//${parsed.host}`;
    }
    return trimTrailingSlash(url);
  } catch {
    return trimTrailingSlash(url);
  }
}

export function getConnectionVerifyCommand(url: string): string {
  return `curl ${toVerifyBaseUrl(url)}/json/version`;
}

export function getDefaultConnectionHint(url: string): string {
  return `Run: ${getConnectionVerifyCommand(url)} — if this fails, your Excel add-in isn't exposing the debug port.`;
}

function formatReason(
  reason: ConnectionErrorReason,
  hint: string,
): ConnectionErrorReason | string {
  if (reason === 'unstable' && hint) {
    return hint;
  }
  return reason;
}

function shouldAppendHint(options: ConnectionErrorOptions): boolean {
  return Boolean(
    options.hint &&
    options.reason !== 'unstable' &&
    options.hint !== getDefaultConnectionHint(options.url),
  );
}

function formatConnectionError(options: ConnectionErrorOptions): string {
  const lines = [
    'Excel WebView2 debug endpoint is not reachable.',
    `  Endpoint: ${options.url}`,
    `  Reason:   ${formatReason(options.reason, options.hint)}`,
  ];
  if (shouldAppendHint(options)) {
    lines.push(`  Hint:     ${options.hint}`);
  }
  lines.push(`  Verify:   ${getConnectionVerifyCommand(options.url)}`);
  lines.push(`  Docs:     ${CONNECTION_DOCS_URL}`);
  return lines.join('\n');
}

export class ConnectionError extends Error {
  readonly url: string;
  readonly reason: ConnectionErrorReason;
  readonly hint: string;

  constructor(options: ConnectionErrorOptions) {
    super(formatConnectionError(options), {cause: options.cause});
    this.name = 'ConnectionError';
    this.url = options.url;
    this.reason = options.reason;
    this.hint = options.hint;
  }

  format(): string {
    return formatConnectionError({
      url: this.url,
      reason: this.reason,
      hint: this.hint,
      cause: this.cause,
    });
  }
}
