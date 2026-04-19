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

export class ConnectionError extends Error {
  readonly url: string;
  readonly reason: ConnectionErrorReason;
  readonly hint: string;

  constructor(options: ConnectionErrorOptions) {
    super(
      `Excel WebView2 debug endpoint is not reachable at ${options.url} (${options.reason}). ${options.hint}`,
      {cause: options.cause},
    );
    this.name = 'ConnectionError';
    this.url = options.url;
    this.reason = options.reason;
    this.hint = options.hint;
  }
}
