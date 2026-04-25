import {ConnectionError} from '../connection/error.js';

export type ErrorCategory =
  | 'auth'
  | 'not_found'
  | 'rate_limit'
  | 'timeout'
  | 'validation'
  | 'connection'
  | 'protocol'
  | 'cancelled'
  | 'unsupported'
  | 'internal';

export interface ToolErrorContext {
  toolName: string;
  attempted: string;
  failed: string;
  params?: unknown;
  details?: Record<string, unknown>;
}

export interface StructuredToolError {
  isError: true;
  errorCategory: ErrorCategory;
  isRetryable: boolean;
  context: ToolErrorContext;
}

export interface ToolErrorInit {
  category: ErrorCategory;
  isRetryable: boolean;
  message: string;
  context: ToolErrorContext;
  cause?: unknown;
}

export class ToolError extends Error {
  readonly category: ErrorCategory;
  readonly isRetryable: boolean;
  readonly context: ToolErrorContext;

  constructor(init: ToolErrorInit) {
    super(init.message, {cause: init.cause});
    this.name = 'ToolError';
    this.category = init.category;
    this.isRetryable = init.isRetryable;
    this.context = init.context;
  }

  toStructured(): StructuredToolError {
    return {
      isError: true,
      errorCategory: this.category,
      isRetryable: this.isRetryable,
      context: this.context,
    };
  }
}

function nameOf(err: unknown): string {
  if (err && typeof err === 'object' && 'name' in err) {
    const n = (err as {name?: unknown}).name;
    if (typeof n === 'string') return n;
  }
  return '';
}

function messageOf(err: unknown): string {
  if (err instanceof Error) return err.message;
  if (err && typeof err === 'object' && 'message' in err) {
    const m = (err as {message?: unknown}).message;
    if (typeof m === 'string') return m;
  }
  return String(err);
}

function classifyConnectionError(err: ConnectionError): {
  category: ErrorCategory;
  isRetryable: boolean;
} {
  const reason = err.reason;
  if (
    reason === 'timeout' ||
    reason === 'unreachable' ||
    reason === 'unstable' ||
    reason === 'connect-failed'
  ) {
    return {category: 'connection', isRetryable: true};
  }
  if (reason === 'invalid-response') {
    return {category: 'connection', isRetryable: false};
  }
  if (typeof reason === 'string' && reason.startsWith('http-error:')) {
    const code = Number(reason.split(':')[1]);
    if (code === 401 || code === 403) {
      return {category: 'auth', isRetryable: false};
    }
    if (code === 404) {
      return {category: 'not_found', isRetryable: false};
    }
    if (code === 408 || code === 429) {
      return {
        category: code === 429 ? 'rate_limit' : 'timeout',
        isRetryable: true,
      };
    }
    if (code >= 500) {
      return {category: 'connection', isRetryable: true};
    }
    return {category: 'connection', isRetryable: false};
  }
  return {category: 'connection', isRetryable: false};
}

export function classifyUnknownError(
  err: unknown,
  toolName: string,
  params?: unknown,
): ToolError {
  if (err instanceof ToolError) {
    return err;
  }

  const message = messageOf(err);
  const name = nameOf(err);

  if (err instanceof ConnectionError) {
    const {category, isRetryable} = classifyConnectionError(err);
    return new ToolError({
      category,
      isRetryable,
      message: err.format(),
      cause: err,
      context: {
        toolName,
        attempted: toolName,
        failed: `connection (${err.reason})`,
        params,
        details: {url: err.url, reason: err.reason},
      },
    });
  }

  if (name === 'TimeoutError' || /timed? ?out/i.test(message)) {
    return new ToolError({
      category: 'timeout',
      isRetryable: true,
      message,
      cause: err,
      context: {toolName, attempted: toolName, failed: 'timeout', params},
    });
  }

  if (name === 'AbortError' || /aborted/i.test(message)) {
    return new ToolError({
      category: 'cancelled',
      isRetryable: false,
      message,
      cause: err,
      context: {toolName, attempted: toolName, failed: 'cancelled', params},
    });
  }

  if (name === 'ProtocolError') {
    return new ToolError({
      category: 'protocol',
      isRetryable: false,
      message,
      cause: err,
      context: {
        toolName,
        attempted: toolName,
        failed: 'CDP protocol error',
        params,
      },
    });
  }

  if (name === 'ZodError' || name === 'ValidationError') {
    return new ToolError({
      category: 'validation',
      isRetryable: false,
      message,
      cause: err,
      context: {
        toolName,
        attempted: toolName,
        failed: 'input validation',
        params,
      },
    });
  }

  if (err instanceof TypeError || err instanceof ReferenceError) {
    return new ToolError({
      category: 'internal',
      isRetryable: false,
      message,
      cause: err,
      context: {
        toolName,
        attempted: toolName,
        failed: `${name || 'internal'} thrown`,
        params,
      },
    });
  }

  return new ToolError({
    category: 'internal',
    isRetryable: false,
    message,
    cause: err,
    context: {
      toolName,
      attempted: toolName,
      failed: 'unclassified error',
      params,
    },
  });
}
