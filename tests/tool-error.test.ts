import assert from 'node:assert';
import {describe, it} from 'node:test';

import {ConnectionError} from '../src/connection/error.js';
import {ToolError, classifyUnknownError} from '../src/tools/ToolError.js';

describe('classifyUnknownError', () => {
  it('returns ToolError as-is', () => {
    const original = new ToolError({
      category: 'not_found',
      isRetryable: false,
      message: 'gone',
      context: {toolName: 't', attempted: 'a', failed: 'f'},
    });
    const result = classifyUnknownError(original, 't');
    assert.strictEqual(result, original);
  });

  it('classifies ConnectionError(timeout) as retryable connection error', () => {
    const err = new ConnectionError({
      url: 'http://x',
      reason: 'timeout',
      hint: '',
    });
    const result = classifyUnknownError(err, 'list_pages');
    assert.strictEqual(result.category, 'connection');
    assert.strictEqual(result.isRetryable, true);
    assert.strictEqual(result.context.toolName, 'list_pages');
  });

  it('classifies ConnectionError(http-error:401) as auth, not retryable', () => {
    const err = new ConnectionError({
      url: 'http://x',
      reason: 'http-error:401',
      hint: '',
    });
    const result = classifyUnknownError(err, 'list_pages');
    assert.strictEqual(result.category, 'auth');
    assert.strictEqual(result.isRetryable, false);
  });

  it('classifies ConnectionError(http-error:429) as rate_limit, retryable', () => {
    const err = new ConnectionError({
      url: 'http://x',
      reason: 'http-error:429',
      hint: '',
    });
    const result = classifyUnknownError(err, 'list_pages');
    assert.strictEqual(result.category, 'rate_limit');
    assert.strictEqual(result.isRetryable, true);
  });

  it('classifies TimeoutError by name as timeout, retryable', () => {
    const err = new Error('Navigation timeout of 30000 ms exceeded');
    err.name = 'TimeoutError';
    const result = classifyUnknownError(err, 'navigate');
    assert.strictEqual(result.category, 'timeout');
    assert.strictEqual(result.isRetryable, true);
  });

  it('classifies AbortError as cancelled, not retryable', () => {
    const err = new Error('aborted');
    err.name = 'AbortError';
    const result = classifyUnknownError(err, 'navigate');
    assert.strictEqual(result.category, 'cancelled');
    assert.strictEqual(result.isRetryable, false);
  });

  it('classifies ProtocolError as protocol, not retryable', () => {
    const err = new Error('Protocol error (Target.foo): bad command');
    err.name = 'ProtocolError';
    const result = classifyUnknownError(err, 'evaluate_script');
    assert.strictEqual(result.category, 'protocol');
    assert.strictEqual(result.isRetryable, false);
  });

  it('classifies TypeError as internal, not retryable', () => {
    const err = new TypeError(
      "Cannot read properties of undefined (reading 'x')",
    );
    const result = classifyUnknownError(err, 'click');
    assert.strictEqual(result.category, 'internal');
    assert.strictEqual(result.isRetryable, false);
  });

  it('classifies unknown errors as internal, not retryable', () => {
    const result = classifyUnknownError(new Error('something weird'), 'foo');
    assert.strictEqual(result.category, 'internal');
    assert.strictEqual(result.isRetryable, false);
    assert.strictEqual(result.context.toolName, 'foo');
  });

  it('captures params in context', () => {
    const result = classifyUnknownError(new Error('boom'), 'foo', {a: 1});
    assert.deepStrictEqual(result.context.params, {a: 1});
  });
});

describe('ToolError.toStructured', () => {
  it('produces the structured envelope', () => {
    const err = new ToolError({
      category: 'not_found',
      isRetryable: false,
      message: 'page not found',
      context: {
        toolName: 'list_pages',
        attempted: 'resolve page',
        failed: 'page id missing',
        details: {pageId: 42},
      },
    });
    assert.deepStrictEqual(err.toStructured(), {
      isError: true,
      errorCategory: 'not_found',
      isRetryable: false,
      context: {
        toolName: 'list_pages',
        attempted: 'resolve page',
        failed: 'page id missing',
        details: {pageId: 42},
      },
    });
  });
});
