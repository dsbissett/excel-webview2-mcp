import assert from 'node:assert';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import {describe, it, afterEach, beforeEach} from 'node:test';

import {
  assertDaemonIsNotRunning,
  assertDaemonIsRunning,
  runCli,
} from '../utils.js';

describe('excel-webview2', () => {
  beforeEach(async () => {
    await runCli(['stop']);
    await assertDaemonIsNotRunning();
  });

  afterEach(async () => {
    await runCli(['stop']);
    await assertDaemonIsNotRunning();
  });

  it('can start and stop the daemon', async () => {
    await assertDaemonIsNotRunning();

    const startResult = await runCli(['start']);
    assert.strictEqual(
      startResult.status,
      0,
      `start command failed: ${startResult.stderr}`,
    );

    await assertDaemonIsRunning();

    const stopResult = await runCli(['stop']);
    assert.strictEqual(
      stopResult.status,
      0,
      `stop command failed: ${stopResult.stderr}`,
    );

    await assertDaemonIsNotRunning();
  });

  it('can start the daemon with userDataDir', async () => {
    const userDataDir = path.join(
      os.tmpdir(),
      `excel-webview2-test-${crypto.randomUUID()}`,
    );
    fs.mkdirSync(userDataDir, {recursive: true});

    const startResult = await runCli(['start', '--userDataDir', userDataDir]);
    assert.strictEqual(
      startResult.status,
      0,
      `start command failed: ${startResult.stderr}`,
    );
    assert.ok(
      !startResult.stderr.includes(
        'Arguments userDataDir and isolated are mutually exclusive',
      ),
      `unexpected conflict error: ${startResult.stderr}`,
    );

    await assertDaemonIsRunning();
  });
});
