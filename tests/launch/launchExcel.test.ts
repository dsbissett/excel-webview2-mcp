import assert from 'node:assert';
import * as child_process from 'node:child_process';
import {ChildProcess} from 'node:child_process';
import {EventEmitter} from 'node:events';
import {PassThrough} from 'node:stream';
import {afterEach, beforeEach, describe, it} from 'node:test';

import sinon from 'sinon';

import {
  createLaunchExcelForTesting,
  LaunchError,
} from '../../src/launch/launchExcel.js';
import type {AddinProject} from '../../src/launch/detectAddin.js';

class FakeProcess extends EventEmitter {
  env: NodeJS.ProcessEnv;
  platform: NodeJS.Platform = 'win32';
  exitCalls: number[] = [];

  constructor(env: NodeJS.ProcessEnv = {}) {
    super();
    this.env = env;
  }

  exit(code = 0): never {
    this.exitCalls.push(code);
    throw new Error(`process.exit(${code})`);
  }
}

class FakeChildProcess extends EventEmitter {
  exitCode: number | null = null;
  killed = false;
  pid?: number;
  signalCode: NodeJS.Signals | null = null;
  stderr = new PassThrough();
  stdout = new PassThrough();

  constructor(pid?: number) {
    super();
    this.pid = pid;
  }

  kill(): boolean {
    this.killed = true;
    return true;
  }

  close(code: number | null = 0, signal: NodeJS.Signals | null = null): void {
    this.exitCode = code;
    this.signalCode = signal;
    this.emit('close', code, signal);
  }

  fail(error: Error): void {
    this.emit('error', error);
  }
}

function makeProject(): AddinProject {
  return {
    root: 'C:\\repo',
    manifestKind: 'xml',
    manifestPath: 'C:\\repo\\manifest.xml',
    packageManager: 'npm',
  };
}

describe('launchExcel', () => {
  let clockMs: number;

  beforeEach(() => {
    clockMs = 0;
  });

  afterEach(() => {
    sinon.restore();
  });

  it('launches with the local office-addin-debugging shim and waits for CDP readiness', async () => {
    const fakeProcess = new FakeProcess({
      PATH: 'C:\\tools',
      PATHEXT: '.EXE;.CMD',
    });
    const launchChild = new FakeChildProcess(4321);
    const spawnStub = sinon
      .stub()
      .returns(launchChild as unknown as ChildProcess);
    const probeStub = sinon
      .stub()
      .onFirstCall()
      .resolves({ok: false, reason: 'unreachable'})
      .onSecondCall()
      .resolves({ok: true, version: 'Edge/136.0.0.0'});
    const accessStub = sinon.stub().callsFake(async filePath => {
      if (
        filePath === 'C:\\repo\\node_modules\\.bin\\office-addin-debugging.cmd'
      ) {
        return undefined;
      }
      throw new Error(`ENOENT: ${filePath}`);
    });
    const sleepCalls: number[] = [];

    const executor = createLaunchExcelForTesting({
      access: accessStub,
      now: () => clockMs,
      probe: probeStub,
      processRef: fakeProcess,
      sleep: async (ms: number) => {
        sleepCalls.push(ms);
        clockMs += ms;
      },
      spawn: spawnStub as unknown as typeof spawnStub,
    });

    const result = await executor.launchExcel({
      project: makeProject(),
      port: 9333,
      extraBrowserArgs: ['--foo=bar', '--baz'],
      timeoutMs: 5000,
    });

    assert.strictEqual(result.pid, 4321);
    assert.strictEqual(result.cdpUrl, 'http://localhost:9333');
    assert.deepStrictEqual(executor.getTrackedPids(), [4321]);
    assert.deepStrictEqual(sleepCalls, [500]);
    assert.strictEqual(spawnStub.callCount, 1);
    assert.deepStrictEqual(spawnStub.firstCall.args.slice(0, 2), [
      'C:\\repo\\node_modules\\.bin\\office-addin-debugging.cmd',
      ['start', 'C:\\repo\\manifest.xml'],
    ]);
    assert.strictEqual(
      spawnStub.firstCall.args[2]?.env?.[
        'WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS'
      ],
      '--remote-debugging-port=9333 --foo=bar --baz',
    );

    executor.reset();
  });

  it('falls back to npx --no-install and stop() uses the same launcher', async () => {
    const fakeProcess = new FakeProcess({
      PATH: 'C:\\tools',
      PATHEXT: '.EXE;.CMD',
    });
    const launchChild = new FakeChildProcess(2468);
    const stopChild = new FakeChildProcess(1357);
    const spawnStub = sinon
      .stub()
      .onFirstCall()
      .returns(launchChild as unknown as ChildProcess)
      .onSecondCall()
      .returns(stopChild as unknown as ChildProcess);
    const accessStub = sinon.stub().callsFake(async filePath => {
      if (filePath === 'C:\\tools\\npx.cmd') {
        return undefined;
      }
      throw new Error(`ENOENT: ${filePath}`);
    });

    const executor = createLaunchExcelForTesting({
      access: accessStub,
      now: () => clockMs,
      probe: async () => ({ok: true, version: 'Edge/136.0.0.0'}),
      processRef: fakeProcess,
      sleep: async (ms: number) => {
        if (ms === 10_000) {
          stopChild.close(0);
          return new Promise(() => {});
        }
        clockMs += ms;
      },
      spawn: spawnStub as unknown as typeof spawnStub,
    });

    const result = await executor.launchExcel({
      project: makeProject(),
      timeoutMs: 2000,
    });
    await result.stop();

    assert.strictEqual(spawnStub.callCount, 2);
    assert.deepStrictEqual(spawnStub.firstCall.args.slice(0, 2), [
      'C:\\tools\\npx.cmd',
      [
        '--no-install',
        'office-addin-debugging',
        'start',
        'C:\\repo\\manifest.xml',
      ],
    ]);
    assert.deepStrictEqual(spawnStub.secondCall.args.slice(0, 2), [
      'C:\\tools\\npx.cmd',
      [
        '--no-install',
        'office-addin-debugging',
        'stop',
        'C:\\repo\\manifest.xml',
      ],
    ]);
    assert.strictEqual(launchChild.killed, true);
    assert.deepStrictEqual(executor.getTrackedPids(), []);

    executor.reset();
  });

  it('refuses to overwrite an existing remote-debugging port env var', async () => {
    const fakeProcess = new FakeProcess({
      WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS:
        '--remote-debugging-port=9555 --foo=bar',
    });
    const spawnStub = sinon.stub();

    const executor = createLaunchExcelForTesting({
      access: sinon.stub().resolves(),
      probe: async () => ({ok: true, version: 'Edge/136.0.0.0'}),
      processRef: fakeProcess,
      sleep: async () => undefined,
      spawn: spawnStub as unknown as typeof spawnStub,
    });

    await assert.rejects(
      () =>
        executor.launchExcel({
          project: makeProject(),
        }),
      error => {
        assert.ok(error instanceof LaunchError);
        assert.strictEqual(error.reason, 'port-already-configured');
        return true;
      },
    );

    assert.strictEqual(spawnStub.callCount, 0);
    executor.reset();
  });

  it('fails with launcher-missing when neither a local shim nor npx is available', async () => {
    const fakeProcess = new FakeProcess({
      PATH: 'C:\\tools',
      PATHEXT: '.EXE;.CMD',
    });
    const executor = createLaunchExcelForTesting({
      access: sinon.stub().rejects(new Error('ENOENT')),
      probe: async () => ({ok: true, version: 'Edge/136.0.0.0'}),
      processRef: fakeProcess,
      sleep: async () => undefined,
      spawn: sinon.stub() as unknown as typeof child_process.spawn,
    });

    await assert.rejects(
      () =>
        executor.launchExcel({
          project: makeProject(),
        }),
      error => {
        assert.ok(error instanceof LaunchError);
        assert.strictEqual(error.reason, 'launcher-missing');
        return true;
      },
    );

    executor.reset();
  });

  it('kills the launch process and reports cdp-not-ready on timeout', async () => {
    const fakeProcess = new FakeProcess();
    const launchChild = new FakeChildProcess(999);
    const spawnStub = sinon
      .stub()
      .returns(launchChild as unknown as ChildProcess);

    const executor = createLaunchExcelForTesting({
      access: sinon.stub().resolves(),
      now: () => clockMs,
      probe: async () => ({ok: false, reason: 'unreachable'}),
      processRef: fakeProcess,
      sleep: async (ms: number) => {
        clockMs += ms;
      },
      spawn: spawnStub as unknown as typeof spawnStub,
    });

    await assert.rejects(
      () =>
        executor.launchExcel({
          project: makeProject(),
          timeoutMs: 1200,
        }),
      error => {
        assert.ok(error instanceof LaunchError);
        assert.strictEqual(error.reason, 'cdp-not-ready');
        return true;
      },
    );

    assert.strictEqual(launchChild.killed, true);
    executor.reset();
  });

  it('refuses to launch on non-Windows platforms', async () => {
    const fakeProcess = new FakeProcess();
    fakeProcess.platform = 'linux';
    const spawnStub = sinon.stub();

    const executor = createLaunchExcelForTesting({
      access: sinon.stub().resolves(),
      probe: async () => ({ok: true, version: 'Edge/136.0.0.0'}),
      processRef: fakeProcess,
      sleep: async () => undefined,
      spawn: spawnStub as unknown as typeof spawnStub,
    });

    await assert.rejects(
      () =>
        executor.launchExcel({
          project: makeProject(),
        }),
      error => {
        assert.ok(error instanceof LaunchError);
        assert.strictEqual(error.reason, 'unsupported-platform');
        return true;
      },
    );

    assert.strictEqual(spawnStub.callCount, 0);
    executor.reset();
  });

  it('aborts the launch when the AbortSignal fires before CDP is ready', async () => {
    const fakeProcess = new FakeProcess();
    const launchChild = new FakeChildProcess(555);
    const spawnStub = sinon
      .stub()
      .returns(launchChild as unknown as ChildProcess);
    const controller = new AbortController();

    const executor = createLaunchExcelForTesting({
      access: sinon.stub().resolves(),
      now: () => clockMs,
      probe: async () => {
        controller.abort();
        return {ok: false, reason: 'unreachable'};
      },
      processRef: fakeProcess,
      sleep: async (ms: number) => {
        clockMs += ms;
      },
      spawn: spawnStub as unknown as typeof spawnStub,
    });

    await assert.rejects(
      () =>
        executor.launchExcel({
          project: makeProject(),
          timeoutMs: 5000,
          signal: controller.signal,
        }),
      error => {
        assert.ok(error instanceof LaunchError);
        assert.strictEqual(error.reason, 'aborted');
        return true;
      },
    );

    assert.strictEqual(launchChild.killed, true);
    executor.reset();
  });

  it('reports launch-failed when the launcher exits before the CDP endpoint is ready', async () => {
    const fakeProcess = new FakeProcess();
    const launchChild = new FakeChildProcess(777);
    const spawnStub = sinon
      .stub()
      .returns(launchChild as unknown as ChildProcess);
    let probeCalls = 0;

    const executor = createLaunchExcelForTesting({
      access: sinon.stub().resolves(),
      now: () => clockMs,
      probe: async () => {
        probeCalls += 1;
        if (probeCalls === 1) {
          queueMicrotask(() => launchChild.close(3));
        }
        return {ok: false, reason: 'unreachable'};
      },
      processRef: fakeProcess,
      sleep: async (ms: number) => {
        clockMs += ms;
      },
      spawn: spawnStub as unknown as typeof spawnStub,
    });

    await assert.rejects(
      () =>
        executor.launchExcel({
          project: makeProject(),
          timeoutMs: 5000,
        }),
      error => {
        assert.ok(error instanceof LaunchError);
        assert.strictEqual(error.reason, 'launch-failed');
        return true;
      },
    );

    executor.reset();
  });

  it('surfaces stop-failed when the stop command exits non-zero', async () => {
    const fakeProcess = new FakeProcess();
    const launchChild = new FakeChildProcess(8080);
    const stopChild = new FakeChildProcess(8081);
    const spawnStub = sinon
      .stub()
      .onFirstCall()
      .returns(launchChild as unknown as ChildProcess)
      .onSecondCall()
      .returns(stopChild as unknown as ChildProcess);

    const executor = createLaunchExcelForTesting({
      access: sinon.stub().resolves(),
      now: () => clockMs,
      probe: async () => ({ok: true, version: 'Edge/136.0.0.0'}),
      processRef: fakeProcess,
      sleep: async (ms: number) => {
        if (ms === 10_000) {
          return new Promise(() => {});
        }
        clockMs += ms;
      },
      spawn: spawnStub as unknown as typeof spawnStub,
    });

    const result = await executor.launchExcel({
      project: makeProject(),
      timeoutMs: 2000,
    });

    queueMicrotask(() => stopChild.close(1));

    await assert.rejects(
      () => result.stop(),
      error => {
        assert.ok(error instanceof LaunchError);
        assert.strictEqual(error.reason, 'stop-failed');
        return true;
      },
    );

    executor.reset();
  });
});
