import assert from 'node:assert';
import {afterEach, describe, it} from 'node:test';

import sinon from 'sinon';

import type {AddinProject} from '../../src/launch/detectAddin.js';
import {LaunchError} from '../../src/launch/launchExcel.js';
import {runAutoLaunch} from '../../src/launch/runAutoLaunch.js';
import {
  excelDetectAddin,
  excelLaunchAddin,
  excelStopAddin,
} from '../../src/tools/lifecycle.js';
import {
  resetLifecycleDepsForTesting,
  setLifecycleDepsForTesting,
} from '../../src/tools/lifecycleState.js';

interface CapturedResponse {
  lines: string[];
  structured?: object;
}

function makeResponse(): {
  response: {
    appendResponseLine: (value: string) => void;
    setStructuredContent: (value: object) => void;
  };
  captured: CapturedResponse;
} {
  const captured: CapturedResponse = {lines: []};
  const response = {
    appendResponseLine(value: string) {
      captured.lines.push(value);
    },
    setStructuredContent(value: object) {
      captured.structured = value;
    },
  };
  return {response, captured};
}

function project(manifest = 'C:\\repo\\manifest.xml'): AddinProject {
  return {
    root: 'C:\\repo',
    manifestKind: 'xml',
    manifestPath: manifest,
    packageManager: 'npm',
  };
}

describe('lifecycle tools', () => {
  afterEach(() => {
    resetLifecycleDepsForTesting();
    sinon.restore();
  });

  describe('excel_detect_addin', () => {
    it('returns the detected project as structured content', async () => {
      const detected = project();
      setLifecycleDepsForTesting({
        detectExcelAddin: async () => detected,
        cwd: () => 'C:\\repo',
      });

      const {response, captured} = makeResponse();
      const tool = excelDetectAddin;
      await tool.handler({params: {}}, response as never, {} as never);

      assert.deepStrictEqual(captured.structured, {
        detected: true,
        project: detected,
      });
    });

    it('reports no detection when the directory is not an add-in', async () => {
      setLifecycleDepsForTesting({
        detectExcelAddin: async () => null,
        cwd: () => 'C:\\other',
      });

      const {response, captured} = makeResponse();
      await excelDetectAddin.handler(
        {params: {cwd: 'C:\\other'}},
        response as never,
        {} as never,
      );

      assert.deepStrictEqual(captured.structured, {
        detected: false,
        cwd: 'C:\\other',
      });
      assert.match(captured.lines[0] ?? '', /No Excel add-in project/);
    });
  });

  describe('excel_launch_addin', () => {
    it('launches and auto-connects on first invocation, then reuses on second', async () => {
      const detected = project();
      const launchStub = sinon.stub().resolves({
        pid: 123,
        cdpUrl: 'http://localhost:9222',
        stop: sinon.stub().resolves(),
      });
      const connectStub = sinon.stub().resolves({connected: true});
      setLifecycleDepsForTesting({
        detectExcelAddin: async () => detected,
        launchExcel: launchStub,
        ensureBrowserConnected: connectStub as never,
      });

      const tool = excelLaunchAddin();

      const first = makeResponse();
      await tool.handler({params: {}}, first.response as never, {} as never);
      assert.strictEqual(launchStub.callCount, 1);
      assert.strictEqual(connectStub.callCount, 1);
      assert.deepStrictEqual(first.captured.structured, {
        reused: false,
        pid: 123,
        cdpUrl: 'http://localhost:9222',
        project: detected,
      });

      const second = makeResponse();
      await tool.handler({params: {}}, second.response as never, {} as never);
      assert.strictEqual(launchStub.callCount, 1);
      assert.strictEqual(connectStub.callCount, 2);
      assert.deepStrictEqual(second.captured.structured, {
        reused: true,
        pid: 123,
        cdpUrl: 'http://localhost:9222',
        project: detected,
      });
    });

    it('skips connect when autoConnect is false', async () => {
      const detected = project();
      const connectStub = sinon.stub().resolves({connected: true});
      setLifecycleDepsForTesting({
        detectExcelAddin: async () => detected,
        launchExcel: async () => ({
          pid: 1,
          cdpUrl: 'http://localhost:9222',
          stop: async () => undefined,
        }),
        ensureBrowserConnected: connectStub as never,
      });

      const tool = excelLaunchAddin();
      const {response} = makeResponse();
      await tool.handler(
        {params: {autoConnect: false}},
        response as never,
        {} as never,
      );

      assert.strictEqual(connectStub.callCount, 0);
    });

    it('surfaces LaunchError reason without throwing', async () => {
      const detected = project();
      setLifecycleDepsForTesting({
        detectExcelAddin: async () => detected,
        launchExcel: async () => {
          throw new LaunchError('cdp-not-ready', 'timed out', {
            output: ['line-1'],
          });
        },
      });

      const tool = excelLaunchAddin();
      const {response, captured} = makeResponse();
      await tool.handler({params: {}}, response as never, {} as never);

      assert.ok(
        captured.lines.some(line => line.includes('cdp-not-ready')),
        `expected cdp-not-ready in output; got ${captured.lines.join('\n')}`,
      );
      assert.ok(captured.lines.includes('line-1'));
    });

    it('uses CLI-provided defaults when no params are supplied', async () => {
      const detected = project();
      const launchStub = sinon.stub().resolves({
        pid: 77,
        cdpUrl: 'http://localhost:9444',
        stop: async () => undefined,
      });
      setLifecycleDepsForTesting({
        detectExcelAddin: async () => detected,
        launchExcel: launchStub,
        ensureBrowserConnected: (async () => ({})) as never,
      });

      const tool = excelLaunchAddin({
        launchPort: 9444,
        launchTimeout: 1234,
      } as never);
      const {response} = makeResponse();
      await tool.handler({params: {}}, response as never, {} as never);

      const callArgs = launchStub.firstCall.args[0];
      assert.strictEqual(callArgs.port, 9444);
      assert.strictEqual(callArgs.timeoutMs, 1234);
    });
  });

  describe('excel_stop_addin', () => {
    it('stops the tracked launch and removes it from the map', async () => {
      const detected = project();
      const stopStub = sinon.stub().resolves();
      setLifecycleDepsForTesting({
        detectExcelAddin: async () => detected,
        launchExcel: async () => ({
          pid: 42,
          cdpUrl: 'http://localhost:9222',
          stop: stopStub,
        }),
        ensureBrowserConnected: (async () => ({})) as never,
      });

      const launchTool = excelLaunchAddin();
      const firstLaunch = makeResponse();
      await launchTool.handler(
        {params: {autoConnect: false}},
        firstLaunch.response as never,
        {} as never,
      );

      const stopResponse = makeResponse();
      await excelStopAddin.handler(
        {params: {}},
        stopResponse.response as never,
        {} as never,
      );

      assert.strictEqual(stopStub.callCount, 1);
      assert.ok(
        stopResponse.captured.lines.some(line =>
          line.includes('Stopped launch'),
        ),
      );

      // Second stop finds nothing tracked.
      const noop = makeResponse();
      await excelStopAddin.handler(
        {params: {}},
        noop.response as never,
        {} as never,
      );
      assert.ok(noop.captured.lines.some(line => line.includes('No tracked')));
    });

    it('reports missing manifest without stopping anything', async () => {
      setLifecycleDepsForTesting({});
      const {response, captured} = makeResponse();
      await excelStopAddin.handler(
        {params: {manifestPath: 'C:\\nope\\manifest.xml'}},
        response as never,
        {} as never,
      );
      assert.ok(
        captured.lines.some(line => line.includes('No tracked launch for')),
      );
    });
  });

  describe('runAutoLaunch', () => {
    it('is a no-op when no add-in is detected', async () => {
      const launchStub = sinon.stub();
      setLifecycleDepsForTesting({
        detectExcelAddin: async () => null,
        launchExcel: launchStub as never,
      });
      const logs: string[] = [];
      await runAutoLaunch({cwd: 'C:\\other', logger: m => logs.push(m)});
      assert.strictEqual(launchStub.callCount, 0);
      assert.ok(logs.some(l => l.includes('no Excel add-in detected')));
    });

    it('launches once and records the tracked entry', async () => {
      const detected = project();
      const launchStub = sinon.stub().resolves({
        pid: 9,
        cdpUrl: 'http://localhost:9222',
        stop: async () => undefined,
      });
      setLifecycleDepsForTesting({
        detectExcelAddin: async () => detected,
        launchExcel: launchStub,
      });
      await runAutoLaunch({cwd: detected.root});
      assert.strictEqual(launchStub.callCount, 1);

      // Second call is idempotent.
      await runAutoLaunch({cwd: detected.root});
      assert.strictEqual(launchStub.callCount, 1);
    });
  });
});
