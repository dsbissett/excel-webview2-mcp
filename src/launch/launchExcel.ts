import {spawn, type ChildProcessWithoutNullStreams} from 'node:child_process';
import fs from 'node:fs/promises';
import net from 'node:net';
import path from 'node:path';

import {probeCdpEndpoint, type ProbeResult} from '../connection/probe.js';

import type {AddinProject} from './detectAddin.js';

export interface LaunchOptions {
  project: AddinProject;
  port?: number;
  extraBrowserArgs?: string[];
  timeoutMs?: number;
  signal?: AbortSignal;
  skipDevServer?: boolean;
  devServerTimeoutMs?: number;
}

export interface LaunchResult {
  pid: number;
  cdpUrl: string;
  stop: () => Promise<void>;
}

export type LaunchErrorReason =
  | 'unsupported-platform'
  | 'launcher-missing'
  | 'port-already-configured'
  | 'launch-failed'
  | 'cdp-not-ready'
  | 'dev-server-not-ready'
  | 'dev-server-failed'
  | 'stop-failed'
  | 'aborted';

export class LaunchError extends Error {
  readonly reason: LaunchErrorReason;
  readonly output: string[];

  constructor(
    reason: LaunchErrorReason,
    message: string,
    options?: {
      cause?: unknown;
      output?: string[];
    },
  ) {
    super(message, {cause: options?.cause});
    this.name = 'LaunchError';
    this.reason = reason;
    this.output = options?.output ?? [];
  }
}

interface LauncherCommand {
  command: string;
  argsPrefix: string[];
}

type LaunchChildProcess = ChildProcessWithoutNullStreams;

interface ChildOutcome {
  type: 'close' | 'error';
  code?: number | null;
  signal?: NodeJS.Signals | null;
  error?: unknown;
}

interface TrackedLaunch {
  child: LaunchChildProcess;
  cdpUrl: string;
  launcher: LauncherCommand;
  output: string[];
  pid: number;
  project: AddinProject;
  stopPromise?: Promise<void>;
  devServer?: DevServerHandle;
}

interface DevServerHandle {
  child: LaunchChildProcess;
  output: string[];
  port: number;
  preexisting: boolean;
}

interface CleanupProcess {
  env: NodeJS.ProcessEnv;
  on: (event: string, listener: (...args: unknown[]) => void) => void;
  off: (event: string, listener: (...args: unknown[]) => void) => void;
  exit: (code?: number) => never;
  platform: NodeJS.Platform;
}

interface LaunchExcelDeps {
  access: typeof fs.access;
  now: () => number;
  probe: (url: string, timeoutMs: number) => Promise<ProbeResult>;
  processRef: CleanupProcess;
  sleep: (ms: number) => Promise<void>;
  spawn: typeof spawn;
}

interface LaunchRuntimeState {
  cleanupRegistered: boolean;
  exitHandler?: (...args: unknown[]) => void;
  sigintHandler?: (...args: unknown[]) => void;
  sigtermHandler?: (...args: unknown[]) => void;
  tracked: Map<number, TrackedLaunch>;
}

interface LaunchExcelTestingApi {
  launchExcel: (options: LaunchOptions) => Promise<LaunchResult>;
  getTrackedPids: () => number[];
  reset: () => void;
}

const DEFAULT_CDP_PORT = 9222;
const DEFAULT_TIMEOUT_MS = 60_000;
const DEFAULT_DEV_SERVER_TIMEOUT_MS = 90_000;
const PROBE_INTERVAL_MS = 500;
const PROBE_TIMEOUT_MS = 1000;
const DEV_SERVER_PROBE_TIMEOUT_MS = 1500;
const STOP_TIMEOUT_MS = 10_000;
const MAX_OUTPUT_LINES = 200;
const LAUNCHER_NAME = 'office-addin-debugging';

const defaultState = createLaunchRuntimeState();

const defaultDeps: LaunchExcelDeps = {
  access: fs.access,
  now: Date.now,
  probe: probeCdpEndpoint,
  processRef: process as CleanupProcess,
  sleep: (ms: number) => new Promise(resolve => setTimeout(resolve, ms)),
  spawn,
};

export const launchExcel = createLaunchExcelImpl(defaultDeps, defaultState);

export function createLaunchExcelForTesting(
  overrides: Partial<LaunchExcelDeps> = {},
): LaunchExcelTestingApi {
  const state = createLaunchRuntimeState();
  const deps: LaunchExcelDeps = {
    ...defaultDeps,
    ...overrides,
  };

  return {
    launchExcel: createLaunchExcelImpl(deps, state),
    getTrackedPids: () => {
      return [...state.tracked.keys()];
    },
    reset: () => {
      resetRuntimeState(state, deps.processRef);
    },
  };
}

export function resetLaunchStateForTesting(): void {
  resetRuntimeState(defaultState, defaultDeps.processRef);
}

function createLaunchExcelImpl(
  deps: LaunchExcelDeps,
  state: LaunchRuntimeState,
) {
  return async function launchExcelImpl(
    options: LaunchOptions,
  ): Promise<LaunchResult> {
    assertWindowsOnly(deps.processRef.platform);

    const port = options.port ?? DEFAULT_CDP_PORT;
    const timeoutMs = options.timeoutMs ?? DEFAULT_TIMEOUT_MS;
    const devServerTimeoutMs =
      options.devServerTimeoutMs ?? DEFAULT_DEV_SERVER_TIMEOUT_MS;
    const cdpUrl = `http://localhost:${port}`;
    const launcher = await resolveLauncher(options.project.root, deps);
    const env = buildLaunchEnv({
      env: deps.processRef.env,
      port,
      extraBrowserArgs: options.extraBrowserArgs,
      projectRoot: options.project.root,
    });

    const devServer = options.skipDevServer
      ? undefined
      : await ensureDevServerRunning({
          deps,
          project: options.project,
          env,
          timeoutMs: devServerTimeoutMs,
          signal: options.signal,
        });

    let launchChild: LaunchChildProcess;
    try {
      launchChild = spawnLauncher(
        deps,
        launcher,
        ['start', options.project.manifestPath],
        {
          cwd: options.project.root,
          env,
          stdio: 'pipe',
          windowsHide: false,
        },
      );
    } catch (error) {
      if (devServer && !devServer.preexisting) {
        killChild(devServer.child);
      }
      throw new LaunchError(
        'launch-failed',
        `Failed to spawn ${LAUNCHER_NAME} (${launcher.command}): ${(error as Error).message}`,
        {cause: error},
      );
    }

    const pid = launchChild.pid;
    if (!pid || pid <= 0) {
      throw new LaunchError(
        'launch-failed',
        `Failed to launch ${LAUNCHER_NAME}; child process did not report a pid.`,
      );
    }

    const output = createOutputBuffer(launchChild);
    const childOutcome = monitorChild(launchChild);

    try {
      await waitForCdpReady({
        childOutcome,
        cdpUrl,
        now: deps.now,
        probe: deps.probe,
        signal: options.signal,
        sleep: deps.sleep,
        timeoutMs,
        output,
      });
    } catch (error) {
      killChild(launchChild);
      if (devServer && !devServer.preexisting) {
        killChild(devServer.child);
      }
      throw error;
    }

    ensureCleanupHandlersRegistered(state, deps);

    const trackedLaunch: TrackedLaunch = {
      child: launchChild,
      cdpUrl,
      launcher,
      output,
      pid,
      project: options.project,
      devServer,
    };
    state.tracked.set(pid, trackedLaunch);
    void childOutcome.finally(() => {
      state.tracked.delete(pid);
    });

    return {
      pid,
      cdpUrl,
      stop: () => stopTrackedLaunch(trackedLaunch, state, deps),
    };
  };
}

function createLaunchRuntimeState(): LaunchRuntimeState {
  return {
    cleanupRegistered: false,
    tracked: new Map(),
  };
}

function resetRuntimeState(
  state: LaunchRuntimeState,
  processRef: CleanupProcess,
): void {
  if (state.exitHandler) {
    processRef.off('exit', state.exitHandler);
  }
  if (state.sigintHandler) {
    processRef.off('SIGINT', state.sigintHandler);
  }
  if (state.sigtermHandler) {
    processRef.off('SIGTERM', state.sigtermHandler);
  }

  for (const trackedLaunch of state.tracked.values()) {
    killChild(trackedLaunch.child);
    if (trackedLaunch.devServer && !trackedLaunch.devServer.preexisting) {
      killChild(trackedLaunch.devServer.child);
    }
  }

  state.tracked.clear();
  state.cleanupRegistered = false;
  state.exitHandler = undefined;
  state.sigintHandler = undefined;
  state.sigtermHandler = undefined;
}

function assertWindowsOnly(platform: NodeJS.Platform): void {
  if (platform !== 'win32') {
    throw new LaunchError(
      'unsupported-platform',
      'Launching Excel add-ins is only supported on Windows because WebView2 is Windows-only.',
    );
  }
}

async function resolveLauncher(
  root: string,
  deps: LaunchExcelDeps,
): Promise<LauncherCommand> {
  for (const candidate of localLauncherCandidates(root)) {
    if (await pathExists(candidate, deps.access)) {
      return {
        command: candidate,
        argsPrefix: [],
      };
    }
  }

  const npxPath = await resolveCommandOnPath('npx', deps);
  if (npxPath) {
    return {
      command: npxPath,
      argsPrefix: ['--no-install', LAUNCHER_NAME],
    };
  }

  throw new LaunchError(
    'launcher-missing',
    `Could not find ${LAUNCHER_NAME}. Install it as a devDependency in ${root} or make npx available on PATH.`,
  );
}

function localLauncherCandidates(root: string): string[] {
  const binRoot = path.join(root, 'node_modules', '.bin');
  return [
    path.join(binRoot, `${LAUNCHER_NAME}.cmd`),
    path.join(binRoot, LAUNCHER_NAME),
    path.join(binRoot, `${LAUNCHER_NAME}.exe`),
  ];
}

async function resolveCommandOnPath(
  command: string,
  deps: LaunchExcelDeps,
): Promise<string | null> {
  const rawPath = deps.processRef.env['PATH'] ?? deps.processRef.env['Path'];
  if (!rawPath) {
    return null;
  }

  const pathEntries = rawPath
    .split(path.delimiter)
    .map(entry => {
      return entry.trim();
    })
    .filter(Boolean);
  const pathExts = (deps.processRef.env['PATHEXT'] ?? '.EXE;.CMD;.BAT;.COM')
    .split(';')
    .map(ext => {
      return ext.toLowerCase();
    })
    .filter(Boolean);

  for (const entry of pathEntries) {
    const baseCandidate = path.join(entry, command);
    for (const candidate of [
      baseCandidate,
      ...pathExts.map(ext => `${baseCandidate}${ext}`),
    ]) {
      if (await pathExists(candidate, deps.access)) {
        return candidate;
      }
    }
  }

  return null;
}

function buildLaunchEnv(options: {
  env: NodeJS.ProcessEnv;
  port: number;
  extraBrowserArgs?: string[];
  projectRoot: string;
}): NodeJS.ProcessEnv {
  const existingArgs = options.env['WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS'];
  if (
    typeof existingArgs === 'string' &&
    /(?:^|\s)--remote-debugging-port(?:\s|=|$)/.test(existingArgs)
  ) {
    throw new LaunchError(
      'port-already-configured',
      'WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS already contains --remote-debugging-port; unset it before launching Excel from the MCP server.',
    );
  }

  const browserArgs = [
    `--remote-debugging-port=${options.port}`,
    ...(options.extraBrowserArgs ?? []).filter(Boolean),
  ].join(' ');

  const binDir = path.join(options.projectRoot, 'node_modules', '.bin');
  const pathKey = 'PATH' in options.env ? 'PATH' : 'Path';
  const existingPath = options.env[pathKey] ?? '';
  const augmentedPath = existingPath
    ? `${binDir}${path.delimiter}${existingPath}`
    : binDir;

  return {
    ...options.env,
    [pathKey]: augmentedPath,
    WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS: browserArgs.trim(),
  };
}

function spawnLauncher(
  deps: LaunchExcelDeps,
  launcher: LauncherCommand,
  trailingArgs: string[],
  options: {
    cwd: string;
    env: NodeJS.ProcessEnv;
    stdio: 'pipe';
    windowsHide: boolean;
  },
): LaunchChildProcess {
  const args = [...launcher.argsPrefix, ...trailingArgs];
  const isWindowsBatch =
    deps.processRef.platform === 'win32' &&
    /\.(cmd|bat)$/i.test(launcher.command);

  if (isWindowsBatch) {
    return deps.spawn(`"${launcher.command}"`, args.map(quoteWindowsShellArg), {
      ...options,
      shell: true,
    }) as LaunchChildProcess;
  }

  return deps.spawn(launcher.command, args, options) as LaunchChildProcess;
}

function quoteWindowsShellArg(arg: string): string {
  if (arg === '') {
    return '""';
  }
  if (!/[\s"&|<>^%]/.test(arg)) {
    return arg;
  }
  return `"${arg.replace(/"/g, '\\"')}"`;
}

function createOutputBuffer(child: LaunchChildProcess): string[] {
  const output: string[] = [];

  const append = (chunk: string) => {
    const lines = chunk
      .split(/\r?\n/)
      .map(line => {
        return line.trimEnd();
      })
      .filter(Boolean);

    for (const line of lines) {
      output.push(line);
    }
    if (output.length > MAX_OUTPUT_LINES) {
      output.splice(0, output.length - MAX_OUTPUT_LINES);
    }
  };

  child.stdout?.on('data', chunk => {
    append(String(chunk));
  });
  child.stderr?.on('data', chunk => {
    append(String(chunk));
  });

  return output;
}

function monitorChild(child: LaunchChildProcess): Promise<ChildOutcome> {
  return new Promise(resolve => {
    child.once('error', error => {
      resolve({
        type: 'error',
        error,
      });
    });
    child.once('close', (code, signal) => {
      resolve({
        type: 'close',
        code,
        signal,
      });
    });
  });
}

async function waitForCdpReady(options: {
  childOutcome: Promise<ChildOutcome>;
  cdpUrl: string;
  now: () => number;
  output: string[];
  probe: (url: string, timeoutMs: number) => Promise<ProbeResult>;
  signal?: AbortSignal;
  sleep: (ms: number) => Promise<void>;
  timeoutMs: number;
}): Promise<void> {
  const deadline = options.now() + options.timeoutMs;
  let lastProbeReason: string | undefined;
  let outcome: ChildOutcome | undefined;

  void options.childOutcome.then(result => {
    outcome = result;
  });

  while (options.now() <= deadline) {
    throwIfAborted(options.signal, options.output);

    if (outcome) {
      throw childFailureToError(outcome, options.output);
    }

    const remaining = deadline - options.now();
    const probeTimeoutMs = Math.max(
      1,
      Math.min(PROBE_TIMEOUT_MS, remaining || PROBE_TIMEOUT_MS),
    );
    const probe = await options.probe(options.cdpUrl, probeTimeoutMs);
    if (probe.ok) {
      return;
    }
    lastProbeReason = probe.reason;

    if (options.now() >= deadline) {
      break;
    }

    await options.sleep(Math.min(PROBE_INTERVAL_MS, deadline - options.now()));
  }

  throw new LaunchError(
    'cdp-not-ready',
    `Timed out waiting for the Excel WebView2 CDP endpoint at ${options.cdpUrl} to become ready.${lastProbeReason ? ` Last probe result: ${lastProbeReason}.` : ''}`,
    {output: options.output},
  );
}

async function ensureDevServerRunning(args: {
  deps: LaunchExcelDeps;
  project: AddinProject;
  env: NodeJS.ProcessEnv;
  timeoutMs: number;
  signal?: AbortSignal;
}): Promise<DevServerHandle | undefined> {
  const devServer = args.project.devServer;
  if (!devServer) {
    return undefined;
  }

  const url = `http://localhost:${devServer.port}`;

  if (await isPortListening(devServer.port, DEV_SERVER_PROBE_TIMEOUT_MS)) {
    return {
      child: undefined as unknown as LaunchChildProcess,
      output: [],
      port: devServer.port,
      preexisting: true,
    };
  }

  const runner = resolvePackageRunner(args.project.packageManager);
  const output: string[] = [];
  let child: LaunchChildProcess;
  try {
    child = spawnPackageScript(args.deps, runner, devServer.script, {
      cwd: args.project.root,
      env: args.env,
      stdio: 'pipe',
      windowsHide: false,
    });
  } catch (error) {
    if (await isPortListening(devServer.port, DEV_SERVER_PROBE_TIMEOUT_MS)) {
      return {
        child: undefined as unknown as LaunchChildProcess,
        output: [],
        port: devServer.port,
        preexisting: true,
      };
    }
    throw new LaunchError(
      'dev-server-failed',
      `Failed to spawn dev server (${runner} run ${devServer.script}): ${(error as Error).message}. If the dev server is already running elsewhere, re-invoke excel_launch_addin with skipDevServer: true.`,
      {cause: error},
    );
  }

  attachOutputBuffer(child, output);
  const childOutcome = monitorChild(child);

  const deadline = args.deps.now() + args.timeoutMs;
  let outcome: ChildOutcome | undefined;
  void childOutcome.then(result => {
    outcome = result;
  });

  while (args.deps.now() <= deadline) {
    if (args.signal?.aborted) {
      killChild(child);
      throw new LaunchError('aborted', 'Excel add-in launch was aborted.', {
        output,
      });
    }
    if (outcome) {
      throw new LaunchError(
        'dev-server-failed',
        `Dev server script '${devServer.script}' exited before ${url} became ready.`,
        {output},
      );
    }
    if (await isPortListening(devServer.port, DEV_SERVER_PROBE_TIMEOUT_MS)) {
      return {child, output, port: devServer.port, preexisting: false};
    }
    await args.deps.sleep(
      Math.min(PROBE_INTERVAL_MS, Math.max(1, deadline - args.deps.now())),
    );
  }

  killChild(child);
  throw new LaunchError(
    'dev-server-not-ready',
    `Timed out waiting for dev server at ${url} (script '${devServer.script}').`,
    {output},
  );
}

async function isPortListening(
  port: number,
  timeoutMs: number,
): Promise<boolean> {
  return new Promise(resolve => {
    const socket = net.createConnection({host: '127.0.0.1', port});
    let settled = false;
    const finish = (listening: boolean) => {
      if (settled) return;
      settled = true;
      socket.destroy();
      resolve(listening);
    };
    socket.setTimeout(timeoutMs);
    socket.once('connect', () => finish(true));
    socket.once('timeout', () => finish(false));
    socket.once('error', () => finish(false));
  });
}

function resolvePackageRunner(
  packageManager: AddinProject['packageManager'],
): string {
  return packageManager;
}

function spawnPackageScript(
  deps: LaunchExcelDeps,
  runner: string,
  script: string,
  options: {
    cwd: string;
    env: NodeJS.ProcessEnv;
    stdio: 'pipe';
    windowsHide: boolean;
  },
): LaunchChildProcess {
  const isWindows = deps.processRef.platform === 'win32';
  const command = isWindows ? `${runner}.cmd` : runner;
  const args = ['run', script];

  if (isWindows) {
    return deps.spawn(`"${command}"`, args.map(quoteWindowsShellArg), {
      ...options,
      shell: true,
    }) as LaunchChildProcess;
  }

  return deps.spawn(command, args, options) as LaunchChildProcess;
}

function attachOutputBuffer(child: LaunchChildProcess, output: string[]): void {
  const append = (chunk: string) => {
    const lines = chunk
      .split(/\r?\n/)
      .map(line => line.trimEnd())
      .filter(Boolean);
    for (const line of lines) {
      output.push(line);
    }
    if (output.length > MAX_OUTPUT_LINES) {
      output.splice(0, output.length - MAX_OUTPUT_LINES);
    }
  };
  child.stdout?.on('data', chunk => append(String(chunk)));
  child.stderr?.on('data', chunk => append(String(chunk)));
}

function throwIfAborted(
  signal: AbortSignal | undefined,
  output: string[],
): void {
  if (signal?.aborted) {
    throw new LaunchError('aborted', 'Excel add-in launch was aborted.', {
      output,
    });
  }
}

function childFailureToError(
  outcome: ChildOutcome,
  output: string[],
): LaunchError {
  if (outcome.type === 'error') {
    return new LaunchError(
      'launch-failed',
      `The ${LAUNCHER_NAME} process failed to start.`,
      {
        cause: outcome.error,
        output,
      },
    );
  }

  return new LaunchError(
    'launch-failed',
    `${LAUNCHER_NAME} exited before the CDP endpoint was ready (code=${outcome.code ?? 'null'}, signal=${outcome.signal ?? 'null'}).`,
    {output},
  );
}

async function stopTrackedLaunch(
  trackedLaunch: TrackedLaunch,
  state: LaunchRuntimeState,
  deps: LaunchExcelDeps,
): Promise<void> {
  if (!trackedLaunch.stopPromise) {
    trackedLaunch.stopPromise = (async () => {
      try {
        await runLauncherCommand(
          trackedLaunch.launcher,
          'stop',
          trackedLaunch.project,
          deps,
          trackedLaunch.output,
          STOP_TIMEOUT_MS,
        );
      } finally {
        state.tracked.delete(trackedLaunch.pid);
        killChild(trackedLaunch.child);
        if (trackedLaunch.devServer && !trackedLaunch.devServer.preexisting) {
          killChild(trackedLaunch.devServer.child);
        }
      }
    })();
  }

  return trackedLaunch.stopPromise;
}

async function runLauncherCommand(
  launcher: LauncherCommand,
  action: 'start' | 'stop',
  project: AddinProject,
  deps: LaunchExcelDeps,
  output: string[],
  timeoutMs?: number,
): Promise<void> {
  let child: LaunchChildProcess;
  try {
    child = spawnLauncher(deps, launcher, [action, project.manifestPath], {
      cwd: project.root,
      env: deps.processRef.env,
      stdio: 'pipe',
      windowsHide: false,
    });
  } catch (error) {
    throw new LaunchError(
      'stop-failed',
      `Failed to spawn ${LAUNCHER_NAME} ${action} (${launcher.command}): ${(error as Error).message}`,
      {cause: error, output},
    );
  }

  const commandOutput = createOutputBuffer(child);
  const completion = monitorChild(child);

  if (timeoutMs === undefined) {
    const outcome = await completion;
    mergeOutput(output, commandOutput);
    if (outcome.type !== 'close' || outcome.code !== 0) {
      throw stopFailureFromOutcome(outcome, output);
    }
    return;
  }

  const timedResult = await Promise.race([
    completion.then(outcome => {
      return {
        type: 'outcome' as const,
        outcome,
      };
    }),
    deps.sleep(timeoutMs).then(() => {
      return {
        type: 'timeout' as const,
      };
    }),
  ]);

  mergeOutput(output, commandOutput);

  if (timedResult.type === 'timeout') {
    killChild(child);
    throw new LaunchError(
      'stop-failed',
      `Timed out waiting for ${LAUNCHER_NAME} ${action} to finish.`,
      {output},
    );
  }

  if (timedResult.outcome.type !== 'close' || timedResult.outcome.code !== 0) {
    throw stopFailureFromOutcome(timedResult.outcome, output);
  }
}

function stopFailureFromOutcome(
  outcome: ChildOutcome,
  output: string[],
): LaunchError {
  if (outcome.type === 'error') {
    return new LaunchError(
      'stop-failed',
      `The ${LAUNCHER_NAME} stop command failed to start.`,
      {
        cause: outcome.error,
        output,
      },
    );
  }

  return new LaunchError(
    'stop-failed',
    `${LAUNCHER_NAME} stop exited unsuccessfully (code=${outcome.code ?? 'null'}, signal=${outcome.signal ?? 'null'}).`,
    {output},
  );
}

function mergeOutput(target: string[], source: string[]): void {
  for (const line of source) {
    target.push(line);
  }
  if (target.length > MAX_OUTPUT_LINES) {
    target.splice(0, target.length - MAX_OUTPUT_LINES);
  }
}

function killChild(child: LaunchChildProcess): void {
  if (child.killed || child.exitCode !== null) {
    return;
  }

  try {
    if (process.platform === 'win32' && child.pid) {
      spawn('taskkill', ['/pid', String(child.pid), '/T', '/F'], {
        stdio: 'ignore',
        windowsHide: true,
      });
      return;
    }
    child.kill();
  } catch {
    // Best-effort cleanup only.
  }
}

function ensureCleanupHandlersRegistered(
  state: LaunchRuntimeState,
  deps: LaunchExcelDeps,
): void {
  if (state.cleanupRegistered) {
    return;
  }

  state.exitHandler = () => {
    for (const trackedLaunch of state.tracked.values()) {
      killChild(trackedLaunch.child);
      if (trackedLaunch.devServer && !trackedLaunch.devServer.preexisting) {
        killChild(trackedLaunch.devServer.child);
      }
    }
  };
  state.sigintHandler = () => {
    void stopAllTrackedLaunches(state, deps).finally(() => {
      deps.processRef.exit(0);
    });
  };
  state.sigtermHandler = () => {
    void stopAllTrackedLaunches(state, deps).finally(() => {
      deps.processRef.exit(0);
    });
  };

  deps.processRef.on('exit', state.exitHandler);
  deps.processRef.on('SIGINT', state.sigintHandler);
  deps.processRef.on('SIGTERM', state.sigtermHandler);
  state.cleanupRegistered = true;
}

async function stopAllTrackedLaunches(
  state: LaunchRuntimeState,
  deps: LaunchExcelDeps,
): Promise<void> {
  await Promise.allSettled(
    [...state.tracked.values()].map(trackedLaunch => {
      return stopTrackedLaunch(trackedLaunch, state, deps);
    }),
  );
}

async function pathExists(
  filePath: string,
  access: typeof fs.access,
): Promise<boolean> {
  try {
    await access(filePath);
    return true;
  } catch {
    return false;
  }
}
