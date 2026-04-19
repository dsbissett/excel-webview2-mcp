/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */

import {execSync} from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

import {connectWithRetry} from './connection/retry.js';
import {
  beginReconnect,
  getStickyError,
  isSessionStale,
  markDisconnected,
} from './connection/session.js';
import type {ConnectionEndpointSource} from './connection/status.js';
import {
  markConnectionAttached,
  markConnectionDetached,
  recordProbeResult,
  setConnectionEndpoint,
} from './connection/status.js';
import {logger} from './logger.js';
import type {
  Browser,
  ChromeReleaseChannel,
  LaunchOptions,
  Target,
} from './third_party/index.js';
import {puppeteer} from './third_party/index.js';

let browser: Browser | undefined;
let activeDisconnectListener: (() => void) | undefined;

export const WEBVIEW2_DEBUG_URL = 'http://localhost:9222';

function detachDisconnectListener(): void {
  activeDisconnectListener?.();
  activeDisconnectListener = undefined;
}

function attachDisconnectListener(b: Browser): void {
  detachDisconnectListener();
  const listener = () => {
    markConnectionDetached();
    markDisconnected();
  };
  b.on('disconnected', listener);
  activeDisconnectListener = () => {
    b.off('disconnected', listener);
  };
}

function makeTargetFilter(enableExtensions = false) {
  const ignoredPrefixes = new Set(['chrome://', 'chrome-untrusted://']);
  if (!enableExtensions) {
    ignoredPrefixes.add('chrome-extension://');
  }

  return function targetFilter(target: Target): boolean {
    if (target.url() === 'chrome://newtab/') {
      return true;
    }
    // Could be the only page opened in the browser.
    if (target.url().startsWith('chrome://inspect')) {
      return true;
    }
    for (const prefix of ignoredPrefixes) {
      if (target.url().startsWith(prefix)) {
        return false;
      }
    }
    return true;
  };
}

// WebView2 exposes all targets (including about:blank frames) — accept everything.
function makeWebView2TargetFilter() {
  return function targetFilter(_target: Target): boolean {
    return true;
  };
}

export async function ensureBrowserConnected(options: {
  browserURL?: string;
  wsEndpoint?: string;
  wsHeaders?: Record<string, string>;
  devtools: boolean;
  channel?: Channel;
  userDataDir?: string;
  enableExtensions?: boolean;
  webview2?: boolean;
  connectTimeout?: number;
  connectRetryBudget?: number;
  connectRetryVerbose?: boolean;
  endpointSource?: ConnectionEndpointSource;
}) {
  const {channel, enableExtensions} = options;

  const sticky = getStickyError();
  if (sticky) {
    throw sticky;
  }

  if (browser?.connected && !isSessionStale()) {
    return browser;
  }

  if (isSessionStale()) {
    const reconnectUrl = options.browserURL ?? WEBVIEW2_DEBUG_URL;
    const blocker = beginReconnect(reconnectUrl);
    if (blocker) {
      throw blocker;
    }
    detachDisconnectListener();
    browser = undefined;
  }

  const isWebView2 = options.webview2 ?? false;
  const connectOptions: Parameters<typeof puppeteer.connect>[0] = {
    targetFilter: isWebView2
      ? makeWebView2TargetFilter()
      : makeTargetFilter(enableExtensions),
    defaultViewport: null,
    handleDevToolsAsPage: !isWebView2,
  };

  let autoConnect = false;
  const endpointSource =
    options.endpointSource ??
    (options.wsEndpoint
      ? 'wsEndpoint'
      : options.browserURL
        ? 'browserUrl'
        : options.channel || options.userDataDir
          ? 'autoDetect'
          : 'default');
  if (options.wsEndpoint) {
    connectOptions.browserWSEndpoint = options.wsEndpoint;
    setConnectionEndpoint(options.wsEndpoint, endpointSource);
    if (options.wsHeaders) {
      connectOptions.headers = options.wsHeaders;
    }
  } else if (options.browserURL) {
    connectOptions.browserURL = options.browserURL;
    setConnectionEndpoint(options.browserURL, endpointSource);
  } else if (channel || options.userDataDir) {
    const userDataDir = options.userDataDir;
    if (userDataDir) {
      autoConnect = true;
      // TODO: re-expose this logic via Puppeteer.
      const portPath = path.join(userDataDir, 'DevToolsActivePort');
      try {
        const fileContent = await fs.promises.readFile(portPath, 'utf8');
        const [rawPort, rawPath] = fileContent
          .split('\n')
          .map(line => {
            return line.trim();
          })
          .filter(line => {
            return !!line;
          });
        if (!rawPort || !rawPath) {
          throw new Error(`Invalid DevToolsActivePort '${fileContent}' found`);
        }
        const port = parseInt(rawPort, 10);
        if (isNaN(port) || port <= 0 || port > 65535) {
          throw new Error(`Invalid port '${rawPort}' found`);
        }
        const browserWSEndpoint = `ws://127.0.0.1:${port}${rawPath}`;
        connectOptions.browserWSEndpoint = browserWSEndpoint;
        setConnectionEndpoint(browserWSEndpoint, endpointSource);
      } catch (error) {
        throw new Error(
          `Could not connect to Chrome in ${userDataDir}. Check if Chrome is running and remote debugging is enabled by going to chrome://inspect/#remote-debugging.`,
          {
            cause: error,
          },
        );
      }
    } else {
      if (!channel) {
        throw new Error('Channel must be provided if userDataDir is missing');
      }
      connectOptions.channel = (
        channel === 'stable' ? 'chrome' : `chrome-${channel}`
      ) as ChromeReleaseChannel;
      setConnectionEndpoint('', endpointSource);
    }
  } else {
    throw new Error(
      'Either browserURL, wsEndpoint, channel or userDataDir must be provided',
    );
  }

  logger('Connecting Puppeteer to ', JSON.stringify(connectOptions));
  if (connectOptions.browserURL) {
    browser = await connectWithRetry({
      browserURL: connectOptions.browserURL,
      probeTimeoutMs: options.connectTimeout ?? 5000,
      retryBudgetMs: options.connectRetryBudget ?? 15000,
      verbose: options.connectRetryVerbose ?? false,
      connect: () => puppeteer.connect(connectOptions),
      onProbeResult: recordProbeResult,
    });
    markConnectionAttached(connectOptions.browserURL);
    attachDisconnectListener(browser);
  } else {
    try {
      browser = await puppeteer.connect(connectOptions);
      markConnectionAttached(
        endpointSource === 'browserUrl' || endpointSource === 'default'
          ? connectOptions.browserURL
          : browser.wsEndpoint(),
      );
      attachDisconnectListener(browser);
    } catch (err) {
      throw new Error(
        `Could not connect to Chrome. ${autoConnect ? `Check if Chrome is running and remote debugging is enabled by going to chrome://inspect/#remote-debugging.` : `Check if Chrome is running.`}`,
        {
          cause: err,
        },
      );
    }
  }
  logger('Connected Puppeteer');
  return browser;
}

interface McpLaunchOptions {
  acceptInsecureCerts?: boolean;
  executablePath?: string;
  channel?: Channel;
  userDataDir?: string;
  headless: boolean;
  isolated: boolean;
  logFile?: fs.WriteStream;
  viewport?: {
    width: number;
    height: number;
  };
  chromeArgs?: string[];
  ignoreDefaultChromeArgs?: string[];
  devtools: boolean;
  enableExtensions?: boolean;
  viaCli?: boolean;
}

export function detectDisplay(): void {
  // Only detect display on Linux/UNIX.
  if (os.platform() === 'win32' || os.platform() === 'darwin') {
    return;
  }
  if (!process.env['DISPLAY']) {
    try {
      const result = execSync(
        `ps -u $(id -u) -o pid= | xargs -I{} cat /proc/{}/environ 2>/dev/null | tr '\\0' '\\n' | grep -m1 '^DISPLAY=' | cut -d= -f2`,
      );
      const display = result.toString('utf8').trim();
      process.env['DISPLAY'] = display;
    } catch {
      // no-op
    }
  }
}

export async function launch(options: McpLaunchOptions): Promise<Browser> {
  const {channel, executablePath, headless, isolated} = options;
  const profileDirName =
    channel && channel !== 'stable'
      ? `chrome-profile-${channel}`
      : 'chrome-profile';

  let userDataDir = options.userDataDir;
  if (!isolated && !userDataDir) {
    userDataDir = path.join(
      os.homedir(),
      '.cache',
      options.viaCli ? 'excel-webview2-mcp-cli' : 'excel-webview2-mcp',
      profileDirName,
    );
    await fs.promises.mkdir(userDataDir, {
      recursive: true,
    });
  }

  const args: LaunchOptions['args'] = [
    ...(options.chromeArgs ?? []),
    '--hide-crash-restore-bubble',
  ];
  const ignoreDefaultArgs: LaunchOptions['ignoreDefaultArgs'] =
    options.ignoreDefaultChromeArgs ?? false;

  if (headless) {
    args.push('--screen-info={3840x2160}');
  }
  let puppeteerChannel: ChromeReleaseChannel | undefined;
  if (options.devtools) {
    args.push('--auto-open-devtools-for-tabs');
  }
  if (!executablePath) {
    puppeteerChannel =
      channel && channel !== 'stable'
        ? (`chrome-${channel}` as ChromeReleaseChannel)
        : 'chrome';
  }

  if (!headless) {
    detectDisplay();
  }

  try {
    const browser = await puppeteer.launch({
      channel: puppeteerChannel,
      targetFilter: makeTargetFilter(options.enableExtensions),
      executablePath,
      defaultViewport: null,
      userDataDir,
      pipe: true,
      headless,
      args,
      ignoreDefaultArgs: ignoreDefaultArgs,
      acceptInsecureCerts: options.acceptInsecureCerts,
      handleDevToolsAsPage: true,
      enableExtensions: options.enableExtensions,
    });
    if (options.logFile) {
      // FIXME: we are probably subscribing too late to catch startup logs. We
      // should expose the process earlier or expose the getRecentLogs() getter.
      browser.process()?.stderr?.pipe(options.logFile);
      browser.process()?.stdout?.pipe(options.logFile);
    }
    if (options.viewport) {
      const [page] = await browser.pages();
      await page?.resize({
        contentWidth: options.viewport.width,
        contentHeight: options.viewport.height,
      });
    }
    return browser;
  } catch (error) {
    if (
      userDataDir &&
      (error as Error).message.includes('The browser is already running')
    ) {
      throw new Error(
        `The browser is already running for ${userDataDir}. Use --isolated to run multiple browser instances.`,
        {
          cause: error,
        },
      );
    }
    throw error;
  }
}

export async function ensureBrowserLaunched(
  options: McpLaunchOptions,
): Promise<Browser> {
  if (browser?.connected) {
    return browser;
  }
  browser = await launch(options);
  return browser;
}

export type Channel = 'stable' | 'canary' | 'beta' | 'dev';
