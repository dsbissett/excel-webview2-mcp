import '../polyfill.js';

import process from 'node:process';

import {createMcpServer, logDisclaimers} from '../index.js';
import {shutdownAllLaunches} from '../launch/launchExcel.js';
import {runAutoLaunch} from '../launch/runAutoLaunch.js';
import {logger, saveLogsToFile} from '../logger.js';
import {computeFlagUsage} from '../telemetry/flagUtils.js';
import {StdioServerTransport} from '../third_party/index.js';
import {checkForUpdates} from '../utils/check-for-updates.js';
import {VERSION} from '../version.js';

import {cliOptions, parseArguments} from './excel-webview2-mcp-cli-options.js';

await checkForUpdates(
  'Run `npm install @dsbissett/excel-webview2-mcp@latest` to update.',
);

export const args = parseArguments(VERSION);

const logFile = args.logFile ? saveLogsToFile(args.logFile) : undefined;
if (
  process.env['CI'] ||
  process.env['EXCEL_WEBVIEW2_MCP_NO_USAGE_STATISTICS']
) {
  console.error(
    "turning off usage statistics. process.env['CI'] || process.env['EXCEL_WEBVIEW2_MCP_NO_USAGE_STATISTICS'] is set.",
  );
  args.usageStatistics = false;
}

if (process.env['EXCEL_WEBVIEW2_MCP_CRASH_ON_UNCAUGHT'] !== 'true') {
  process.on('unhandledRejection', (reason, promise) => {
    logger('Unhandled promise rejection', promise, reason);
  });
}

logger(`Starting Excel WebView2 MCP Server v${VERSION}`);
if (args.autoLaunch) {
  await runAutoLaunch({
    launchPort: args.launchPort,
    launchTimeout: args.launchTimeout,
    logger: msg => logger(msg),
  });
}
const {server, clearcutLogger} = await createMcpServer(args, {
  logFile,
});
const transport = new StdioServerTransport();
await server.connect(transport);
logger('Excel WebView2 MCP Server connected');
logDisclaimers(args);
void clearcutLogger?.logDailyActiveIfNeeded();
void clearcutLogger?.logServerStart(computeFlagUsage(args, cliOptions));

let shuttingDown = false;
const onParentDeath = (reason: string) => {
  if (shuttingDown) {
    return;
  }
  shuttingDown = true;
  logger(`Parent death detected (${reason}); stopping tracked launches.`);
  void shutdownAllLaunches()
    .catch(err => logger('shutdownAllLaunches failed', err))
    .finally(() => process.exit(0));
};
process.stdin.on('end', () => onParentDeath('stdin end'));
process.stdin.on('close', () => onParentDeath('stdin close'));
process.stdout.on('error', err => {
  if ((err as NodeJS.ErrnoException).code === 'EPIPE') {
    onParentDeath('stdout EPIPE');
  }
});
process.on('disconnect', () => onParentDeath('ipc disconnect'));
