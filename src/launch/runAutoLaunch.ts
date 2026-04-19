import {getLifecycleDeps, trackedByManifest} from '../tools/lifecycleState.js';

import {LaunchError} from './launchExcel.js';

export async function runAutoLaunch(args: {
  launchPort?: number;
  launchTimeout?: number;
  cwd?: string;
  logger?: (msg: string) => void;
}): Promise<void> {
  const log = args.logger ?? (() => undefined);
  const deps = getLifecycleDeps();
  const workingDir = args.cwd ?? deps.cwd();
  const project = await deps.detectExcelAddin(workingDir);
  if (!project) {
    log(`auto-launch: no Excel add-in detected at ${workingDir}; skipping.`);
    return;
  }

  if (trackedByManifest.has(project.manifestPath)) {
    log(`auto-launch: already tracking ${project.manifestPath}; skipping.`);
    return;
  }

  const port = args.launchPort ?? 9222;
  const timeoutMs = args.launchTimeout ?? 60_000;

  try {
    log(`auto-launch: launching ${project.manifestPath} on port ${port}...`);
    const result = await deps.launchExcel({project, port, timeoutMs});
    trackedByManifest.set(project.manifestPath, {project, result});
    log(`auto-launch: launched pid=${result.pid} at ${result.cdpUrl}.`);
  } catch (error) {
    if (error instanceof LaunchError) {
      log(`auto-launch failed [${error.reason}]: ${error.message}`);
      return;
    }
    throw error;
  }
}
