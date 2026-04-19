import {ensureBrowserConnected} from '../browser.js';
import {
  detectExcelAddin as detectExcelAddinImpl,
  type AddinProject,
} from '../launch/detectAddin.js';
import {
  launchExcel as launchExcelImpl,
  type LaunchOptions,
  type LaunchResult,
} from '../launch/launchExcel.js';

export type DetectFn = (cwd: string) => Promise<AddinProject | null>;
export type LaunchFn = (options: LaunchOptions) => Promise<LaunchResult>;
export type ConnectFn = typeof ensureBrowserConnected;

export interface LifecycleDeps {
  detectExcelAddin: DetectFn;
  launchExcel: LaunchFn;
  ensureBrowserConnected: ConnectFn;
  cwd: () => string;
}

export interface TrackedLaunchEntry {
  project: AddinProject;
  result: LaunchResult;
}

const defaultDeps: LifecycleDeps = {
  detectExcelAddin: detectExcelAddinImpl,
  launchExcel: launchExcelImpl,
  ensureBrowserConnected,
  cwd: () => process.cwd(),
};

let currentDeps: LifecycleDeps = defaultDeps;
export const trackedByManifest = new Map<string, TrackedLaunchEntry>();

export function getLifecycleDeps(): LifecycleDeps {
  return currentDeps;
}

export function setLifecycleDepsForTesting(
  overrides: Partial<LifecycleDeps>,
): void {
  currentDeps = {...defaultDeps, ...overrides};
  trackedByManifest.clear();
}

export function resetLifecycleDepsForTesting(): void {
  currentDeps = defaultDeps;
  trackedByManifest.clear();
}
