import type {ParsedArguments} from '../bin/excel-webview2-mcp-cli-options.js';
import {type AddinProject} from '../launch/detectAddin.js';
import {LaunchError} from '../launch/launchExcel.js';
import {zod} from '../third_party/index.js';

import {ToolCategory} from './categories.js';
import {
  getLifecycleDeps,
  trackedByManifest,
  type TrackedLaunchEntry,
} from './lifecycleState.js';
import {defineTool} from './ToolDefinition.js';

async function resolveProject(
  cwdOverride: string | undefined,
  manifestPath: string | undefined,
): Promise<AddinProject> {
  const deps = getLifecycleDeps();
  const workingDir = cwdOverride ?? deps.cwd();
  const detected = await deps.detectExcelAddin(workingDir);
  if (detected) {
    if (manifestPath && manifestPath !== detected.manifestPath) {
      return {...detected, manifestPath};
    }
    return detected;
  }
  throw new Error(
    `Could not detect an Excel add-in project at ${workingDir}. Provide cwd or manifestPath.`,
  );
}

export const excelDetectAddin = defineTool({
  name: 'excel_detect_addin',
  description:
    'Inspects a working directory and reports whether it looks like an Excel add-in project (manifest location, manifest kind, package manager, and any existing remote-debugging script).',
  annotations: {
    category: ToolCategory.LIFECYCLE,
    readOnlyHint: true,
  },
  requiresContext: false,
  schema: {
    cwd: zod
      .string()
      .optional()
      .describe(
        'Directory to inspect. Defaults to the MCP server working directory.',
      ),
  },
  handler: async (request, response) => {
    const deps = getLifecycleDeps();
    const workingDir = request.params.cwd ?? deps.cwd();
    const project = await deps.detectExcelAddin(workingDir);
    if (!project) {
      response.appendResponseLine(
        `No Excel add-in project detected at ${workingDir}.`,
      );
      response.setStructuredContent({detected: false, cwd: workingDir});
      return;
    }
    response.setStructuredContent({detected: true, project});
    response.appendResponseLine(JSON.stringify(project, null, 2));
  },
});

export const excelLaunchAddin = defineTool((args?: ParsedArguments) => ({
  name: 'excel_launch_addin',
  description:
    'Launches Excel with the detected add-in and WebView2 remote debugging enabled. Idempotent per manifest path: re-calling returns the tracked launch instead of spawning a duplicate.',
  annotations: {
    category: ToolCategory.LIFECYCLE,
    readOnlyHint: false,
  },
  requiresContext: false,
  schema: {
    cwd: zod.string().optional(),
    port: zod.number().int().positive().optional(),
    manifestPath: zod.string().optional(),
    extraBrowserArgs: zod.array(zod.string()).optional(),
    timeoutMs: zod.number().int().positive().optional(),
    autoConnect: zod.boolean().optional(),
  },
  handler: async (request, response) => {
    const deps = getLifecycleDeps();
    const port = request.params.port ?? args?.launchPort ?? 9222;
    const timeoutMs = request.params.timeoutMs ?? args?.launchTimeout ?? 60_000;
    const autoConnect = request.params.autoConnect ?? true;

    const project = await resolveProject(
      request.params.cwd,
      request.params.manifestPath,
    );

    response.appendResponseLine(
      `Detected Excel add-in at ${project.root} (manifest: ${project.manifestPath}).`,
    );

    const existing = trackedByManifest.get(project.manifestPath);
    if (existing) {
      response.appendResponseLine(
        `Reusing tracked launch (pid=${existing.result.pid}) at ${existing.result.cdpUrl}.`,
      );
      await maybeConnect(autoConnect, existing.result.cdpUrl, response);
      response.setStructuredContent({
        reused: true,
        pid: existing.result.pid,
        cdpUrl: existing.result.cdpUrl,
        project: existing.project,
      });
      return;
    }

    response.appendResponseLine(
      `Launching office-addin-debugging on port ${port} (timeout ${timeoutMs}ms)...`,
    );

    try {
      const result = await deps.launchExcel({
        project,
        port,
        extraBrowserArgs: request.params.extraBrowserArgs,
        timeoutMs,
      });
      trackedByManifest.set(project.manifestPath, {project, result});
      response.appendResponseLine(
        `Launched (pid=${result.pid}); CDP endpoint is ${result.cdpUrl}.`,
      );
      await maybeConnect(autoConnect, result.cdpUrl, response);
      response.setStructuredContent({
        reused: false,
        pid: result.pid,
        cdpUrl: result.cdpUrl,
        project,
      });
    } catch (error) {
      if (error instanceof LaunchError) {
        response.appendResponseLine(
          `ERROR [${error.reason}]: ${error.message}`,
        );
        if (error.output.length > 0) {
          response.appendResponseLine('--- launcher output ---');
          for (const line of error.output) {
            response.appendResponseLine(line);
          }
        }
        return;
      }
      throw error;
    }
  },
}));

export const excelStopAddin = defineTool({
  name: 'excel_stop_addin',
  description:
    'Stops the most recent Excel add-in launched by excel_launch_addin (or a specific manifest). Runs office-addin-debugging stop and kills the process if it does not exit cleanly.',
  annotations: {
    category: ToolCategory.LIFECYCLE,
    readOnlyHint: false,
  },
  requiresContext: false,
  schema: {
    manifestPath: zod.string().optional(),
  },
  handler: async (request, response) => {
    const manifestPath = request.params.manifestPath;
    const entries: TrackedLaunchEntry[] = [];
    if (manifestPath) {
      const entry = trackedByManifest.get(manifestPath);
      if (!entry) {
        response.appendResponseLine(
          `No tracked launch for manifest ${manifestPath}.`,
        );
        return;
      }
      entries.push(entry);
    } else {
      entries.push(...trackedByManifest.values());
    }

    if (entries.length === 0) {
      response.appendResponseLine('No tracked Excel add-in launches to stop.');
      return;
    }

    for (const entry of entries) {
      try {
        await entry.result.stop();
        trackedByManifest.delete(entry.project.manifestPath);
        response.appendResponseLine(
          `Stopped launch for ${entry.project.manifestPath} (pid=${entry.result.pid}).`,
        );
      } catch (error) {
        trackedByManifest.delete(entry.project.manifestPath);
        if (error instanceof LaunchError) {
          response.appendResponseLine(
            `WARN [${error.reason}]: ${error.message}`,
          );
          continue;
        }
        throw error;
      }
    }
  },
});

async function maybeConnect(
  autoConnect: boolean,
  cdpUrl: string,
  response: {appendResponseLine: (value: string) => void},
): Promise<void> {
  if (!autoConnect) {
    return;
  }
  try {
    await getLifecycleDeps().ensureBrowserConnected({
      browserURL: cdpUrl,
      devtools: false,
      webview2: true,
      endpointSource: 'browserUrl',
    });
    response.appendResponseLine(`Connected to ${cdpUrl}.`);
  } catch (error) {
    response.appendResponseLine(
      `WARN: autoConnect failed: ${(error as Error).message}`,
    );
  }
}
