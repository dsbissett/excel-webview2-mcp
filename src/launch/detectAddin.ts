import fs from 'node:fs/promises';
import path from 'node:path';

export interface AddinProject {
  root: string;
  manifestPath: string;
  manifestKind: 'xml' | 'json';
  packageManager: 'npm' | 'pnpm' | 'yarn';
  existingCdpScript?: string;
  devServer?: {
    script: string;
    port: number;
  };
}

interface PackageJson {
  scripts?: Record<string, string>;
  config?: {
    dev_server_port?: unknown;
  };
}

export async function detectExcelAddin(
  cwd: string,
): Promise<AddinProject | null> {
  const root = await findPackageRoot(cwd);
  if (!root) {
    return null;
  }

  const packageJson = await readPackageJson(path.join(root, 'package.json'));
  if (!packageJson) {
    return null;
  }

  const manifest = await detectManifest(root);
  if (!manifest) {
    return null;
  }

  return {
    root,
    manifestPath: manifest.manifestPath,
    manifestKind: manifest.manifestKind,
    packageManager: await detectPackageManager(root),
    existingCdpScript: detectExistingCdpScript(packageJson.scripts),
    devServer: detectDevServer(packageJson),
  };
}

function detectDevServer(packageJson: PackageJson): AddinProject['devServer'] {
  const scripts = packageJson.scripts;
  if (!scripts) {
    return undefined;
  }
  const scriptName = ['dev-server', 'dev:server', 'dev'].find(name => {
    return typeof scripts[name] === 'string' && scripts[name].length > 0;
  });
  if (!scriptName) {
    return undefined;
  }
  const rawPort = packageJson.config?.dev_server_port;
  const port = typeof rawPort === 'number' ? rawPort : Number(rawPort);
  if (!Number.isFinite(port) || port <= 0) {
    return undefined;
  }
  return {script: scriptName, port};
}

async function findPackageRoot(cwd: string): Promise<string | null> {
  let currentDir = path.resolve(cwd);

  for (let depth = 0; depth <= 5; depth += 1) {
    if (await pathExists(path.join(currentDir, 'package.json'))) {
      return currentDir;
    }

    const parentDir = path.dirname(currentDir);
    if (parentDir === currentDir) {
      break;
    }
    currentDir = parentDir;
  }

  return null;
}

async function readPackageJson(
  packageJsonPath: string,
): Promise<PackageJson | null> {
  try {
    const content = await fs.readFile(packageJsonPath, 'utf8');
    return JSON.parse(content) as PackageJson;
  } catch {
    return null;
  }
}

async function detectManifest(
  root: string,
): Promise<Pick<AddinProject, 'manifestPath' | 'manifestKind'> | null> {
  const xmlManifestPath = path.join(root, 'manifest.xml');
  if (await isWorkbookXmlManifest(xmlManifestPath)) {
    return {
      manifestPath: xmlManifestPath,
      manifestKind: 'xml',
    };
  }

  const jsonManifestPath = path.join(root, 'manifest.json');
  if (await isWorkbookJsonManifest(jsonManifestPath)) {
    return {
      manifestPath: jsonManifestPath,
      manifestKind: 'json',
    };
  }

  return null;
}

async function isWorkbookXmlManifest(manifestPath: string): Promise<boolean> {
  try {
    const content = await fs.readFile(manifestPath, 'utf8');
    return (
      /<OfficeApp\b/i.test(content) &&
      /<Host\b[^>]*Name\s*=\s*["']Workbook["']/i.test(content)
    );
  } catch {
    return false;
  }
}

async function isWorkbookJsonManifest(manifestPath: string): Promise<boolean> {
  try {
    const content = await fs.readFile(manifestPath, 'utf8');
    const manifest = JSON.parse(content) as {
      extensions?: Array<{
        requirements?: {
          scopes?: unknown;
        };
      }>;
    };

    return (manifest.extensions ?? []).some(extension => {
      const scopes = extension.requirements?.scopes;
      return (
        Array.isArray(scopes) &&
        scopes.some(scope => String(scope).toLowerCase() === 'workbook')
      );
    });
  } catch {
    return false;
  }
}

function detectExistingCdpScript(
  scripts: Record<string, string> | undefined,
): string | undefined {
  if (!scripts) {
    return undefined;
  }

  for (const [name, command] of Object.entries(scripts)) {
    if (
      command.includes('--remote-debugging-port') ||
      command.includes('WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS')
    ) {
      return name;
    }
  }

  return undefined;
}

async function detectPackageManager(
  root: string,
): Promise<AddinProject['packageManager']> {
  if (await pathExists(path.join(root, 'pnpm-lock.yaml'))) {
    return 'pnpm';
  }
  if (await pathExists(path.join(root, 'yarn.lock'))) {
    return 'yarn';
  }
  return 'npm';
}

async function pathExists(filePath: string): Promise<boolean> {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}
