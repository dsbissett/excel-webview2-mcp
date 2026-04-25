import {execFile} from 'node:child_process';
import {promisify} from 'node:util';

const execFileAsync = promisify(execFile);

export interface RemainingPort {
  port: number;
  pid: number;
  state?: string;
}

export interface ForceShutdownResult {
  taskkillOutput: string;
  port3000CleanupOutput: string;
  finalCleanupOutput: string;
  remaining: RemainingPort[];
}

const PORT_3000_CLEANUP_PS = `Get-NetTCPConnection -LocalPort 3000 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique | ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue; "killed pid $_" }`;

const FINAL_CLEANUP_PS = `$conns = Get-NetTCPConnection -LocalPort 3000,9222 -ErrorAction SilentlyContinue; $conns | Format-Table LocalPort,State,OwningProcess; $conns.OwningProcess | Sort-Object -Unique | Where-Object { $_ -gt 4 } | ForEach-Object { try { $p = Get-Process -Id $_ -ErrorAction Stop; "killing $($p.Id) $($p.ProcessName)"; Stop-Process -Id $_ -Force } catch {} }`;

const VERIFY_PS = `$c = Get-NetTCPConnection -LocalPort 3000,9222 -ErrorAction SilentlyContinue | Where-Object { $_.OwningProcess -gt 4 }; if ($c) { $c | ForEach-Object { "REMAINING $($_.LocalPort) $($_.State) $($_.OwningProcess)" } }`;

async function runPowerShell(script: string): Promise<string> {
  try {
    const {stdout, stderr} = await execFileAsync(
      'powershell.exe',
      ['-NoProfile', '-NonInteractive', '-Command', script],
      {windowsHide: true, maxBuffer: 4 * 1024 * 1024},
    );
    return [stdout, stderr].filter(Boolean).join('\n').trim();
  } catch (err) {
    const e = err as {stdout?: string; stderr?: string; message?: string};
    return [e.stdout ?? '', e.stderr ?? '', e.message ?? String(err)]
      .filter(Boolean)
      .join('\n')
      .trim();
  }
}

async function runTaskkillFallback(): Promise<string> {
  const lines: string[] = [];
  for (const args of [
    ['/F', '/IM', 'excel.exe'],
    ['/F', '/IM', 'node.exe', '/FI', 'WINDOWTITLE eq *office-addin*'],
  ]) {
    try {
      const {stdout, stderr} = await execFileAsync('taskkill', args, {
        windowsHide: true,
      });
      lines.push([stdout, stderr].filter(Boolean).join('\n').trim());
    } catch (err) {
      const e = err as {stdout?: string; stderr?: string; message?: string};
      lines.push(
        [e.stdout ?? '', e.stderr ?? '', e.message ?? String(err)]
          .filter(Boolean)
          .join('\n')
          .trim(),
      );
    }
  }
  return lines.filter(Boolean).join('\n');
}

function parseRemaining(verifyOutput: string): RemainingPort[] {
  const remaining: RemainingPort[] = [];
  for (const line of verifyOutput.split(/\r?\n/)) {
    const match = line.match(/^REMAINING\s+(\d+)\s+(\S+)\s+(\d+)/);
    if (match) {
      remaining.push({
        port: Number(match[1]),
        state: match[2],
        pid: Number(match[3]),
      });
    }
  }
  return remaining;
}

export async function forceShutdownAddinProcesses(): Promise<ForceShutdownResult> {
  if (process.platform !== 'win32') {
    return {
      taskkillOutput: '',
      port3000CleanupOutput: '',
      finalCleanupOutput: '',
      remaining: [],
    };
  }

  const taskkillOutput = await runTaskkillFallback();
  const port3000CleanupOutput = await runPowerShell(PORT_3000_CLEANUP_PS);
  const finalCleanupOutput = await runPowerShell(FINAL_CLEANUP_PS);
  const verifyOutput = await runPowerShell(VERIFY_PS);
  const remaining = parseRemaining(verifyOutput);

  return {
    taskkillOutput,
    port3000CleanupOutput,
    finalCleanupOutput,
    remaining,
  };
}
