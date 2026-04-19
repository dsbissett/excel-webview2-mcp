# Changelog

All notable changes since commit `412b24e` are documented here.

## [0.0.2]

### Added

#### Excel Add-in Launch & Lifecycle
- New `src/launch/` module with `launchExcel` and `detectAddin` for spawning Excel with WebView2 remote debugging (CDP on port 9222).
- Three lifecycle tools under a new `LIFECYCLE` category:
  - `excel_detect_addin` — detect add-in manifests in a project.
  - `excel_launch_addin` — launch Excel + dev server, idempotent per manifest.
  - `excel_stop_addin` — stop tracked Excel/dev-server processes (force-kills process tree on Windows via `taskkill`).
- `runAutoLaunch` module and `--auto-launch` CLI option for launching an add-in at MCP server startup.
- Centralized lifecycle state module (`src/tools/lifecycleState.ts`).
- TCP socket-based port detection replacing HTTP fetch probes (better Node 24 compatibility).
- CLI options for add-in launch configuration (manifest path, project root, dev-server command, etc.).

#### Connection Robustness
- New `src/connection/` module split into focused files:
  - `probe.ts` — endpoint availability probing.
  - `retry.ts` — retry loop with configurable budget.
  - `session.ts` — session-level reconnect management with circuit breaker.
  - `error.ts` — structured connection error reasons (including new `unstable` reason).
  - `status.ts` — connection status reporting.
- New `connection_status` tool for inspecting current CDP connection health.
- Disconnect listener attached after successful WebView2 connection to trigger reconnect flow.
- New CLI options for connection tuning (retry budget, probe interval, etc.).

### Changed
- Removed per-file license/copyright headers from 155+ source and test files (Chrome DevTools upstream headers no longer apply to this fork).
- Removed custom `check-license-rule` ESLint rule.
- README and troubleshooting docs updated with WebView2 connection guidance.
- `excel-webview2` skill documentation updated to cover agent-driven Excel launching and lifecycle tools.
- Refactored `runAutoLaunch` out of the lifecycle tools module so tool registries only export tool definitions (fixes docs generator).

### Fixed
- Import order in main entry point.
- Unused `ChildProcess` import removed.
- Dev server now receives correct `projectRoot` when launched via `excel_launch_addin`.
- `spawn EINVAL` on Windows wrapped in try-catch for clearer error reporting.
- Dev server cleanup failure in `excel_stop_addin` resolved via Windows process-tree kill.

### Removed
- Obsolete Chrome DevTools eval scenarios (console, network, input, snapshot, emulation, performance, page-focus-keyboard).

### Version
- Bumped from `0.0.1` → `0.0.2` across `src/version.ts`, `package.json`, `package-lock.json`, `server.json`, `.release-please-manifest.json`, `.claude-plugin/plugin.json`, and `.claude-plugin/marketplace.json`.
