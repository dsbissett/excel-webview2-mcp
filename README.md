# Excel WebView2 MCP

Professional MCP connectivity for Microsoft Excel add-ins running inside WebView2.

| Item              | Value                                        |
| ----------------- | -------------------------------------------- |
| Package           | `@dsbissett/excel-webview2-mcp`              |
| Upstream          | Fork of `ChromeDevTools/chrome-devtools-mcp` |
| Target runtime    | Excel add-ins hosted in WebView2             |
| Debug endpoint    | `http://127.0.0.1:9222`                      |
| Recommended setup | Claude Code plugin marketplace               |

> [!IMPORTANT]
> As of **v0.0.2**, this server can launch Excel and sideload your add-in for you. You no longer need to start the dev server or attach a debugger before using it — `excel_launch_addin` will start your local dev server (if configured) and launch Excel with WebView2 remote debugging enabled on port `9222`.
> Manual pre-launch is still supported: if an Excel add-in is already running with remote debugging on `9222`, the server will attach to it.

## What's New in v0.0.2

- **Add-in lifecycle management** — new tools launch and stop Excel add-ins directly from Claude Code:
  - [excel_detect_addin](src/tools/lifecycle.ts) — discover manifest and dev server configuration for the current project.
  - [excel_launch_addin](src/tools/lifecycle.ts) — start the dev server (if needed) and launch Excel with the add-in sideloaded and CDP port `9222` enabled.
  - [excel_stop_addin](src/tools/lifecycle.ts) — tear down the launched Excel session and dev server process tree.
- **Read-only Excel inspection tools** — a broad set of read operations for inspecting workbooks, worksheets, ranges, tables, pivots, charts, and more. None of these mutate workbook state:
  - Context & structure: `excel_context_info`, `excel_workbook_info`, `excel_list_worksheets`, `excel_worksheet_info`, `excel_active_range`, `excel_used_range`.
  - Range reads: `excel_read_range`, `excel_range_properties`, `excel_range_formulas`, `excel_range_special_cells`, `excel_find_in_range`.
  - Formatting & validation: `excel_list_conditional_formats`, `excel_list_data_validations`.
  - Tables & names: `excel_list_tables`, `excel_table_info`, `excel_table_rows`, `excel_table_filters`, `excel_list_named_items`.
  - Comments & shapes: `excel_list_comments`, `excel_list_shapes`.
  - Calculation & pivots: `excel_calculation_state`, `excel_list_pivot_tables`, `excel_pivot_table_info`, `excel_pivot_table_values`.
  - Charts: `excel_list_charts`, `excel_chart_info`, `excel_chart_image`.
  - Misc: `excel_custom_xml_parts`, `excel_settings_get`.
- **Socket-based port detection** replaces HTTP polling for more reliable dev-server and CDP readiness checks on Node 24.
- **Robust cleanup on Windows** — `excel_stop_addin` now force-kills the dev server process tree via `taskkill`.

## Fork Notice

This repository is a fork of the Chrome DevTools MCP repository, [`ChromeDevTools/chrome-devtools-mcp`](https://github.com/ChromeDevTools/chrome-devtools-mcp). It preserves the upstream DevTools and MCP foundation, while adapting the connection model for Microsoft Excel add-ins hosted in WebView2.

## What This Project Does

- Connects Claude Code to a locally running Excel add-in through the Chrome DevTools Protocol (CDP).
- Exposes MCP tools for inspection, automation, screenshots, console access, network inspection, and performance analysis.
- Targets the embedded WebView2 runtime used by Excel add-ins instead of a standalone Chrome session.

## Connection Model

```text
Claude Code
    |
    v
excel-webview2-mcp
    |
    v
WebView2 remote debugging endpoint (localhost:9222)
    |
    v
Locally running Excel add-in
```

That separation matters: `excel-webview2-mcp` is a bridge to an existing debug session. It is not the thing that launches or hosts the add-in.

## Prerequisites

You have two supported workflows:

### Auto-launch (recommended, v0.0.2+)

1. Your Office add-in project (with a `manifest.xml` and a dev server script) lives on disk.
2. Excel desktop is installed on Windows.
3. Node.js is installed and `npx @dsbissett/excel-webview2-mcp@latest` is runnable.

Call `excel_detect_addin` first to confirm the project is discovered, then `excel_launch_addin` to start the dev server and sideload the add-in into Excel with CDP port `9222` enabled. Use `excel_stop_addin` to tear everything down.

### Manual / pre-attached

1. Your Excel add-in is already loaded and running in the local Excel desktop client.
2. WebView2 remote debugging is enabled and bound to port `9222` (see [Launching Excel with the debug port](#launching-excel-with-the-debug-port)).
3. The debugging endpoint is reachable at `http://127.0.0.1:9222`.

Verify with:

```sh
curl http://127.0.0.1:9222/json
```

## Launching Excel with the debug port

This section applies to **Excel desktop on Windows**. Per Microsoft's WebView2 and Office add-in documentation, the supported way to pass Chromium flags into the WebView2 runtime is `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS`. The WebView2 team also documents a registry-based fallback for persistent configuration.

If you are using **Excel for Mac**, this MCP server does not apply. Microsoft documents Excel for Mac debugging through Safari Web Inspector instead of a WebView2 CDP port.

Sources:

- WebView2 debug arguments and registry policy: <https://learn.microsoft.com/en-us/microsoft-edge/webview2/how-to/debug-visual-studio-code>
- Office add-ins debugging with Edge DevTools: <https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-chromium>
- Office add-ins debugging overview, including Mac: <https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-add-ins-overview>

### Windows: preferred local-dev setup

Set the environment variable before launching Excel so the Excel process inherits it:

```powershell
$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS="--remote-debugging-port=9222"
```

Then:

1. Launch Excel from that same shell or from a parent process that inherited the variable.
2. Start your add-in locally.
3. Confirm the debug endpoint is live.

Worked example:

```powershell
$env:WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS="--remote-debugging-port=9222"
start excel.exe
```

The Office add-ins team documents this environment-variable mechanism generically for WebView2-hosted add-ins, and the WebView2 team documents `--remote-debugging-port=9222` as a valid browser argument.

### Windows: persistent registry fallback

If you need a persistent machine-local setting, the WebView2 team documents this registry policy path:

```text
HKEY_CURRENT_USER\Software\Policies\Microsoft\Edge\WebView2\AdditionalBrowserArguments
```

Use:

- Value name: `EXCEL.EXE`
- Value data: `--remote-debugging-port=9222`

Use the environment variable first for local development. It is easier to turn on and off and avoids leaving a persistent machine-wide setting behind.

### Verify the endpoint

Run:

```sh
curl http://127.0.0.1:9222/json/version
```

A healthy response is JSON and includes fields such as:

```json
{
  "Browser": "...",
  "Protocol-Version": "...",
  "webSocketDebuggerUrl": "ws://127.0.0.1:9222/devtools/browser/..."
}
```

If that `curl` command fails, Excel is not exposing the WebView2 debug port yet, and `excel-webview2-mcp` will not be able to attach.

## Installation

The server ships as the npm package `@dsbissett/excel-webview2-mcp` and is invoked via `npx`. The configuration below is the same across every host — only the location of the config file changes.

Canonical MCP entry:

```json
{
  "mcpServers": {
    "excel-webview2": {
      "command": "npx",
      "args": ["@dsbissett/excel-webview2-mcp@latest"]
    }
  }
}
```

<details>
<summary><strong>Claude Code (CLI)</strong></summary>

One-liner:

```sh
claude mcp add excel-webview2 -- npx @dsbissett/excel-webview2-mcp@latest
```

Use `claude mcp add --scope user ...` to make the server available in every project, or `--scope project` to check the config into `.mcp.json` for teammates.

Alternatively, install the bundled plugin:

1. Add the marketplace from [`.claude-plugin/marketplace.json`](.claude-plugin/marketplace.json) with `/plugin marketplace add dsbissett/excel-webview2-mcp`.
2. Install with `/plugin install excel-webview2-mcp`.

</details>

<details>
<summary><strong>Claude Code (VS Code extension)</strong></summary>

1. Open the Claude Code side panel in VS Code.
2. Open the command palette and run **Claude Code: Manage MCP Servers** (or click the MCP icon in the Claude panel).
3. Choose **Add Server** and paste the canonical MCP entry above, or run the `claude mcp add` command in the VS Code integrated terminal — the extension reads the same config.

</details>

<details>
<summary><strong>Cursor</strong></summary>

Edit `~/.cursor/mcp.json` (global) or `.cursor/mcp.json` in the project root and add the canonical MCP entry above. Restart Cursor, then open **Settings → MCP** to confirm `excel-webview2` shows as connected.

</details>

<details>
<summary><strong>Codex (OpenAI Codex CLI)</strong></summary>

Codex reads MCP servers from `~/.codex/config.toml`. Add:

```toml
[mcp_servers.excel-webview2]
command = "npx"
args = ["@dsbissett/excel-webview2-mcp@latest"]
```

Then launch `codex` and confirm the server appears in `/mcp`.

</details>

<details>
<summary><strong>GitHub Copilot (VS Code)</strong></summary>

GitHub Copilot Chat in VS Code supports MCP servers through the agent-mode configuration.

1. Create or edit `.vscode/mcp.json` in your workspace (or the user-level `mcp.json` via **MCP: Open User Configuration** from the command palette).
2. Add the server entry:

   ```json
   {
     "servers": {
       "excel-webview2": {
         "command": "npx",
         "args": ["@dsbissett/excel-webview2-mcp@latest"]
       }
     }
   }
   ```

3. Open Copilot Chat, switch to **Agent** mode, and click the **Tools** icon to confirm the `excel-webview2` tools are available. Use **MCP: List Servers** from the command palette to inspect status or restart the server.

</details>

### Verifying the install

After adding the server, ask the model to call `excel_detect_addin` from your Office add-in project directory. A successful response confirms the MCP server is wired up; from there `excel_launch_addin` will take care of starting Excel.

By default, the server connects to the local WebView2 debugging endpoint at `http://127.0.0.1:9222`.

## Local Development

```sh
npm install
npm run build
npm start
```

Supported Node.js versions are `^20.19.0`, `^22.12.0`, or `>=23`.
