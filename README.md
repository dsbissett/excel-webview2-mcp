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
> This project does **not** launch Excel, start your add-in, or create a browser session.
> It only connects Claude Code to an **already-running** WebView2 remote debugging endpoint on port `9222`.
> Your Excel add-in must already be running locally and actively being debugged before this server can do anything useful.

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

Before using this server, make sure all of the following are true:

1. Your Excel add-in is loaded in the local Excel desktop client.
2. The add-in is already running under local debugging.
3. WebView2 remote debugging is enabled and bound to port `9222`.
4. The debugging endpoint is reachable at `http://127.0.0.1:9222`.
5. Claude Code can run `npx @dsbissett/excel-webview2-mcp@latest`.

The exact way you enable WebView2 remote debugging depends on your Office add-in launch workflow, but the end result must be a live CDP endpoint on port `9222`.

To verify that the endpoint is available before starting Claude Code:

```sh
curl http://127.0.0.1:9222/json
```

> [!WARNING]
> If the add-in is not already running locally with remote debugging enabled on port `9222`, this server has nothing to attach to.

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

## Claude Code Setup

The recommended path is to install this project through the bundled Claude Code plugin marketplace metadata included in this repository.

1. Add the plugin marketplace from [`.claude-plugin/marketplace.json`](.claude-plugin/marketplace.json) to Claude Code.
2. Install the `excel-webview2-mcp` plugin from that marketplace.
3. Start Excel and run your add-in locally with WebView2 remote debugging enabled on port `9222`.
4. Open Claude Code and connect through the installed plugin.

The plugin definition is stored in [`.claude-plugin/plugin.json`](.claude-plugin/plugin.json) and uses this MCP server entry:

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

## Direct MCP Configuration

If you prefer to configure Claude Code manually instead of installing the plugin, add the same server entry directly:

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

By default, the server connects to the local WebView2 debugging endpoint at `http://localhost:9222`.

## Local Development

```sh
npm install
npm run build
npm start
```

Supported Node.js versions are `^20.19.0`, `^22.12.0`, or `>=23`.
