# Connection Robustness — Reference Findings (Phase 0)

Verified facts with source URLs. Created 2026-04-19 per Phase 0 of
[connection-robustness-plan.md](./connection-robustness-plan.md). This document
contains research only — no implementation.

Claims flagged **NOT VERIFIED** must not be relied on in implementation code or
user-facing docs.

---

## 1. Chrome DevTools Protocol HTTP discovery endpoints

Source: <https://chromedevtools.github.io/devtools-protocol/>

### `GET /json/version`

- Returns browser/version metadata as JSON.
- Documented response fields:
  - `Browser` (string) — Chrome/Edge version string.
  - `Protocol-Version` (string) — CDP version.
  - `User-Agent` (string).
  - `V8-Version` (string).
  - `WebKit-Version` (string).
  - `webSocketDebuggerUrl` (string) — browser-level WebSocket endpoint Puppeteer
    resolves to when given `browserURL`.
- Served on the same host:port that was passed as `--remote-debugging-port`.
- Standard success status: `200 OK`. No authentication. CORS headers are present
  but not required for server-side `fetch`.

### `GET /json` (alias `GET /json/list`)

- Returns an array of available debuggable targets.
- Per-target fields documented: `description`, `devtoolsFrontendUrl`, `id`,
  `title`, `type` (e.g. `"page"`), `url`, `webSocketDebuggerUrl`.
- Used by Puppeteer after `/json/version` to enumerate pages.

### Implication for the probe

`/json/version` is the minimal liveness check. A 2xx response with a parseable
JSON body containing a `Browser` string is sufficient evidence that a CDP host
is reachable. Any other outcome (network error, timeout, non-2xx, non-JSON)
means `puppeteer.connect` will also fail.

---

## 2. Puppeteer `connect` / disconnect semantics

Source: <https://github.com/puppeteer/puppeteer/blob/main/docs/api/puppeteer.connectoptions.md>
Source: <https://github.com/puppeteer/puppeteer/blob/main/docs/api/puppeteer.browserevent.md>
Source: <https://github.com/puppeteer/puppeteer/blob/main/docs/api/puppeteer.connect.md>
Source: <https://github.com/puppeteer/puppeteer/blob/main/docs/guides/browser-management.md>

### `ConnectOptions` fields relevant to this plan

- `browserURL` (string, optional) — HTTP URL of an existing browser. Puppeteer
  fetches `/json/version` internally to resolve the WebSocket endpoint.
- `browserWSEndpoint` (string, optional) — direct WebSocket URL. Skips HTTP
  discovery.
- `headers` (Record<string,string>, optional) — headers on the WebSocket
  connection. Node-only.
- `protocolTimeout` (number, optional, default `180000` ms) — timeout for a
  single CDP call. Not a connect-phase timeout.
- `targetFilter` (TargetFilterCallback, optional) — allows filtering which
  targets Puppeteer attaches to.
- `acceptInsecureCerts`, `defaultViewport`, `slowMo`, `protocol`, `channel`,
  `transport`, `handleDevToolsAsPage`, `networkEnabled`, `issuesEnabled`,
  `downloadBehavior`, `capabilities` — present but not load-bearing for this
  plan.

### `browser.on('disconnected')`

- Emitted when Puppeteer loses its connection to the browser instance. Documented
  triggers: browser crashes, browser closes, or explicit
  `browser.disconnect()` call.
- No documented payload on the event.
- `browser.disconnect()` detaches Puppeteer without closing the browser (unlike
  `browser.close()`).

### Connect example (canonical shape)

```ts
const browser = await puppeteer.connect({
  browserWSEndpoint: 'ws://127.0.0.1:9222/devtools/browser/...',
});
```

### Not explicitly documented (flagged)

- **NOT VERIFIED** — Whether `disconnected` fires when Excel merely _pauses_ a
  WebView2 (task pane hidden but process alive) vs. only when the process/WS
  tears down. Treat pause behavior as unknown; the plan's reconnect budget
  (Phase 3) is the safety net.
- **NOT VERIFIED** — Whether `puppeteer.connect` with a `browserURL` emits any
  retry of its own when `/json/version` returns 5xx. Assume it does not and
  implement retries explicitly (Phase 2).

---

## 3. Enabling the WebView2 remote debugging port for Office add-ins

### Canonical Microsoft-documented mechanism (Windows)

Source (WebView2 team): <https://learn.microsoft.com/en-us/microsoft-edge/webview2/how-to/debug-visual-studio-code>

WebView2 exposes a CDP endpoint when the host process is launched with either
of these equivalent mechanisms:

1. **Environment variable** (per-process, preferred for local dev):

   ```
   WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222
   ```

   The Office application must be launched _after_ the variable is set and must
   inherit it. Applies to all WebView2 instances in that process.

2. **Registry value** (machine-wide, persistent):
   - Key: `HKEY_CURRENT_USER\Software\Policies\Microsoft\Edge\WebView2\AdditionalBrowserArguments`
   - Value name: the executable file name, e.g. `EXCEL.EXE`.
   - Value data: `--remote-debugging-port=9222` (REG_SZ).

Both are documented by the Edge/WebView2 team as interchangeable ways to pass
additional browser arguments into the WebView2 runtime.

### Office-add-in specific confirmation

Source (Office add-ins team): <https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-chromium>

The Office add-ins team documents `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` as
the supported channel for passing flags into the WebView2 runtime that hosts an
Office add-in task pane — the article uses
`--auto-open-devtools-for-tabs` as its worked example, but the mechanism is
generic: any Chromium/WebView2 browser argument can be passed through this
variable. `--remote-debugging-port=9222` follows the same pattern.

> "Set the `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` environment variable to
> include the value ... Open the Office application. Run the add-in."

Combined, these two Microsoft-authored docs establish
`WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=9222` as the
supported path for this project's use case on Windows.

### Platform split

- **Windows (Excel desktop)**: WebView2-based; the env var / registry
  mechanism above applies. This is the only platform this MCP server targets.
- **Mac (Excel for Mac)**: Does _not_ use WebView2. Microsoft documents
  debugging via the Safari Web Inspector instead, with no CDP port equivalent.
  Source: <https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-add-ins-overview>
  (see the "Debug on Mac" section). Users on Mac cannot use this MCP server;
  Phase 6 docs should say so explicitly.
- **Linux / Office on the web**: No local CDP surface; out of scope.

### Caveats

- The `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` variable must be set in the
  parent process that launches Excel and must be inherited. Setting it in a
  cmd window that didn't spawn Excel has no effect.
- The CDP port is shared across all WebView2 controls in the host process.
  The VS Code WebView2 doc notes: "After the first match is found in the URL,
  the debugger stops. You cannot debug two WebView2 controls at the same
  time, because the CDP port is shared by all WebView2 controls, and uses a
  single port number." Implication: if Excel hosts multiple WebView2 task
  panes, `/json/list` returns all of them and the consumer must pick.
- **NOT VERIFIED** — Whether group-policy restrictions can block the
  `Software\Policies\Microsoft\Edge\WebView2\AdditionalBrowserArguments` key
  from taking effect on managed enterprise machines. Phase 6 docs should
  recommend the env var first and mention the registry key as a fallback.

---

## 4. Summary table for implementation phases

| Need (from plan)                       | Verified source                                                                             |
| -------------------------------------- | ------------------------------------------------------------------------------------------- |
| `/json/version` response shape         | chromedevtools.github.io/devtools-protocol/                                                 |
| `browser.on('disconnected')` semantics | puppeteer docs — `puppeteer.browserevent.md`                                                |
| `puppeteer.connect` option names       | puppeteer docs — `puppeteer.connectoptions.md`                                              |
| Debug-port env var (Windows)           | learn.microsoft.com — webview2/how-to/debug-visual-studio-code                              |
| Debug-port registry key (Windows)      | learn.microsoft.com — webview2/how-to/debug-visual-studio-code                              |
| Office-add-in confirmation of env var  | learn.microsoft.com — office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-chromium |
| Mac uses Safari Web Inspector (no CDP) | learn.microsoft.com — office/dev/add-ins/testing/debug-add-ins-overview                     |

---

## 5. Items explicitly NOT VERIFIED

Do not cite these in code comments or user-facing docs without further
research:

- Whether `browser.on('disconnected')` fires on WebView2 task-pane hide/pause
  vs. only on process/WS tear-down.
- Whether `puppeteer.connect({browserURL})` performs any internal retry on
  transient `/json/version` failures.
- Whether managed/enterprise group policy can suppress the
  `AdditionalBrowserArguments` registry key.
- Whether there is any supported way to enable a CDP port on Excel for Mac.
