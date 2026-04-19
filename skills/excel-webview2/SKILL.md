---
name: excel-webview2
description: Uses Excel WebView2 via MCP to debug, inspect, and automate an Excel add-in's embedded WebView2 browser. Use when inspecting or interacting with a running Excel add-in's task pane or web content. Requires the add-in to already be running with remote debugging enabled on port 9222.
---

## Prerequisites

This skill connects to an **already-running** Excel add-in WebView2 instance. The MCP server does **not** launch Chrome or create a browser — it attaches to the existing WebView2 remote debugging endpoint at `http://localhost:9222`.

Before using any tools, the user must have their add-in running locally with the WebView2 debug port enabled. Setup instructions live in one place:

[README.md#launching-excel-with-the-debug-port](../../README.md#launching-excel-with-the-debug-port)

If tools fail to connect, verify the debuggable target is available: `curl http://localhost:9222/json/version`

## Launching Excel from the agent

When the user is working inside an Excel add-in repository, this server can launch Excel directly instead of asking them to run an `npm run start:cdp`-style script in another terminal. Three lifecycle tools cover the flow:

- `excel_detect_addin` — inspects `cwd` (or a passed-in path) and reports whether it looks like an add-in repo. Detection signals: a `manifest.xml` (classic) or `manifest.json` (unified) at/above the working directory, plus signals from `package.json` (`office-addin-debugging` devDep, `--remote-debugging-port` in any script). Returns the detected `manifestPath`, `manifestKind`, `packageManager`, and any existing CDP-enabled script.
- `excel_launch_addin` — spawns `office-addin-debugging start <manifest>` with `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS=--remote-debugging-port=<port>` injected into the child process env. Polls the CDP endpoint until ready, then (by default, `autoConnect: true`) calls the connect path so subsequent tools work without an extra step. **Idempotent per manifest** — a second call returns the existing tracked launch instead of spawning a duplicate.
- `excel_stop_addin` — runs `office-addin-debugging stop <manifest>` against the tracked launch (or all tracked launches when no `manifestPath` is given). Falls back to killing the child if stop does not exit cleanly.

### Env-var contract

The launcher only sets `WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS` in the spawned child's environment — it does not mutate the user's shell. If that variable is **already** set in `process.env` and contains `--remote-debugging-port`, the launch refuses with `port-already-configured` rather than silently overwriting. Tell the user to unset it (or just run the tool without the manual env var).

### When to suggest the launch tools

Suggest `excel_launch_addin` when:

- Connection tools fail because no CDP endpoint is up, AND
- `excel_detect_addin` confirms the working directory is an Excel add-in repo.

Do NOT suggest the launch tools on macOS/Linux — WebView2 is Windows-only and the launcher refuses with a platform error.

The user's existing `start:cdp` npm script remains a valid manual alternative; the launch tools simply move that step into the agent loop. If the user prefers running the script themselves, just connect to the resulting endpoint.

### Auto-launch at server startup

The `--auto-launch` CLI flag runs `excel_launch_addin` once during MCP server startup if the working directory is detected as an add-in repo. `--launch-port` (default 9222) and `--launch-timeout` (default 60000ms) control the launch.

## Core Concepts

**Connection**: The MCP server attaches to the WebView2 instance via the Chrome DevTools Protocol (CDP) at `http://localhost:9222`. No browser is launched.

**Page selection**: WebView2 may expose multiple debuggable targets (task pane, dialog, etc.). Use `list_pages` to see available targets, then `select_page` to switch context.

**Element interaction**: Use `take_snapshot` to get page structure with element `uid`s. Each element has a unique `uid` for interaction. If an element isn't found, take a fresh snapshot — the element may have been removed or the page changed.

## Workflow Patterns

### Starting a session

1. `list_pages` — confirm the WebView2 target is visible and select the right one
2. `take_snapshot` — understand the current page structure before interacting

### Before interacting with a page

1. Snapshot: `take_snapshot` to understand page structure
2. Interact: Use element `uid`s from snapshot for `click`, `fill`, etc.
3. Wait: `wait_for` if an action triggers async UI updates

**Note**: The MCP server exposes no navigation or tab-creation tools — the add-in host owns its URL and window. Attach to an already-running add-in and interact with the pages it exposes.

**For Office.js-aware debugging** (inspecting the selected range, reading `Office.context`, handling dialogs, probing requirement sets), see the [`excel-addin-debugging`](../excel-addin-debugging/SKILL.md) skill.

### Verifying workbook state after add-in code changes

The `excel_*` read tools run inside the add-in page via `Excel.run` and return structured JSON. They are all read-only (no mutation) and cap grid payloads at 1000 cells (`truncated: true` when exceeded). Use them to confirm that a code change produced the expected workbook state rather than only relying on UI screenshots.

- **Workbook surface**: `excel_context_info`, `excel_workbook_info`, `excel_list_worksheets`, `excel_worksheet_info`, `excel_calculation_state`, `excel_list_named_items`, `excel_custom_xml_parts`, `excel_settings_get`.
- **Ranges**: `excel_active_range` (current selection), `excel_read_range` (by A1 address or `Sheet!A1:C10`), `excel_used_range`, `excel_range_formulas` (A1 + R1C1 side-by-side with values), `excel_range_properties` (valueTypes, hidden flags, optional font/fill/alignment), `excel_range_special_cells` (constants/formulas/blanks/visible), `excel_find_in_range`.
- **Tables**: `excel_list_tables`, `excel_table_info`, `excel_table_rows`, `excel_table_filters`.
- **PivotTables**: `excel_list_pivot_tables`, `excel_pivot_table_info`, `excel_pivot_table_values`.
- **Charts & shapes**: `excel_list_charts`, `excel_chart_info`, `excel_chart_image` (base64 PNG), `excel_list_shapes`.
- **Validation & formatting rules**: `excel_list_conditional_formats`, `excel_list_data_validations`, `excel_list_comments`.

Range-targeting tools accept either `{sheet, address}` or a fully qualified `'Sheet1!A1:C10'` address; omit the address to operate on the active selection. When a tool returns `{error: "Excel API not available on this target"}`, the selected page is not an Excel add-in context — re-run `list_pages` / `select_page`.

### Efficient data retrieval

- Use `filePath` parameter for large outputs (screenshots, snapshots, traces)
- Use pagination (`pageIdx`, `pageSize`) and filtering (`types`) to minimize data
- Set `includeSnapshot: false` on input actions unless you need updated page state

### Tool selection

- **Automation/interaction**: `take_snapshot` (text-based, faster, better for automation)
- **Visual inspection**: `take_screenshot` (when user needs to see visual state)
- **Additional details**: `evaluate_script` for data not in accessibility tree

### Parallel execution

You can send multiple tool calls in parallel, but maintain correct order: snapshot → interact → wait (if needed) → snapshot again.

## Troubleshooting

**Cannot connect / no pages listed**: The add-in is not running, or remote debugging is not enabled on port 9222. Ask the user to verify their add-in is started and the debug port is configured.

**Element not found**: Take a fresh snapshot — the add-in may have re-rendered the page.

**Unexpected navigation**: The add-in controls routing internally. If the taskpane changes URL on its own, re-run `list_pages` and `select_page` to re-attach to the new target.

For DevTools protocol reference: https://developer.chrome.com/docs/devtools
