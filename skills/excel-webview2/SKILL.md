---
name: excel-webview2
description: Uses Excel WebView2 via MCP to debug, inspect, and automate an Excel add-in's embedded WebView2 browser. Use when inspecting or interacting with a running Excel add-in's task pane or web content. Requires the add-in to already be running with remote debugging enabled on port 9222.
---

## Prerequisites

This skill connects to an **already-running** Excel add-in WebView2 instance. The MCP server does **not** launch Chrome or create a browser — it attaches to the existing WebView2 remote debugging endpoint at `http://localhost:9222`.

Before using any tools, the user must:

1. Have their Excel add-in loaded and running in Excel
2. Have remote debugging enabled on port 9222 (configured in the add-in host or launch settings)

If tools fail to connect, verify the debuggable target is available: `curl http://localhost:9222/json`

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
