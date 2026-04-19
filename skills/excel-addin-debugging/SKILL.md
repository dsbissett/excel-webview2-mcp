---
name: excel-addin-debugging
description: Debug and automate Excel add-ins running in WebView2 with Office.js awareness — active range inspection, Office context detection, dialog API handling, and requirement-set probing. Use when the user mentions Office.js, Excel.run, taskpane, ribbon command, named ranges, or an add-in dialog.
---

## When to use

Trigger this skill when any of the following apply:

- The user mentions **Office.js**, `Office.context`, `Excel.run`, or requirement sets (`ExcelApi`, `DialogApi`, etc.).
- The user is debugging a **taskpane**, **ribbon command**, or **add-in dialog** running inside Excel's WebView2.
- The user asks about **the currently selected range**, **named ranges**, or **active worksheet state**.
- The target at `http://localhost:9222` is an Excel add-in rather than a generic web page.

For generic DOM/CDP interaction against the WebView2 target (snapshot, click, screenshot), defer to the `excel-webview2` skill.

## Readiness check (always run first)

Call [`excel_context_info`](../../src/tools/excel.ts) **before** any Office.js-aware tool. If it returns `hasOfficeGlobal: false`, stop and report the target is not an Excel add-in — subsequent Office.js tools will fail.

The tool returns:

- `hasOfficeGlobal`, `hasExcelGlobal` — presence of the runtime globals.
- `hostInfo` — `{ host, platform, version }` from `Office.context.diagnostics`.
- `contentLanguage`, `displayLanguage` — locale info.
- `requirementSets` — which `ExcelApi` / `DialogApi` / `SharedRuntime` / `ExcelApiOnline` / `RibbonApi` / `IdentityAPI` versions the host supports.

Branch behavior on `requirementSets`. For example, only call `getItemOrNullObject` patterns if `ExcelApi 1.4` is reported.

## Common workflows

### Inspect the current selection

1. `excel_context_info` — confirm `hasExcelGlobal: true`.
2. `excel_active_range` — returns `{ address, values, rowCount, columnCount, formulas?, numberFormat? }`.
3. For ranges over 1000 cells, the tool truncates `values` and notes the truncation in the response. Narrow the selection or read a named sub-range instead.

### Read a named range

Use `evaluate_script` with the pattern from [references/office-js-cheatsheet.md](references/office-js-cheatsheet.md) §3. Always guard with `typeof Excel === 'undefined'` and wrap in `Excel.run`.

### Capture taskpane state after a ribbon click

1. Ask the user to invoke the ribbon command.
2. `list_pages` — a new CDP target may have appeared (shared-runtime commands reuse the taskpane target; UI-less commands may not expose one).
3. `take_snapshot` on the taskpane target to see post-command DOM.
4. `list_console_messages` to surface any Office.js errors from the command handler.

### Dialog API flow

`Office.context.ui.displayDialogAsync` opens a **new debuggable CDP target**. After the add-in invokes it:

1. `list_pages` again — the dialog appears as a separate target.
2. `select_page` to switch into the dialog.
3. Interact via `take_snapshot` / `click` / `evaluate_script` as usual.
4. The dialog posts messages back via `Office.context.ui.messageParent` (see cheatsheet §7); the parent handles them through `Office.EventType.DialogMessageReceived`.

The dialog's `startAddress` must be HTTPS and same-origin with the taskpane. On Excel Online the dialog is an iframe modal — it is not a separate CDP target there, but this skill targets Excel desktop WebView2.

### Requirement-set branching

```js
// Inside evaluate_script
if (Office.context.requirements.isSetSupported('ExcelApi', '1.12')) {
  // Use Excel API 1.12 features
} else {
  // Fall back or report unsupported
}
```

Prefer probing over version-string comparison against `diagnostics.version`.

## Anti-patterns

- **Never invent Office.js APIs.** Only call members listed in [references/office-js-cheatsheet.md](references/office-js-cheatsheet.md). That file is sourced from [docs/plans/office-js-api-reference.md](../../docs/plans/office-js-api-reference.md) — if you need something not in the cheatsheet, add it there first (with a source URL) rather than guessing.
- **Never `load('*')`.** Excel throws `PropertyNotLoaded`. Enumerate property names explicitly.
- **Never read loaded properties without `await ctx.sync()`** inside `Excel.run`.
- **Never assume `Office` or `Excel` globals exist** — always guard with `typeof`.
- **Do not duplicate generic CDP debugging advice** already in the `excel-webview2` skill. Cross-link instead.

## Troubleshooting

**`hasOfficeGlobal: false`** — the attached target is not an add-in page. It may be the taskpane's sign-in iframe, a dialog, or the wrong tab. Run `list_pages` and `select_page` to switch.

**`Excel.run` throws `PropertyNotLoaded`** — a property was read without first being loaded. Add it to the `range.load([...])` call and re-sync.

**Dialog never appears in `list_pages`** — the add-in may be running on Excel Online (where dialogs are iframes, not CDP targets), or the dialog's HTTPS/same-origin check failed. Check `list_console_messages` for the Office.js error.

**Requirement set reports `false` unexpectedly** — confirm Excel version via `excel_context_info.hostInfo.version`. `ExcelApi` versions are tied to specific Microsoft 365 builds.
