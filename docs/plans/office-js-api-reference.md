# Office.js API Reference (Phase 0 Discovery)

**Purpose:** Authoritative list of Office.js / Excel JavaScript API members that later phases (3, 4, 5, 6) are allowed to depend on. Every entry was verified against Microsoft's published reference via context7 on 2026-04-18. If an API later phases need is **not** in this file, the phase must pause and re-run discovery rather than guessing.

**Primary context7 libraries consulted**

- `/officedev/office-js-docs-reference` — formal API reference (3229 snippets, benchmark 50.7).
- `/officedev/office-js-docs-pr` — concept / how-to docs (2278 snippets, benchmark 75.58).
- `/websites/learn_microsoft_en-us_office_dev_add-ins` — learn.microsoft.com mirror (10186 snippets, benchmark 71.65).

All source URLs below are the pages returned by context7.

---

## 1. Readiness and global detection

### 1.1 `Office.onReady(callback?)`

- **Shape:** `Office.onReady(): Promise<{ host: Office.HostType; platform: Office.PlatformType }>` — also accepts a callback with the same info object.
- **Min requirement set:** Common 1.1 (always present once Office.js is loaded).
- **Source:** https://learn.microsoft.com/en-us/office/dev/add-ins/develop/initialize-add-in
- **Snippet (copy-ready):**

  ```js
  Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel) {
      // Excel-specific boot
    }
    if (info.platform === Office.PlatformType.PC) {
      // Windows-specific boot
    }
    console.log(`Office.js ready in ${info.host} on ${info.platform}`);
  });
  ```

- **Notes:** If the Office.js library is blocked (firewall, extension), the promise never resolves — tools must time-box any `await Office.onReady()`.

### 1.2 `Office.initialize` (legacy)

- **Shape:** `Office.initialize = function(reason?) { ... }`
- **Source:** https://learn.microsoft.com/en-us/office/dev/add-ins/develop/initialize-add-in
- **Notes:** Predecessor to `Office.onReady`. Still supported. Do not _require_ it in new tooling; prefer `Office.onReady`. Presence alone is a weak signal — `typeof Office !== 'undefined'` is stronger.

### 1.3 Global detection (for tool handlers)

No Microsoft-published "is Office.js loaded" helper exists. The canonical pattern from the tutorials is:

```js
if (typeof Office !== 'undefined' && Office.context && Office.context.host) {
  // Office.js is live
}
if (typeof Excel !== 'undefined') {
  // Excel namespace available
}
```

**Source inference:** Combined from the Excel tutorial and `Office.context` property docs (both URLs below). No single Microsoft page says this verbatim, so tools should wrap any such probe inside a `try/catch` and treat `ReferenceError` as "not loaded."

---

## 2. `Office.context` properties

Base URL: https://learn.microsoft.com/en-us/office/dev/add-ins/reference/javascript-api-for-office
Context7 source: https://github.com/officedev/office-js-docs-reference/blob/main/docs/requirement-sets/outlook/requirement-set-1.10/office.context.md (the Outlook requirement set page also documents the shared `Office.context` surface).

| Property                         | Type                       | Min req. set | Notes                                                         |
| -------------------------------- | -------------------------- | ------------ | ------------------------------------------------------------- |
| `Office.context.host`            | `Office.HostType` enum     | Common 1.1   | e.g. `"Excel"`, `"Word"`, `"Outlook"`.                        |
| `Office.context.platform`        | `Office.PlatformType` enum | Common 1.1   | e.g. `"PC"`, `"Mac"`, `"OfficeOnline"`, `"iOS"`, `"Android"`. |
| `Office.context.contentLanguage` | `string`                   | Common 1.1   | RFC-1766 tag (`"en-US"`). User's **editing** language.        |
| `Office.context.displayLanguage` | `string`                   | Common 1.1   | RFC-1766 tag. User's **UI** language.                         |
| `Office.context.diagnostics`     | `ContextInformation`       | Common 1.1   | `{ host, version, platform }`.                                |
| `Office.context.requirements`    | `RequirementSetSupport`    | Common 1.1   | Provides `isSetSupported(name, version?)`.                    |
| `Office.context.ui`              | `UI`                       | Common 1.1   | Provides `displayDialogAsync`, `messageParent`.               |

**Copy-ready snippet:**

```js
const info = {
  host: Office.context.diagnostics.host,
  platform: Office.context.diagnostics.platform,
  version: Office.context.diagnostics.version,
  contentLanguage: Office.context.contentLanguage,
  displayLanguage: Office.context.displayLanguage,
};
```

---

## 3. `Office.HostType` and `Office.PlatformType` enums

- **Source:** https://learn.microsoft.com/en-us/office/dev/add-ins/develop/initialize-add-in
- **HostType values (verified in docs):** `Excel`, `Word`, `PowerPoint`, `Outlook`, `OneNote`, `Project`.
- **PlatformType values (verified in docs):** `PC`, `Mac`, `OfficeOnline`, `iOS`, `Android`.
- **Usage:** Compare stringly via enum members, e.g. `info.host === Office.HostType.Excel`. Do not compare to raw strings — the enum member is the source of truth even if it currently equals `"Excel"`.

---

## 4. Requirement set probing

### 4.1 `Office.context.requirements.isSetSupported(name, minVersion?)`

- **Shape:** `isSetSupported(name: string, minVersion?: string): boolean`
- **Source:** https://learn.microsoft.com/en-us/office/dev/add-ins/develop/specify-api-requirements-runtime
- **Notes:**
  - First param (name) is required; second (version) is optional and defaults to `"1.1"`.
  - Returns `true` if the current host supports _at least_ that version.
  - Never throws — returns `false` for unknown names.

**Copy-ready snippet:**

```js
if (Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
  // safe to use Excel API 1.7 members
}
```

### 4.2 Requirement-set names verified via docs

The following set names are referenced directly by Microsoft docs and safe to probe. Versions past these are valid but not separately enumerated here — `isSetSupported` returns `false` for unavailable versions.

| Set name               | Relevant to Excel add-in?         | Source                                                                                                                         |
| ---------------------- | --------------------------------- | ------------------------------------------------------------------------------------------------------------------------------ |
| `ExcelApi`             | Yes — primary Excel API           | https://learn.microsoft.com/en-us/office/dev/add-ins/reference/requirement-sets/excel-api/excel-api-requirement-sets           |
| `ExcelApiOnline`       | Yes — Excel on the web only       | https://learn.microsoft.com/en-us/office/dev/add-ins/develop/platform-specific-requirement-sets                                |
| `SharedRuntime`        | Yes — shared runtime features     | https://github.com/officedev/office-js-docs-reference/blob/main/docs/requirement-sets/common/office-add-in-requirement-sets.md |
| `DialogApi`            | Yes — `displayDialogAsync`        | https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins                                      |
| `RibbonApi`            | Yes — programmatic ribbon updates | https://learn.microsoft.com/en-us/office/dev/add-ins/reference/requirement-sets/ribbon-api/ribbon-api-requirement-sets         |
| `IdentityAPI`          | Yes — SSO                         | https://learn.microsoft.com/en-us/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets                  |
| `WordApi`              | No — Word host                    | N/A for Excel tooling                                                                                                          |
| `OutlookApi`/`Mailbox` | No — Outlook host                 | N/A for Excel tooling                                                                                                          |

**For Phase 3 `excel_context_info`, probe this list and include any that return `true`:**

```js
const sets = [
  ['ExcelApi', '1.1'],
  ['ExcelApi', '1.7'],
  ['ExcelApi', '1.10'],
  ['ExcelApi', '1.12'],
  ['ExcelApi', '1.14'],
  ['ExcelApiOnline', '1.1'],
  ['SharedRuntime', '1.1'],
  ['DialogApi', '1.1'],
  ['DialogApi', '1.2'],
  ['RibbonApi', '1.1'],
  ['IdentityAPI', '1.3'],
];
```

---

## 5. `Excel.run(batch)` and the `RequestContext` pattern

### 5.1 `Excel.run(async ctx => ...)`

- **Shape:** `Excel.run<T>(batch: (ctx: Excel.RequestContext) => Promise<T>): Promise<T>`
- **Min requirement set:** `ExcelApi 1.1`
- **Source:** https://context7.com/officedev/office-js-docs-reference/llms.txt (snippet: "Execute Excel Commands with Excel.run")
- **Rules:**
  - All property reads go through `range.load(...)` + `await ctx.sync()`. Accessing an unloaded property throws `PropertyNotLoaded`.
  - Never `load('*')` — specify explicit property names.
  - Return values extracted _after_ `ctx.sync()` from inside the callback. Do not `await` the `range` object outside the callback.

**Copy-ready snippet:**

```js
const result = await Excel.run(async ctx => {
  const sheet = ctx.workbook.worksheets.getActiveWorksheet();
  sheet.load('name');
  await ctx.sync();
  return sheet.name;
});
```

### 5.2 `ctx.workbook.getSelectedRange()`

- **Min requirement set:** `ExcelApi 1.1`
- **Source:** https://github.com/officedev/office-js-docs-reference/blob/main/docs/includes/excel-1_1.md
- **Returns:** `Excel.Range` representing the currently selected cells.

### 5.3 `ctx.workbook.worksheets.getActiveWorksheet()`

- **Min requirement set:** `ExcelApi 1.1`
- **Source:** same as above.

---

## 6. `Excel.Range` properties (loadable)

- **Source:** https://github.com/officedev/office-js-docs-reference/blob/main/docs/includes/excel-1_1.md
- **Min requirement set:** `ExcelApi 1.1` for everything listed unless noted.

| Property        | Type                       | Notes                                                           |
| --------------- | -------------------------- | --------------------------------------------------------------- |
| `address`       | `string`                   | A1-style, worksheet-qualified (e.g. `"Sheet1!A1:B3"`).          |
| `values`        | `any[][]`                  | 2-D raw values.                                                 |
| `formulas`      | `string[][]`               | 2-D A1 formulas.                                                |
| `formulasLocal` | `string[][]`               | 2-D locale-formatted formulas.                                  |
| `numberFormat`  | `string[][]`               | 2-D Excel format codes.                                         |
| `rowCount`      | `number`                   | Total rows in the range.                                        |
| `columnCount`   | `number`                   | Total columns in the range.                                     |
| `rowIndex`      | `number`                   | Zero-based row of the first cell.                               |
| `text`          | `string[][]`               | 2-D formatted strings (what the user sees).                     |
| `valueTypes`    | `Excel.RangeValueType[][]` | Per-cell type.                                                  |
| `worksheet`     | `Excel.Worksheet`          | Parent worksheet (load nested: `range.worksheet.load('name')`). |

**Copy-ready snippet (the Phase 4 pattern, verbatim-safe):**

```js
await Excel.run(async ctx => {
  const range = ctx.workbook.getSelectedRange();
  range.load(['address', 'values', 'rowCount', 'columnCount']);
  if (includeFormulas) range.load('formulas');
  if (includeNumberFormat) range.load('numberFormat');
  await ctx.sync();
  return {
    address: range.address,
    values: range.values,
    rowCount: range.rowCount,
    columnCount: range.columnCount,
    formulas: includeFormulas ? range.formulas : undefined,
    numberFormat: includeNumberFormat ? range.numberFormat : undefined,
  };
});
```

---

## 7. Named items (named ranges)

- **Source:** https://github.com/officedev/office-js-docs-reference/blob/main/docs/includes/excel-1_1.md
- **Min requirement set:** `ExcelApi 1.1` (basic get), `ExcelApi 1.4` (add with comment/scope).

### 7.1 Reading a named range

```js
await Excel.run(async ctx => {
  const named = ctx.workbook.names.getItem('MyNamedRange');
  const range = named.getRange();
  range.load(['address', 'values']);
  await ctx.sync();
  return {address: range.address, values: range.values};
});
```

- `workbook.names` — `Excel.NamedItemCollection`.
- `names.getItem(name: string)` — throws if name missing unless you use `getItemOrNullObject` (ExcelApi 1.4+).

---

## 8. Dialog API

### 8.1 `Office.context.ui.displayDialogAsync(startAddress, options?, callback?)`

- **Source:** https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins
- **Min requirement set:** `DialogApi 1.1`
- **Rules:**
  - `startAddress` must be HTTPS and same domain as the host page (WebView2 enforces this).
  - Opens a new CDP target — debugging tools must call `list_pages` again after the dialog opens.

**Copy-ready snippet:**

```js
Office.context.ui.displayDialogAsync(
  'https://localhost:3000/dialog.html',
  {height: 40, width: 30},
  result => {
    const dialog = result.value;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, arg => {
      console.log('Message from dialog:', arg.message);
    });
  },
);
```

### 8.2 `Office.context.ui.messageParent(message, options?)` (called _inside_ the dialog)

- **Source:** same page.
- **Shape:** `messageParent(message: string, options?: { targetOrigin: string })`
- **Notes:** Cross-domain dialogs require `targetOrigin`. `"*"` is allowed for non-sensitive messages.

---

## 9. WebView2 / Excel desktop specifics

### 9.1 Runtime

- Source: https://learn.microsoft.com/en-us/office/dev/add-ins/resources/resources-glossary and https://learn.microsoft.com/en-us/office/dev/add-ins/concepts/browsers-used-by-office-web-add-ins
- On Windows + Microsoft 365 (Version 2101+), Excel desktop runs add-ins inside **Edge WebView2** (`msedgewebview2.exe`).
- On Excel Online, the add-in runs inside the user's regular browser — there is **no** WebView2 process, and remote debugging on port 9222 is not applicable.
- `Office.context.diagnostics.platform === "PC"` combined with the WebView2 host string is the strongest signal that we are attached to Excel desktop's WebView2 instance.

### 9.2 Remote debugging

- Context7 did not return a single canonical Microsoft article about enabling CDP on Excel's WebView2 from the plan's perspective. The VS Code guidance page (https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-desktop-using-edge-chromium) documents the `msedge` / `useWebView: true` debugger attach pattern on port 9229, but that is VS Code's debugger. The server's own README is the authority for how to enable the port-9222 endpoint this project uses.
- **Status:** VERIFIED — WebView2 runtime presence. **NOT VERIFIED via context7 — port-9222 enablement. Defer to the server's README.**

### 9.3 Known differences: Excel Online vs. Excel desktop WebView2

- Source: https://learn.microsoft.com/en-us/office/dev/add-ins/develop/platform-specific-requirement-sets
- `ExcelApiOnline` is present only on Excel on the web. Probing it is the cleanest boolean for "am I in Online?"
- Shared runtime (`SharedRuntime`) is Windows-only for Excel — Online does not offer it.
- Dialog API behavior: on Online, dialogs open as iframe modals, not OS windows, and are not independent CDP targets.

---

## 10. Anti-patterns (enforce in every phase)

1. **No invented properties.** If a property isn't in sections 1–9 above, do not call it. Go back to context7 first.
2. **No `load('*')`.** Excel throws. Always enumerate.
3. **No property reads outside `Excel.run` / after `ctx.sync()`.** Reading a loaded property after the callback resolves is safe; reading _without_ `await ctx.sync()` will throw.
4. **No assumption that `Office` or `Excel` globals exist.** The WebView2 may have navigated, or the target may be a dialog/iframe.
5. **No reliance on Excel Online-only or desktop-only behavior.** Branch on `Office.context.platform` or on `isSetSupported('ExcelApiOnline', '1.1')`.

---

## 11. Entries marked NOT VERIFIED

- **Port-9222 CDP enablement on Excel WebView2** — context7 did not surface the canonical Microsoft article. Use the server's README instead.
- **Complete `Office.HostType` / `Office.PlatformType` enum exhaustiveness** — context7 returned only the commonly-cited members. Do not assume other members exist; probe defensively.

---

## 12. Change log

- 2026-04-18 — Initial authoring. All sections 1–9 verified via context7.
