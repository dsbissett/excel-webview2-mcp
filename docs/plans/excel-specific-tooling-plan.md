# Plan: Excel-Specific Tooling, Skills, and Tool-Surface Pruning

**Target:** Transform this repo from a rebranded Chrome DevTools MCP fork into a genuinely Excel-add-in-aware MCP server. Add Office.js-aware tools, a dedicated Excel skill, prune browser-only tools that don't apply to a WebView2 taskpane, and align eval scenarios.

**Author context:** Created 2026-04-18. Each phase is self-contained and can be executed in a fresh chat.

---

## Discovered Facts (copy into every phase's context)

Evidence-backed findings from discovery against the current repo. Cite these rather than re-discovering.

### Tool definition pattern

- Tools live in `src/tools/*.ts` and are defined via `definePageTool(...)` (page-scoped) or `defineTool(...)` imported from `./ToolDefinition.js`.
- Example shape: see [src/tools/snapshot.ts:12-41](../../src/tools/snapshot.ts#L12-L41) (`takeSnapshot`) and [src/tools/console.ts:40-86](../../src/tools/console.ts#L40-L86) (`listConsoleMessages`).
- Handler signature: `async (request, response, context) => { ... }`
  - `request.params` — zod-parsed params
  - `request.page.pptrPage` — the underlying Puppeteer `Page` (page-scoped tools only)
  - `context` — methods like `waitForTextOnPage`, `saveTemporaryFile`, `getSelectedMcpPage` (see [src/McpContext.ts](../../src/McpContext.ts) and the `Context` type in `ToolDefinition.ts:142-202`).
- Zod schema pattern: `{ paramName: zod.string().optional().describe('...') }`.
- Every tool sets `annotations.category` to a `ToolCategory` enum value.

### Tool registration

- Tools are imported and spread in [src/tools/tools.ts](../../src/tools/tools.ts) inside `createTools()`. To add a new tool module, add `import * as foo from './foo.js'` and `...Object.values(foo)` to the array (approx. lines 9-44).

### Categories

Existing categories in [src/tools/categories.ts:7-16](../../src/tools/categories.ts#L7-L16):
`INPUT`, `NAVIGATION`, `EMULATION`, `PERFORMANCE`, `NETWORK`, `DEBUGGING`, `EXTENSIONS`, `IN_PAGE`.

### License header (MANDATORY for every new .ts source file)

```
/**
 * @license
 * Copyright 2025 Google LLC
 * SPDX-License-Identifier: Apache-2.0
 */
```

ESLint rule `scripts/eslint_rules/check-license-rule.js` enforces this. Omitting it will fail `npm run format`.

### Docs generation

`docs/tool-reference.md` and `docs/slim-tool-reference.md` are **auto-generated** by `npm run gen`. Never edit them manually — let the generator produce them after new tools land.

### Skill structure

- `skills/<name>/SKILL.md` with YAML frontmatter: `name`, `description`.
- Optional `skills/<name>/references/` folder for supporting snippets.
- Example: [skills/excel-webview2/SKILL.md](../../skills/excel-webview2/SKILL.md) is the current entry point and already references tools like `list_pages`, `take_snapshot`.

### Test pattern

E2E pattern in `tests/e2e/` uses `node:test` with `runCli([...])` from `tests/utils.js`. Unit-test pattern for non-tool logic in `tests/*.test.ts`.

### Current tool inventory (classification)

| File             | Classification | Reason                                                                                                 |
| ---------------- | -------------- | ------------------------------------------------------------------------------------------------------ |
| `console.ts`     | KEEP           | Taskpane console logs are core debugging.                                                              |
| `input.ts`       | KEEP           | Click/fill/type/etc. apply directly.                                                                   |
| `snapshot.ts`    | KEEP           | Accessibility tree still useful.                                                                       |
| `screenshot.ts`  | KEEP           | Visual debugging.                                                                                      |
| `script.ts`      | KEEP           | `evaluate_script` is essential.                                                                        |
| `network.ts`     | KEEP           | API/backend debugging.                                                                                 |
| `memory.ts`      | KEEP           | Heap snapshots useful.                                                                                 |
| `performance.ts` | KEEP           | Taskpane rendering perf.                                                                               |
| `inPage.ts`      | KEEP           | Custom add-in tools via `window.__dtmcp`.                                                              |
| `emulation.ts`   | PRUNE (mostly) | Viewport/UA/geolocation N/A; retain network/CPU throttle if feasible.                                  |
| `extensions.ts`  | PRUNE          | Chrome extensions don't exist in WebView2 taskpane context.                                            |
| `pages.ts`       | RETHINK        | Keep `listPages`/`selectPage` (multi-target), drop `newPage`/`navigatePage` — the add-in owns routing. |
| `lighthouse.ts`  | RETHINK        | Likely broken against a taskpane; verify or remove.                                                    |

### Anti-patterns to guard against

- **Do NOT invent Office.js APIs.** Every Office.js call must be verified against Microsoft's official reference before being embedded in a handler or skill snippet.
- **Do NOT hand-edit `docs/tool-reference.md` or `docs/slim-tool-reference.md`.**
- **Do NOT drop the license header.**
- **Do NOT assume `Office` or `Excel` globals are present.** The WebView2 may have navigated away from the add-in page, or the target may be an iframe. Always check before calling.

---

## Phase 0 — Documentation Discovery (EXECUTE FIRST) — ✅ COMPLETE (2026-04-18)

**Goal:** Produce an "Allowed Office.js APIs" reference that later phases copy from, so no handler contains an invented method.

**Deliverable:** [docs/plans/office-js-api-reference.md](office-js-api-reference.md) — **written and verified**. Later phases (3, 4, 5) MUST copy from this file and not call any Office.js member not listed there.

### Phase 0 completion summary

Context7 libraries consulted:

- `/officedev/office-js-docs-reference` (formal API reference).
- `/officedev/office-js-docs-pr` (concept / how-to docs).
- `/websites/learn_microsoft_en-us_office_dev_add-ins` (learn.microsoft.com mirror).

Verified surface covered in [office-js-api-reference.md](office-js-api-reference.md):

1. §1 Readiness & global detection — `Office.onReady`, legacy `Office.initialize`, `typeof Office`/`typeof Excel` probe pattern.
2. §2 `Office.context` properties — `host`, `platform`, `contentLanguage`, `displayLanguage`, `diagnostics`, `requirements`, `ui`.
3. §3 `Office.HostType` and `Office.PlatformType` enums (members verified: Excel/Word/PowerPoint/Outlook/OneNote/Project; PC/Mac/OfficeOnline/iOS/Android).
4. §4 `Office.context.requirements.isSetSupported(name, minVersion?)` — plus a verified probe list (`ExcelApi`, `ExcelApiOnline`, `SharedRuntime`, `DialogApi`, `RibbonApi`, `IdentityAPI`) that Phase 3 must use verbatim.
5. §5 `Excel.run` batched pattern + `ctx.workbook.getSelectedRange()` + `ctx.workbook.worksheets.getActiveWorksheet()`.
6. §6 `Excel.Range` loadable properties: `address`, `values`, `formulas`, `formulasLocal`, `numberFormat`, `rowCount`, `columnCount`, `rowIndex`, `text`, `valueTypes`, `worksheet`.
7. §7 Named items — `workbook.names.getItem(...)` + `getRange()`.
8. §8 Dialog API — `displayDialogAsync`, `messageParent`, HTTPS + same-origin rules.
9. §9 WebView2 specifics — Excel desktop uses `msedgewebview2.exe` on Windows + Microsoft 365 2101+; Excel Online does not.

Items explicitly **marked NOT VERIFIED via context7** (see §11 of the reference file):

- Canonical Microsoft article for enabling port-9222 CDP on Excel WebView2. Defer to the server's README.
- Complete enum exhaustiveness for `Office.HostType` / `Office.PlatformType`. Probe defensively; do not assume absent members exist.

**Downstream impact on this plan:** Phases 3 and 4 may now proceed. The Phase 3 requirement-set probe list in §4.2 of the reference file **supersedes** the bullet in Phase 3's spec ("probe the list defined in Phase 0"). The Phase 4 active-range snippet in §6 of the reference file **is** the handler body verbatim (modulo parameter names).

### Tasks

1. Via `mcp__plugin_context7_context7__resolve-library-id` with `libraryName: "officejs"` (or try `"office-js"`, `"office-add-ins"`), then `query-docs`, fetch docs for:
   - `Office.context` — detecting host app, platform, requirement sets
   - `Office.onReady` / `Office.initialize` — readiness signals
   - `Excel.run(ctx => ...)` — batched Excel API pattern
   - `Excel.Workbook` — getting active worksheet, selected range
   - `Excel.Range` — `.address`, `.values`, `.formulas`, `.numberFormat`
   - `Office.context.ui.displayDialogAsync` — dialog API
   - `Office.context.requirements.isSetSupported` — capability detection

2. Also fetch WebView2-specific facts:
   - How `Office.context.diagnostics` exposes platform/version.
   - Known difference between Excel Online and Excel desktop WebView2 host.

3. For each API captured, write one-line entries in `docs/plans/office-js-api-reference.md` with:
   - Exact method/property name and signature.
   - Minimum Office.js version or requirement set.
   - Source URL.
   - Copy-ready snippet if short.

### Verification

- File exists at `docs/plans/office-js-api-reference.md`.
- Every entry has a source URL (not "per my knowledge").
- `grep -n "TODO\|unclear\|not sure"` returns nothing.

### Anti-pattern guards

- If context7 doesn't return docs for a specific call, mark the entry "NOT VERIFIED — do not use" instead of guessing.
- Do **not** implement any tool in this phase. Discovery only.

---

## Phase 1 — Prune Inapplicable Tools — ✅ COMPLETE (2026-04-18)

**Goal:** Remove tools that don't apply to a single-taskpane WebView2, so the surface reflects reality.

### Phase 1 completion summary

**Tool-surface removals (user-visible):**

- `install_extension`, `uninstall_extension`, `list_extensions`, `get_extension_permissions`, `reload_extension` — deleted with `src/tools/extensions.ts`.
- `new_page`, `navigate_page`, `resize_page` — removed from [src/tools/pages.ts](../../src/tools/pages.ts). Kept: `list_pages`, `select_page`, `close_page`, `handle_dialog`, `get_tab_id`.
- `emulate` schema trimmed to `networkConditions` + `cpuThrottlingRate` in [src/tools/emulation.ts](../../src/tools/emulation.ts); viewport / userAgent / geolocation / colorScheme fields dropped.

**Internal plumbing kept intentionally:** `src/utils/ExtensionRegistry.ts`, `context.installExtension`, and `context.emulate` remain because `McpContext`, `McpResponse`, and `src/tools/script.ts` still import them. The plan's conditional ("if nothing else imports them") is false today, so deleting the registry would cascade across unrelated code. Dead-branch reads of `args.categoryExtensions` in `McpResponse.ts`, `pages.ts` description, and `script.ts` were left in place — they are now unreachable since the CLI flag is gone, and ripping them out belongs to a follow-up cleanup rather than Phase 1.

**Wiring updates:**

- [src/tools/tools.ts](../../src/tools/tools.ts) — dropped the `extensionTools` import + spread; added `as unknown as` cast in the tool-registration loop (required after the surviving union narrowed).
- [src/tools/categories.ts](../../src/tools/categories.ts) — removed `ToolCategory.EXTENSIONS` and its label. `EMULATION` kept (throttling still applies).
- [src/index.ts](../../src/index.ts) — removed the `ToolCategory.EXTENSIONS` filter branch; `enableExtensions` hard-coded to `false` when creating the browser.
- [src/bin/excel-webview2-mcp-cli-options.ts](../../src/bin/excel-webview2-mcp-cli-options.ts) — removed the `categoryExtensions` yargs option and its `conflicts` references on `autoConnect`, `browserUrl`, `wsEndpoint`.
- [src/bin/excel-webview2.ts](../../src/bin/excel-webview2.ts) — switched `commands` import from `./excel-webview2-cli-options.js` → `./cliDefinitions.js`; removed the `delete startCliOptions.categoryExtensions` line.
- [src/bin/excel-webview2-cli-options.ts](../../src/bin/excel-webview2-cli-options.ts) — **deleted**. It was a pre-existing duplicate of the auto-generated `cliDefinitions.ts`; only one consumer (`excel-webview2.ts`) referenced it, and it had gone stale after the Chrome→Excel rename.

**Docs regeneration:** `npm run gen` regenerated [docs/tool-reference.md](../tool-reference.md), [docs/slim-tool-reference.md](../slim-tool-reference.md), [README.md](../../README.md), [src/bin/cliDefinitions.ts](../../src/bin/cliDefinitions.ts), and [src/telemetry/tool_call_metrics.json](../../src/telemetry/tool_call_metrics.json) to reflect the pruned surface.

**Tests:**

- Deleted `tests/tools/pages.test.ts` (+ snapshot), `tests/tools/emulation.test.ts`, `tests/tools/extensions.test.ts` — each exercised a pruned tool surface via a real Chrome puppeteer launch; they have no Excel-WebView2 analogue.
- Edited [tests/tools/script.test.ts](../../tests/tools/script.test.ts) — removed `installExtension` import and the three extension/service-worker tests.
- Edited [tests/McpResponse.test.ts](../../tests/McpResponse.test.ts) — removed `navigatePage` / `newPage` imports and the two "includes in-page tools in navigate_page/new_page response" cases.
- Edited [tests/index.test.ts](../../tests/index.test.ts) — removed the `has experimental extensions tools` e2e case.

**Lighthouse decision:** **DEFERRED to Phase 6** per the plan's own guidance. `lighthouse.ts` and its e2e scenario remain; the empirical viability check against an actual Excel taskpane is Phase 6/7 work.

**Description scrub:** [src/tools/performance.ts](../../src/tools/performance.ts) — the `reload` schema's describe no longer mentions `navigate_page`. Two SKILL.md notes in [skills/excel-webview2/SKILL.md](../../skills/excel-webview2/SKILL.md) were rewritten to stop referencing `navigate_page` / `new_page`.

**Verification:**

- `npm run build` — clean.
- `npm run gen` — clean; auto-generated files regenerated without the pruned tools.
- `npm run test` — 14 pre-existing failures (baseline: 17). The 3 resolved failures correspond to the deleted pages/extensions tests. No _new_ failures were introduced by Phase 1. Pre-existing failures trace to unrelated branding gaps (e.g. `CHROME_DEVTOOLS_MCP_NO_UPDATE_CHECKS` env-var naming in `tests/check-for-updates.test.ts` and the `tests/cli.test.ts` default-args drift) and to Chrome-specific e2e tests (`tests/e2e/chrome-devtools-commands.test.ts`) that require a local Chrome install; these are not in Phase 1's scope.
- `grep -rn "install_extension\|new_page\|navigate_page\|resize_page" src/ tests/` — zero matches except the frozen `src/telemetry/tool_call_metrics.json` (historical call-frequency snapshot, not live code).

**Known follow-ups not in Phase 1 scope:**

- Other skills (`skills/troubleshooting/SKILL.md`, `skills/debug-optimize-lcp/SKILL.md`, `skills/memory-leak-debugging/SKILL.md`, `skills/excel-webview2-cli/SKILL.md`) still mention `navigate_page` / `new_page` / `resize_page`. They are inherited from the upstream Chrome DevTools fork and are candidates for pruning/rewrite in Phase 5 or a dedicated skills-audit pass.
- `scripts/eval_scenarios/*` still references pruned tools. Phase 6 deletes the inapplicable scenarios.
- Dead-branch reads of `args.categoryExtensions` in `McpResponse.ts`, `pages.ts`, `script.ts` should be collapsed in a later cleanup pass.

---

## Phase 1 — Prune Inapplicable Tools (original spec)

**Goal:** Remove tools that don't apply to a single-taskpane WebView2, so the surface reflects reality.

### What to implement

1. **Delete or gate `extensions.ts`.** Remove the import from [src/tools/tools.ts](../../src/tools/tools.ts). Delete `src/tools/extensions.ts` and `src/utils/ExtensionRegistry.ts` if nothing else imports them. Confirm with `grep -rn "ExtensionRegistry\|extensions\\.ts" src/`.
2. **Split `pages.ts`:** Keep `listPages`, `selectPage`, `handleDialog`, `getTabId`. Remove `newPage`, `navigatePage`, and `resizePage` exports. The taskpane's URL is owned by the add-in host, and the host controls the pane dimensions.
3. **Scope `emulation.ts`:** Keep only network/CPU throttling if present; remove device/viewport/userAgent/geolocation/colorScheme code paths. If the file becomes too thin, merge the surviving pieces into `performance.ts` instead.
4. **Decide on `lighthouse.ts`:** Run `lighthouse_audit` against a taskpane target in Phase 6 (verification). If it errors or returns no meaningful data, delete the file and its imports. If it works, keep. **Do not delete preemptively.**
5. Update [src/tools/categories.ts](../../src/tools/categories.ts): remove `EXTENSIONS` if fully pruned, remove `EMULATION` if that tool is fully collapsed.
6. Update [skills/excel-webview2/SKILL.md](../../skills/excel-webview2/SKILL.md) to remove references to any pruned tool.

### Documentation references

- Tool registration pattern: see `src/tools/tools.ts` `createTools()`.
- Category enum: `src/tools/categories.ts:7-16`.

### Verification checklist

- `npm run build` succeeds.
- `npm run test` passes (some pages/emulation tests may need updating — do NOT skip failures; fix or delete the test).
- `grep -rn "install_extension\|new_page\|navigate_page" src/ tests/` returns zero hits (except removal-tracking comments if any).
- `npm run gen` regenerates `docs/tool-reference.md` without the pruned tools.

### Anti-pattern guards

- Do not leave `// removed X tool` comments. If a tool is gone, it's gone.
- Do not rename pruned tools to variants — delete them cleanly.

---

## Phase 2 — Add `EXCEL` Category

**Goal:** Register a new tool category so Office/Excel-aware tools have a discoverable home.

### What to implement

1. In [src/tools/categories.ts](../../src/tools/categories.ts), add `EXCEL = 'excel'` to the `ToolCategory` enum and any matching label map. Preserve alphabetical ordering if the file is alphabetical.
2. If a slim-mode manifest (`src/tools/slim/tools.ts`) references categories, add `EXCEL` there too.
3. No new tools yet — this phase only opens the namespace.

### Documentation references

- Enum location: `src/tools/categories.ts:7-16`.
- Usage example: any existing tool setting `category: ToolCategory.DEBUGGING`.

### Verification checklist

- `npm run build` succeeds.
- `grep -n "EXCEL" src/tools/categories.ts` shows the new enum member.

### Anti-pattern guards

- Do not reuse `DEBUGGING` or `IN_PAGE` for Excel tools — the category must be discoverable.

---

## Phase 3 — Implement `excel_context_info` Tool

**Goal:** Deliver the smallest useful Excel-aware tool: return host/platform/requirement-set info so users can confirm they're attached to a real Excel add-in.

### What to implement

Create `src/tools/excel.ts` (new file, license header required) exporting one tool:

- `excelContextInfo` — `name: 'excel_context_info'`, `category: ToolCategory.EXCEL`.
- Handler runs an in-page `evaluate_script` equivalent that returns a plain JSON object with:
  - `hasOfficeGlobal: boolean` (is `typeof Office !== 'undefined'`)
  - `hasExcelGlobal: boolean`
  - `hostInfo` — from `Office.context.diagnostics` (host, platform, version) — only if `hasOfficeGlobal`.
  - `contentLanguage`, `displayLanguage` — from `Office.context` if available.
  - `requirementSets` — array of probed sets (e.g. `ExcelApi 1.10`) using `Office.context.requirements.isSetSupported(...)` — probe the list defined in Phase 0.
- Schema: no params.
- Response: `response.setIncludeSnapshot(false)`; write the JSON via `response.appendResponseLine(JSON.stringify(result, null, 2))`.

**Copy from:** [src/tools/console.ts:79-86](../../src/tools/console.ts#L79-L86) (handler shape) and [src/tools/snapshot.ts:60-73](../../src/tools/snapshot.ts#L60-L73) (using `request.page.pptrPage`).

### Documentation references

- Office.js APIs used must ALL appear in Phase 0's `office-js-api-reference.md`. If a call isn't there, stop and go back to Phase 0.
- Handler shape: `src/tools/snapshot.ts`.
- Registration: add `import * as excel from './excel.js';` and `...Object.values(excel),` in [src/tools/tools.ts](../../src/tools/tools.ts).

### Verification checklist

- `npm run build` succeeds.
- `npm run format` clean (license header present).
- `npm run gen` regenerates `docs/tool-reference.md` with `excel_context_info` present.
- Manual: connect to a running Excel add-in on port 9222, invoke the tool, and confirm the output reports `hasOfficeGlobal: true` and a non-empty `hostInfo`.
- If the attached target is NOT an add-in (e.g. a random web page), tool returns `hasOfficeGlobal: false` without throwing.

### Anti-pattern guards

- Handler must not crash when `Office` is undefined — wrap `evaluate_script` result and handle `undefined`/`ReferenceError`.
- Do not cache results across calls; the attached target could change.

---

## Phase 4 — Implement `excel_active_range` Tool

**Goal:** Return the current selection (address + values + formulas) using `Excel.run`.

### What to implement

In `src/tools/excel.ts`, add `excelActiveRange`:

- Schema: `{ includeFormulas?: boolean, includeNumberFormat?: boolean }`.
- Handler executes this in-page (via `pptrPage.evaluate`):

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

- Guard at the top: if `typeof Excel === 'undefined'`, return an error via `response.appendResponseLine('ERROR: Excel API not available on this target')` and return early.
- Cap output size: if `rowCount * columnCount > 1000`, truncate `values` and note truncation in the response.

### Documentation references

- `Excel.run` and `Range.load` signatures must come from Phase 0 reference file.
- Handler shape: copy pattern from `excel_context_info` in Phase 3.

### Verification checklist

- Unit-test the parameter schema via a new test file `tests/tools/excel.test.ts` if unit patterns exist; otherwise cover via an e2e scenario.
- Manual: select a 3x3 range in Excel, invoke tool, confirm address and 2-D values array returned.
- Select a 200x200 range, confirm truncation notice fires and response is bounded.

### Anti-pattern guards

- Do not `load('*')` — explicit property loading only (Office.js requirement).
- Do not `await` outside `Excel.run` — the returned values are only valid after `ctx.sync()`.
- Do not use `ctx.workbook.worksheets.getActiveWorksheet().getUsedRange()` as a fallback — that could return megabytes of data.

---

## Phase 5 — Excel-Specific Skill (`skills/excel-addin-debugging/`)

**Goal:** Give Claude a dedicated skill that guides Excel-add-in debugging workflows distinct from generic web debugging.

### What to implement

1. Create `skills/excel-addin-debugging/SKILL.md` with frontmatter:
   ```yaml
   ---
   name: excel-addin-debugging
   description: Debug and automate Excel add-ins running in WebView2 with Office.js awareness — active range inspection, Office context detection, dialog API handling, and requirement-set probing.
   ---
   ```
2. Content sections (follow the structure of [skills/excel-webview2/SKILL.md](../../skills/excel-webview2/SKILL.md)):
   - **When to use** — Triggers: user mentions Office.js, `Excel.run`, taskpane, ribbon command, add-in dialog.
   - **Readiness check** — Always call `excel_context_info` first; bail with a clear message if `hasOfficeGlobal: false`.
   - **Common workflows** — inspecting a selection, reading a named range, capturing taskpane state after a ribbon click.
   - **Dialog API gotcha** — `displayDialogAsync` opens a new debuggable target; use `list_pages` after invoking.
   - **Requirement-set probing** — how to branch behavior on `isSetSupported('ExcelApi', '1.12')`.
3. Add `skills/excel-addin-debugging/references/office-js-cheatsheet.md` with the most-used Office.js snippets (sourced entirely from Phase 0's reference file — nothing invented).
4. Update [skills/excel-webview2/SKILL.md](../../skills/excel-webview2/SKILL.md) to cross-link: "For Office.js-aware debugging, see the `excel-addin-debugging` skill."

### Documentation references

- Skill layout: `skills/a11y-debugging/` (uses `references/` subfolder).
- Frontmatter: `skills/excel-webview2/SKILL.md` top 4 lines.
- Snippets: ONLY from `docs/plans/office-js-api-reference.md`.

### Verification checklist

- `SKILL.md` parses as valid Markdown with frontmatter.
- `grep -nE "Excel\\.run|Office\\.context" skills/excel-addin-debugging/` returns hits only for APIs documented in Phase 0.
- Manual: load the skill via Claude and confirm its description triggers on a prompt like "what's in the currently selected range?"

### Anti-pattern guards

- Do not include Office.js snippets not verified in Phase 0.
- Do not duplicate generic DOM-debugging advice already in `excel-webview2` skill — cross-link instead.

---

## Phase 6 — Align Eval Scenarios

**Goal:** Remove browser-only scenarios and add Excel-flavored ones.

### What to implement

1. **Delete** these files in `scripts/eval_scenarios/` (they don't apply to Excel taskpanes):
   - `emulation_userAgent_test.ts`
   - `emulation_viewport_test.ts`
   - `isolated_context_test.ts`
   - `navigation_test.ts`
   - `page_id_routing_test.ts`
   - `select_page_test.ts`
   - `fix_webpage_issues_test.ts`
2. **Keep** `console_test.ts`, `frontend_snapshot_test.ts`, `input_test.ts`, `input_parallel_test.ts`, `network_test.ts`, `performance_test.ts`, `snapshot_test.ts`, `page_focus_keyboard_test.ts`.
3. **Conditionally keep** `lighthouse_a11y_test.ts`, `lighthouse_best_practices_test.ts` — only if Phase 1 decided Lighthouse stays.
4. **Add** `scripts/eval_scenarios/excel_context_test.ts` and `excel_active_range_test.ts` — self-contained scenarios that assert the new tools from Phases 3 and 4 behave correctly against a fixture taskpane (or against a mock that exposes `Office`/`Excel` globals — see existing eval harness for the fixture pattern).

### Documentation references

- Eval scenario structure: any surviving file in `scripts/eval_scenarios/`.
- Gemini eval runner: `scripts/gemini/*.ts` (already updated to point at `skills/excel-webview2/SKILL.md`).

### Verification checklist

- `npm run test` passes.
- `scripts/eval_scenarios/` has no files referencing `navigate`, `newPage`, or device emulation except through retained shared helpers.
- Eval runner picks up the two new Excel scenarios.

### Anti-pattern guards

- Don't try to mock `Excel.run` inside the MCP server — the scenarios should hit a real (or real-enough) Office.js fixture on the debuggable target.
- Don't inline Office.js snippets in scenarios that aren't in Phase 0's reference.

---

## Phase 7 — Final Verification

**Goal:** Prove the whole change lands cleanly.

### Tasks

1. `npm run format` — clean.
2. `npm run build` — succeeds.
3. `npm run test` — all green.
4. `npm run gen` — regenerates `docs/tool-reference.md` and `docs/slim-tool-reference.md`.
5. Grep sweep for ghosts of pruned tools:
   - `grep -rn "install_extension\|uninstall_extension\|reload_extension" src/ tests/ docs/ skills/` — zero hits (if extensions pruned).
   - `grep -rn "new_page\|navigate_page" src/ tests/` — zero hits in source, only in generated docs if surviving.
6. Manual smoke test: with a running Excel add-in on port 9222,
   - `excel_context_info` returns Office host info.
   - `excel_active_range` returns the currently selected range.
   - `list_pages` still works; `take_snapshot` still works.
7. Run the Gemini eval (`npm run eval` or equivalent) with `--include-skill` pointing at `excel-addin-debugging`. All scenarios should pass.
8. Update [CHANGELOG.md](../../CHANGELOG.md) (if present) with a summary of the tool surface changes.

### Anti-pattern checks

- Confirm `docs/tool-reference.md` was regenerated by `npm run gen`, not hand-edited (check `git diff --stat` — it should be a large regeneration, not a small surgical change).
- Confirm every new `.ts` file has the license header.
- Confirm `skills/excel-addin-debugging/references/office-js-cheatsheet.md` only cites APIs listed in `docs/plans/office-js-api-reference.md`.

---

## Phase Sequencing & Branching Advice

- Phase 0 → 1 → 2 are independent and can be parallelized in worktrees if desired.
- Phase 3 depends on Phases 0 and 2.
- Phase 4 depends on Phase 3.
- Phase 5 depends on Phase 0 (for snippets) and ideally Phases 3–4 (so cross-links exist).
- Phase 6 depends on Phases 3–4 (so the new scenarios have tools to target).
- Phase 7 depends on everything.

## Known Gaps / Confidence Notes

- **Lighthouse viability** is unverified. Phase 1 makes the decision empirically rather than assuming.
- **Office.js version** in the user's WebView2 may vary. Requirement-set probing in Phase 3 makes tools graceful about this, but we can't guarantee every API is present.
- **Context7 coverage** of Office.js was not verified during plan authoring. If Phase 0 finds context7 lacks coverage, fall back to `learn.microsoft.com/en-us/office/dev/add-ins/reference/javascript-api-for-office` and cite URLs directly.
