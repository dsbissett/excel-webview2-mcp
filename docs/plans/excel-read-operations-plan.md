# Excel Read Operations Plan

Expand the `EXCEL` tool category with read-only operations so Claude can verify that add-in code changes produced the expected workbook state. All tools run Office.js via `Excel.run` inside the WebView2 page and return structured JSON. No write operations are in scope.

## Design guardrails

- **Read-only**: every tool sets `annotations.readOnlyHint: true` and never calls mutating Office.js APIs.
- **Requirement-set gated**: each tool declares the minimum `ExcelApi` version it needs and fails cleanly when `Office.context.requirements.isSetSupported` is false.
- **Bounded payloads**: default cell/row caps (reuse the `MAX_CELLS = 1000` pattern from [excel.ts:6](src/tools/excel.ts#L6)); tools return `truncated: true` instead of dumping huge sheets.
- **Address-based targeting**: tools that take a range accept either `{sheet, address}` or an A1 reference like `'Sheet1!A1:C10'`; omitted address means the active selection.
- **Consistent error shape**: when `Excel` global is missing or `Excel.run` throws, return `{error: string}` and let the handler print `ERROR: ...` (same pattern as `excel_active_range`).

## Proposed tools

### Workbook-level

| Tool                      | Purpose                                                                               | Key Office.js calls                                                      | Min ExcelApi |
| ------------------------- | ------------------------------------------------------------------------------------- | ------------------------------------------------------------------------ | ------------ |
| `excel_workbook_info`     | Workbook name, path, save state, calculation mode, protection state.                  | `context.workbook.load('name')`, `application.calculationMode`           | 1.1 / 1.7    |
| `excel_list_worksheets`   | All sheets with name, id, position, visibility, tab color.                            | `workbook.worksheets.load('items/name,id,position,visibility,tabColor')` | 1.1          |
| `excel_list_named_items`  | Workbook-scoped named ranges/formulas with name, type, value, comment.                | `workbook.names.load(...)`                                               | 1.4          |
| `excel_list_tables`       | All tables (ListObjects): name, sheet, address, header row, row count, style.         | `workbook.tables.load(...)`                                              | 1.1          |
| `excel_list_pivot_tables` | All PivotTables: name, sheet, source, layout.                                         | `workbook.pivotTables.load(...)`                                         | 1.8          |
| `excel_list_charts`       | All charts across sheets: name, sheet, type, title, position.                         | `worksheet.charts.load(...)` iterated                                    | 1.1          |
| `excel_calculation_state` | Calculation mode + whether calc is pending/dirty.                                     | `application.calculationState`, `calculationMode`                        | 1.7          |
| `excel_custom_xml_parts`  | List custom XML parts (id, namespaceUri) — useful for add-ins that store state there. | `workbook.customXmlParts.load(...)`                                      | 1.5          |

### Worksheet-level

| Tool                             | Purpose                                                                                                | Key Office.js calls                                                         | Min ExcelApi |
| -------------------------------- | ------------------------------------------------------------------------------------------------------ | --------------------------------------------------------------------------- | ------------ |
| `excel_worksheet_info`           | Single sheet metadata: used range address, visibility, protection, freeze panes, gridlines, tab color. | `worksheet.getUsedRange()`, `worksheet.protection`, `worksheet.freezePanes` | 1.1 / 1.7    |
| `excel_used_range`               | Values / formulas / number format of used range with truncation caps.                                  | `worksheet.getUsedRange().load(...)`                                        | 1.1          |
| `excel_list_sheet_tables`        | Tables on a given sheet.                                                                               | `worksheet.tables.load(...)`                                                | 1.1          |
| `excel_list_sheet_named_items`   | Worksheet-scoped names.                                                                                | `worksheet.names.load(...)`                                                 | 1.4          |
| `excel_list_comments`            | All comments + replies on a sheet (author, text, cell).                                                | `worksheet.comments.load(...)`                                              | 1.10         |
| `excel_list_shapes`              | Shapes/images on a sheet (name, type, position).                                                       | `worksheet.shapes.load(...)`                                                | 1.9          |
| `excel_list_conditional_formats` | Conditional-format rules on a sheet or range (type, priority, range).                                  | `range.conditionalFormats.load(...)`                                        | 1.6          |
| `excel_list_data_validations`    | Data-validation rules on a range (type, formula, error alert).                                         | `range.dataValidation.load(...)`                                            | 1.8          |

### Range-level

| Tool                        | Purpose                                                                                                                                                                                                        | Notes                                                                                      | Min ExcelApi    |
| --------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------ | --------------- |
| `excel_read_range`          | Generalized version of `excel_active_range` that also accepts an explicit address, not just the selection. Returns values, rowCount/columnCount, and optional formulas / numberFormat / text.                  | Supersedes nothing — keep `excel_active_range` for the common case, delegate shared logic. | 1.1             |
| `excel_range_properties`    | Rich property bundle: `valueTypes`, `hasSpill`, `rowHidden`/`columnHidden`, `style`, `format.font`, `format.fill.color`, `format.borders`, `format.horizontalAlignment`. Let callers choose via include flags. | Many small loads — document performance caveat.                                            | 1.1 / 1.7 / 1.9 |
| `excel_range_formulas`      | Returns formulas (A1 and R1C1) and the resolved values side-by-side. Handy for verifying a formula edit.                                                                                                       | `formulas`, `formulasR1C1`, `values`.                                                      | 1.1             |
| `excel_range_special_cells` | List cells matching a category (constants, formulas, blanks, errors) inside a range.                                                                                                                           | `range.getSpecialCellsOrNullObject(...)`.                                                  | 1.9             |
| `excel_find_in_range`       | Find all matches of a string in a range (case/whole-cell options). Returns addresses + values.                                                                                                                 | `range.findAllOrNullObject(...)`.                                                          | 1.9             |

### Table / PivotTable / Chart detail

| Tool                       | Purpose                                                             | Key Office.js calls                                                     | Min ExcelApi |
| -------------------------- | ------------------------------------------------------------------- | ----------------------------------------------------------------------- | ------------ |
| `excel_table_info`         | One table: columns, header names, total row, range, filters active. | `table.columns.load`, `table.getRange()`                                | 1.1          |
| `excel_table_rows`         | Data-body rows of a table with truncation.                          | `table.getDataBodyRange().load('values')`                               | 1.1          |
| `excel_table_filters`      | Active filter criteria per column.                                  | `column.filter.criteria`                                                | 1.2          |
| `excel_pivot_table_info`   | Row/column/data/filter hierarchies and layout.                      | `pivotTable.rowHierarchies/columnHierarchies/dataHierarchies.load(...)` | 1.8          |
| `excel_pivot_table_values` | The rendered pivot output range values.                             | `pivotTable.layout.getRange().load('values')`                           | 1.8          |
| `excel_chart_info`         | Chart type, title, series names, axis titles, source data address.  | `chart.series.load(...)`, `chart.axes...`                               | 1.1          |
| `excel_chart_image`        | Chart as PNG (base64) for visual verification.                      | `chart.getImage()`                                                      | 1.2          |

### Office runtime & diagnostics (adjuncts)

| Tool                          | Purpose                                                                                                                                  |
| ----------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| `excel_settings_get`          | Read add-in document settings (`Office.context.document.settings`) — add-ins store config here.                                          |
| `excel_active_selection_info` | Lighter-weight sibling of `excel_active_range`: just address + sheet + selection type, no values.                                        |
| `excel_runtime_errors`        | Last Office.js error (if any) captured from a ring buffer installed on `Office.onReady`. _Optional; may fold into `excel_context_info`._ |

## Shared implementation helpers

Create `src/tools/excelHelpers.ts`:

- `runExcel<T>(page, fn, args)` — wraps `page.pptrPage.evaluate` + `Excel.run`, normalizes the `{error}` shape.
- `requireRequirementSet(page, set, version)` — guard that returns a clean error when the host is too old.
- `truncateGrid(values, max)` — shared truncation with `{truncated, rowCount, columnCount}` reporting.
- `resolveRangeTarget({sheet, address})` — returns a Range proxy from either an active selection, a sheet+address pair, or a workbook-scoped A1 reference.

## Phasing

1. **Phase 1 — Structural reads** (low risk, high value for verification):
   `excel_workbook_info`, `excel_list_worksheets`, `excel_worksheet_info`, `excel_used_range`, `excel_read_range`, `excel_list_tables`, `excel_list_named_items`.
2. **Phase 2 — Rich range inspection**:
   `excel_range_properties`, `excel_range_formulas`, `excel_range_special_cells`, `excel_find_in_range`, `excel_list_conditional_formats`, `excel_list_data_validations`.
3. **Phase 3 — Derived objects**:
   `excel_table_info`, `excel_table_rows`, `excel_table_filters`, `excel_list_comments`, `excel_list_shapes`, `excel_calculation_state`.
4. **Phase 4 — PivotTables, charts, custom XML, settings**:
   `excel_list_pivot_tables`, `excel_pivot_table_info`, `excel_pivot_table_values`, `excel_list_charts`, `excel_chart_info`, `excel_chart_image`, `excel_custom_xml_parts`, `excel_settings_get`.

Each phase ships with:

- Unit tests stubbing `page.pptrPage.evaluate` (mirroring `tests/tools/excel.test.ts`).
- Doc generation run (`npm run gen`) so `docs/tool-reference.md` picks up the new tools.
- Skill doc update (`skills/excel-webview2/SKILL.md`) listing the new verification workflow.

## Open questions

- Should range-returning tools accept a max-cell override per-call, or keep the global `MAX_CELLS = 1000`?
  - Keep the global `MAX_CELLS = 1000`
- Do we want a single `excel_inspect` mega-tool with include flags, or many focused tools? Current plan favors focused tools for clearer tool descriptions to the model.
  - We want many focused tools over a single monolith, each with clear tool descriptions.
- How should we surface binary output (`excel_chart_image` PNG) — inline base64 or an MCP resource?
  - Unknown at this point.
