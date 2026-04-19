<!-- AUTO GENERATED DO NOT EDIT - run 'npm run gen' to update-->

# Excel WebView2 MCP Tool Reference (~11784 cl100k_base tokens)

- **[Input automation](#input-automation)** (9 tools)
  - [`click`](#click)
  - [`drag`](#drag)
  - [`fill`](#fill)
  - [`fill_form`](#fill_form)
  - [`handle_dialog`](#handle_dialog)
  - [`hover`](#hover)
  - [`press_key`](#press_key)
  - [`type_text`](#type_text)
  - [`upload_file`](#upload_file)
- **[Navigation automation](#navigation-automation)** (4 tools)
  - [`close_page`](#close_page)
  - [`list_pages`](#list_pages)
  - [`select_page`](#select_page)
  - [`wait_for`](#wait_for)
- **[Excel](#excel)** (29 tools)
  - [`excel_active_range`](#excel_active_range)
  - [`excel_calculation_state`](#excel_calculation_state)
  - [`excel_chart_image`](#excel_chart_image)
  - [`excel_chart_info`](#excel_chart_info)
  - [`excel_context_info`](#excel_context_info)
  - [`excel_custom_xml_parts`](#excel_custom_xml_parts)
  - [`excel_find_in_range`](#excel_find_in_range)
  - [`excel_list_charts`](#excel_list_charts)
  - [`excel_list_comments`](#excel_list_comments)
  - [`excel_list_conditional_formats`](#excel_list_conditional_formats)
  - [`excel_list_data_validations`](#excel_list_data_validations)
  - [`excel_list_named_items`](#excel_list_named_items)
  - [`excel_list_pivot_tables`](#excel_list_pivot_tables)
  - [`excel_list_shapes`](#excel_list_shapes)
  - [`excel_list_tables`](#excel_list_tables)
  - [`excel_list_worksheets`](#excel_list_worksheets)
  - [`excel_pivot_table_info`](#excel_pivot_table_info)
  - [`excel_pivot_table_values`](#excel_pivot_table_values)
  - [`excel_range_formulas`](#excel_range_formulas)
  - [`excel_range_properties`](#excel_range_properties)
  - [`excel_range_special_cells`](#excel_range_special_cells)
  - [`excel_read_range`](#excel_read_range)
  - [`excel_settings_get`](#excel_settings_get)
  - [`excel_table_filters`](#excel_table_filters)
  - [`excel_table_info`](#excel_table_info)
  - [`excel_table_rows`](#excel_table_rows)
  - [`excel_used_range`](#excel_used_range)
  - [`excel_workbook_info`](#excel_workbook_info)
  - [`excel_worksheet_info`](#excel_worksheet_info)
- **[Emulation](#emulation)** (1 tools)
  - [`emulate`](#emulate)
- **[Performance](#performance)** (4 tools)
  - [`performance_analyze_insight`](#performance_analyze_insight)
  - [`performance_start_trace`](#performance_start_trace)
  - [`performance_stop_trace`](#performance_stop_trace)
  - [`take_memory_snapshot`](#take_memory_snapshot)
- **[Network](#network)** (2 tools)
  - [`get_network_request`](#get_network_request)
  - [`list_network_requests`](#list_network_requests)
- **[Debugging](#debugging)** (7 tools)
  - [`connection_status`](#connection_status)
  - [`evaluate_script`](#evaluate_script)
  - [`get_console_message`](#get_console_message)
  - [`lighthouse_audit`](#lighthouse_audit)
  - [`list_console_messages`](#list_console_messages)
  - [`take_screenshot`](#take_screenshot)
  - [`take_snapshot`](#take_snapshot)
- **[Add-in lifecycle](#add-in-lifecycle)** (3 tools)
  - [`excel_detect_addin`](#excel_detect_addin)
  - [`excel_launch_addin`](#excel_launch_addin)
  - [`excel_stop_addin`](#excel_stop_addin)

## Input automation

### `click`

**Description:** Clicks on the provided element

**Parameters:**

- **uid** (string) **(required)**: The uid of an element on the page from the page content snapshot
- **dblClick** (boolean) _(optional)_: Set to true for double clicks. Default is false.
- **includeSnapshot** (boolean) _(optional)_: Whether to include a snapshot in the response. Default is false.

---

### `drag`

**Description:** [`Drag`](#drag) an element onto another element

**Parameters:**

- **from_uid** (string) **(required)**: The uid of the element to [`drag`](#drag)
- **to_uid** (string) **(required)**: The uid of the element to drop into
- **includeSnapshot** (boolean) _(optional)_: Whether to include a snapshot in the response. Default is false.

---

### `fill`

**Description:** Type text into an input, text area or select an option from a &lt;select&gt; element.

**Parameters:**

- **uid** (string) **(required)**: The uid of an element on the page from the page content snapshot
- **value** (string) **(required)**: The value to [`fill`](#fill) in
- **includeSnapshot** (boolean) _(optional)_: Whether to include a snapshot in the response. Default is false.

---

### `fill_form`

**Description:** [`Fill`](#fill) out multiple form elements at once

**Parameters:**

- **elements** (array) **(required)**: Elements from snapshot to [`fill`](#fill) out.
- **includeSnapshot** (boolean) _(optional)_: Whether to include a snapshot in the response. Default is false.

---

### `handle_dialog`

**Description:** If a browser dialog was opened, use this command to handle it

**Parameters:**

- **action** (enum: "accept", "dismiss") **(required)**: Whether to dismiss or accept the dialog
- **promptText** (string) _(optional)_: Optional prompt text to enter into the dialog.

---

### `hover`

**Description:** [`Hover`](#hover) over the provided element

**Parameters:**

- **uid** (string) **(required)**: The uid of an element on the page from the page content snapshot
- **includeSnapshot** (boolean) _(optional)_: Whether to include a snapshot in the response. Default is false.

---

### `press_key`

**Description:** Press a key or key combination. Use this when other input methods like [`fill`](#fill)() cannot be used (e.g., keyboard shortcuts, navigation keys, or special key combinations).

**Parameters:**

- **key** (string) **(required)**: A key or a combination (e.g., "Enter", "Control+A", "Control++", "Control+Shift+R"). Modifiers: Control, Shift, Alt, Meta
- **includeSnapshot** (boolean) _(optional)_: Whether to include a snapshot in the response. Default is false.

---

### `type_text`

**Description:** Type text using keyboard into a previously focused input

**Parameters:**

- **text** (string) **(required)**: The text to type
- **submitKey** (string) _(optional)_: Optional key to press after typing. E.g., "Enter", "Tab", "Escape"

---

### `upload_file`

**Description:** Upload a file through a provided element.

**Parameters:**

- **filePath** (string) **(required)**: The local path of the file to upload
- **uid** (string) **(required)**: The uid of the file input element or an element that will open file chooser on the page from the page content snapshot
- **includeSnapshot** (boolean) _(optional)_: Whether to include a snapshot in the response. Default is false.

---

## Navigation automation

### `close_page`

**Description:** Closes the page by its index. The last open page cannot be closed.

**Parameters:**

- **pageId** (number) **(required)**: The ID of the page to close. Call [`list_pages`](#list_pages) to list pages.

---

### `list_pages`

**Description:** Get a list of pages open in the browser.

**Parameters:** None

---

### `select_page`

**Description:** Select a page as a context for future tool calls.

**Parameters:**

- **pageId** (number) **(required)**: The ID of the page to select. Call [`list_pages`](#list_pages) to get available pages.
- **bringToFront** (boolean) _(optional)_: Whether to focus the page and bring it to the top.

---

### `wait_for`

**Description:** Wait for the specified text to appear on the selected page.

**Parameters:**

- **text** (array) **(required)**: Non-empty list of texts. Resolves when any value appears on the page.
- **timeout** (integer) _(optional)_: Maximum wait time in milliseconds. If set to 0, the default timeout will be used.

---

## Excel

### `excel_active_range`

**Description:** Returns the currently selected Excel range (address, dimensions, and values). Optionally includes formulas and number formats. Requires an Excel add-in target with Excel.run available.

**Parameters:**

- **includeFormulas** (boolean) _(optional)_: If true, also return the A1-style formulas for each cell.
- **includeNumberFormat** (boolean) _(optional)_: If true, also return the Excel number-format code per cell.

---

### `excel_calculation_state`

**Description:** Returns the workbook calculation mode (automatic/manual/etc.) and current calculation state (done/calculating/pending).

**Parameters:** None

---

### `excel_chart_image`

**Description:** Returns a chart rendered as a PNG image, encoded as base64. Useful for visual verification.

**Parameters:**

- **name** (string) **(required)**: Chart name.
- **sheet** (string) **(required)**: Worksheet name containing the chart.
- **height** (number) _(optional)_: Image height in pixels. Defaults to the chart’s natural size.
- **width** (number) _(optional)_: Image width in pixels. Defaults to the chart’s natural size.

---

### `excel_chart_info`

**Description:** Returns detailed information about a chart: type, title, series names, axis titles, and source data address.

**Parameters:**

- **name** (string) **(required)**: Chart name.
- **sheet** (string) **(required)**: Worksheet name containing the chart.

---

### `excel_context_info`

**Description:** Returns Office.js and Excel host information for the selected page, including supported requirement sets when available.

**Parameters:** None

---

### `excel_custom_xml_parts`

**Description:** Lists custom XML parts stored in the workbook: id and namespace URI.

**Parameters:** None

---

### `excel_find_in_range`

**Description:** Finds all matches of a text string within a range. Returns the combined match address and cell count.

**Parameters:**

- **text** (string) **(required)**: Text to search for.
- **address** (string) _(optional)_: A1 reference such as 'Sheet1!A1:C10' or 'A1:C10'. If omitted, the active selection is used.
- **completeMatch** (boolean) _(optional)_: If true, require a whole-cell match. Defaults to false.
- **matchCase** (boolean) _(optional)_: If true, the search is case-sensitive. Defaults to false.
- **sheet** (string) _(optional)_: Worksheet name. Used when address omits the sheet prefix; ignored if address includes one.

---

### `excel_list_charts`

**Description:** Lists all charts across worksheets: name, id, worksheet, type, title, position, and size.

**Parameters:**

- **sheet** (string) _(optional)_: Worksheet name. Omit to list charts on all worksheets.

---

### `excel_list_comments`

**Description:** Lists comments and replies on a worksheet: author, content, timestamp, and cell address.

**Parameters:**

- **sheet** (string) _(optional)_: Worksheet name. Omit to use the active worksheet.

---

### `excel_list_conditional_formats`

**Description:** Lists conditional-format rules on a range: id, type, priority, stopIfTrue. Omit address to use the active worksheet’s used range.

**Parameters:**

- **address** (string) _(optional)_: A1 reference such as 'Sheet1!A1:C10' or 'A1:C10'. If omitted, the active selection is used.
- **sheet** (string) _(optional)_: Worksheet name. Used when address omits the sheet prefix; ignored if address includes one.

---

### `excel_list_data_validations`

**Description:** Returns data-validation configuration on a range: type, rule, error alert, and prompt. Omit address to use the active selection.

**Parameters:**

- **address** (string) _(optional)_: A1 reference such as 'Sheet1!A1:C10' or 'A1:C10'. If omitted, the active selection is used.
- **sheet** (string) _(optional)_: Worksheet name. Used when address omits the sheet prefix; ignored if address includes one.

---

### `excel_list_named_items`

**Description:** Lists workbook-scoped named items (named ranges and formulas) with name, type, value, formula, visibility, and comment.

**Parameters:** None

---

### `excel_list_pivot_tables`

**Description:** Lists all PivotTables in the workbook with name, worksheet, layout address, and enabled flags.

**Parameters:** None

---

### `excel_list_shapes`

**Description:** Lists shapes (including images) on a worksheet: name, id, type, position, size, visibility.

**Parameters:**

- **sheet** (string) _(optional)_: Worksheet name. Omit to use the active worksheet.

---

### `excel_list_tables`

**Description:** Lists all tables (ListObjects) in the workbook with name, worksheet, address, header/total row flags, row count, and style.

**Parameters:** None

---

### `excel_list_worksheets`

**Description:** Lists all worksheets in the workbook with name, id, position, visibility, and tab color.

**Parameters:** None

---

### `excel_pivot_table_info`

**Description:** Returns the structure of a PivotTable: row, column, data, and filter hierarchies with their source field names.

**Parameters:**

- **name** (string) **(required)**: PivotTable name.

---

### `excel_pivot_table_values`

**Description:** Returns the rendered values of a PivotTable layout range with truncation when it exceeds the cell cap.

**Parameters:**

- **name** (string) **(required)**: PivotTable name.

---

### `excel_range_formulas`

**Description:** Returns formulas (A1 and R1C1) alongside resolved values for a range. Useful for verifying formula edits.

**Parameters:**

- **address** (string) _(optional)_: A1 reference such as 'Sheet1!A1:C10' or 'A1:C10'. If omitted, the active selection is used.
- **sheet** (string) _(optional)_: Worksheet name. Used when address omits the sheet prefix; ignored if address includes one.

---

### `excel_range_properties`

**Description:** Returns rich properties for a range: value types, hasSpill, row/column hidden flags, and selected format details (font, [`fill`](#fill), alignment, borders). Use include flags to bound payload.

**Parameters:**

- **address** (string) _(optional)_: A1 reference such as 'Sheet1!A1:C10' or 'A1:C10'. If omitted, the active selection is used.
- **includeFormat** (boolean) _(optional)_: If true, include font, [`fill`](#fill), alignment, and border summary per cell.
- **includeStyle** (boolean) _(optional)_: If true, include the named style of each cell.
- **sheet** (string) _(optional)_: Worksheet name. Used when address omits the sheet prefix; ignored if address includes one.

---

### `excel_range_special_cells`

**Description:** Finds cells within a range matching a category: 'constants', 'formulas', 'blanks', or 'visible'. Optionally filter by value type. Returns the resulting address and cell count.

**Parameters:**

- **cellType** (enum: "constants", "formulas", "blanks", "visible") **(required)**: Category of special cells to locate.
- **address** (string) _(optional)_: A1 reference such as 'Sheet1!A1:C10' or 'A1:C10'. If omitted, the active selection is used.
- **sheet** (string) _(optional)_: Worksheet name. Used when address omits the sheet prefix; ignored if address includes one.
- **valueType** (enum: "all", "errors", "logical", "numbers", "text") _(optional)_: For 'constants' or 'formulas', filter by value type. Defaults to 'all'.

---

### `excel_read_range`

**Description:** Reads a range by address (e.g. 'Sheet1!A1:C10' or 'A1:C10' with a sheet param). Omit address to read the active selection. Returns values and optionally formulas / number formats.

**Parameters:**

- **address** (string) _(optional)_: A1 reference such as 'Sheet1!A1:C10' or 'A1:C10'. If omitted, the active selection is used.
- **includeFormulas** (boolean) _(optional)_: If true, also return the A1-style formulas for each cell.
- **includeNumberFormat** (boolean) _(optional)_: If true, also return the Excel number-format code per cell.
- **sheet** (string) _(optional)_: Worksheet name. Used when address omits the sheet prefix; ignored if address includes one.

---

### `excel_settings_get`

**Description:** Reads add-in document settings from Office.context.document.settings. Returns all keys or a single key’s value.

**Parameters:**

- **key** (string) _(optional)_: If provided, return only this setting’s value. Otherwise, return all settings.

---

### `excel_table_filters`

**Description:** Returns the active filter criteria per column for a table. Columns without an active filter have null criteria.

**Parameters:**

- **name** (string) **(required)**: Table name (ListObject name).

---

### `excel_table_info`

**Description:** Returns detail for a single table: name, worksheet, address, row count, columns (name + filter criteria), header/total row flags, and style.

**Parameters:**

- **name** (string) **(required)**: Table name (ListObject name).

---

### `excel_table_rows`

**Description:** Returns the data-body values of a table with truncation when row\*column count exceeds the cell cap.

**Parameters:**

- **name** (string) **(required)**: Table name (ListObject name).
- **includeHeaders** (boolean) _(optional)_: If true, include the header row names.

---

### `excel_used_range`

**Description:** Returns values (and optionally formulas / number formats) for a worksheet’s used range, with truncation when the range exceeds the cell cap.

**Parameters:**

- **includeFormulas** (boolean) _(optional)_: If true, also return the A1-style formulas for each cell.
- **includeNumberFormat** (boolean) _(optional)_: If true, also return the Excel number-format code per cell.
- **sheet** (string) _(optional)_: Worksheet name. Omit to use the active worksheet.
- **valuesOnly** (boolean) _(optional)_: If true (default), only cells with values count toward the used range.

---

### `excel_workbook_info`

**Description:** Returns workbook-level metadata: name, save state, calculation mode and state, and protection state.

**Parameters:** None

---

### `excel_worksheet_info`

**Description:** Returns metadata for a single worksheet: used range address, visibility, protection, gridlines, tab color, and dimensions.

**Parameters:**

- **sheet** (string) _(optional)_: Worksheet name. Omit to use the active worksheet.

---

## Emulation

### `emulate`

**Description:** Throttles network and/or CPU on the selected page.

**Parameters:**

- **cpuThrottlingRate** (number) _(optional)_: Represents the CPU slowdown factor. Omit or set the rate to 1 to disable throttling
- **networkConditions** (enum: "Offline", "Slow 3G", "Fast 3G", "Slow 4G", "Fast 4G") _(optional)_: Throttle network. Omit to disable throttling.

---

## Performance

### `performance_analyze_insight`

**Description:** Provides more detailed information on a specific Performance Insight of an insight set that was highlighted in the results of a trace recording.

**Parameters:**

- **insightName** (string) **(required)**: The name of the Insight you want more information on. For example: "DocumentLatency" or "LCPBreakdown"
- **insightSetId** (string) **(required)**: The id for the specific insight set. Only use the ids given in the "Available insight sets" list.

---

### `performance_start_trace`

**Description:** Start a performance trace on the selected webpage. Use to find frontend performance issues, Core Web Vitals (LCP, INP, CLS), and improve page load speed.

**Parameters:**

- **autoStop** (boolean) _(optional)_: Determines if the trace recording should be automatically stopped.
- **filePath** (string) _(optional)_: The absolute file path, or a file path relative to the current working directory, to save the raw trace data. For example, trace.json.gz (compressed) or trace.json (uncompressed).
- **reload** (boolean) _(optional)_: Determines if, once tracing has started, the current selected page should be automatically reloaded.

---

### `performance_stop_trace`

**Description:** Stop the active performance trace recording on the selected webpage.

**Parameters:**

- **filePath** (string) _(optional)_: The absolute file path, or a file path relative to the current working directory, to save the raw trace data. For example, trace.json.gz (compressed) or trace.json (uncompressed).

---

### `take_memory_snapshot`

**Description:** Capture a heap snapshot of the currently selected page. Use to analyze the memory distribution of JavaScript objects and debug memory leaks.

**Parameters:**

- **filePath** (string) **(required)**: A path to a .heapsnapshot file to save the heapsnapshot to.

---

## Network

### `get_network_request`

**Description:** Gets a network request by an optional reqid, if omitted returns the currently selected request in the DevTools Network panel.

**Parameters:**

- **reqid** (number) _(optional)_: The reqid of the network request. If omitted returns the currently selected request in the DevTools Network panel.
- **requestFilePath** (string) _(optional)_: The absolute or relative path to save the request body to. If omitted, the body is returned inline.
- **responseFilePath** (string) _(optional)_: The absolute or relative path to save the response body to. If omitted, the body is returned inline.

---

### `list_network_requests`

**Description:** List all requests for the currently selected page since the last navigation.

**Parameters:**

- **includePreservedRequests** (boolean) _(optional)_: Set to true to return the preserved requests over the last 3 navigations.
- **pageIdx** (integer) _(optional)_: Page number to return (0-based). When omitted, returns the first page.
- **pageSize** (integer) _(optional)_: Maximum number of requests to return. When omitted, returns all requests.
- **resourceTypes** (array) _(optional)_: Filter requests to only return requests of the specified resource types. When omitted or empty, returns all requests.

---

## Debugging

### `connection_status`

**Description:** Reports whether the server is currently attached to a browser and which CDP endpoint it is tracking.

**Parameters:**

- **probe** (boolean) _(optional)_: If true, re-runs the CDP /json/version probe for the tracked endpoint instead of returning cached probe state.

---

### `evaluate_script`

**Description:** Evaluate a JavaScript function inside the currently selected page. Returns the response as JSON,
so returned values have to be JSON-serializable.

**Parameters:**

- **function** (string) **(required)**: A JavaScript function declaration to be executed by the tool in the currently selected page.
  Example without arguments: `() => {
  return document.title
}` or `async () => {
  return await fetch("example.com")
}`.
  Example with arguments: `(el) => {
  return el.innerText;
}`

- **args** (array) _(optional)_: An optional list of arguments to pass to the function.

---

### `get_console_message`

**Description:** Gets a console message by its ID. You can get all messages by calling [`list_console_messages`](#list_console_messages).

**Parameters:**

- **msgid** (number) **(required)**: The msgid of a console message on the page from the listed console messages

---

### `lighthouse_audit`

**Description:** Get Lighthouse score and reports for accessibility, SEO and best practices. This excludes performance. For performance audits, run [`performance_start_trace`](#performance_start_trace)

**Parameters:**

- **device** (enum: "desktop", "mobile") _(optional)_: Device to [`emulate`](#emulate).
- **mode** (enum: "navigation", "snapshot") _(optional)_: "navigation" reloads &amp; audits. "snapshot" analyzes current state.
- **outputDirPath** (string) _(optional)_: Directory for reports. If omitted, uses temporary files.

---

### `list_console_messages`

**Description:** List all console messages for the currently selected page since the last navigation.

**Parameters:**

- **includePreservedMessages** (boolean) _(optional)_: Set to true to return the preserved messages over the last 3 navigations.
- **pageIdx** (integer) _(optional)_: Page number to return (0-based). When omitted, returns the first page.
- **pageSize** (integer) _(optional)_: Maximum number of messages to return. When omitted, returns all messages.
- **types** (array) _(optional)_: Filter messages to only return messages of the specified resource types. When omitted or empty, returns all messages.

---

### `take_screenshot`

**Description:** Take a screenshot of the page or element.

**Parameters:**

- **filePath** (string) _(optional)_: The absolute path, or a path relative to the current working directory, to save the screenshot to instead of attaching it to the response.
- **format** (enum: "png", "jpeg", "webp") _(optional)_: Type of format to save the screenshot as. Default is "png"
- **fullPage** (boolean) _(optional)_: If set to true takes a screenshot of the full page instead of the currently visible viewport. Incompatible with uid.
- **quality** (number) _(optional)_: Compression quality for JPEG and WebP formats (0-100). Higher values mean better quality but larger file sizes. Ignored for PNG format.
- **uid** (string) _(optional)_: The uid of an element on the page from the page content snapshot. If omitted, takes a page screenshot.

---

### `take_snapshot`

**Description:** Take a text snapshot of the currently selected page based on the a11y tree. The snapshot lists page elements along with a unique
identifier (uid). Always use the latest snapshot. Prefer taking a snapshot over taking a screenshot. The snapshot indicates the element selected
in the DevTools Elements panel (if any).

**Parameters:**

- **filePath** (string) _(optional)_: The absolute path, or a path relative to the current working directory, to save the snapshot to instead of attaching it to the response.
- **verbose** (boolean) _(optional)_: Whether to include all possible information available in the full a11y tree. Default is false.

---

## Add-in lifecycle

### `excel_detect_addin`

**Description:** Inspects a working directory and reports whether it looks like an Excel add-in project (manifest location, manifest kind, package manager, and any existing remote-debugging script).

**Parameters:**

- **cwd** (string) _(optional)_: Directory to inspect. Defaults to the MCP server working directory.

---

### `excel_launch_addin`

**Description:** Launches Excel with the detected add-in and WebView2 remote debugging enabled. Idempotent per manifest path: re-calling returns the tracked launch instead of spawning a duplicate.

**Parameters:**

- **autoConnect** (boolean) _(optional)_
- **cwd** (string) _(optional)_
- **devServerTimeoutMs** (integer) _(optional)_
- **extraBrowserArgs** (array) _(optional)_
- **manifestPath** (string) _(optional)_
- **port** (integer) _(optional)_
- **skipDevServer** (boolean) _(optional)_
- **timeoutMs** (integer) _(optional)_

---

### `excel_stop_addin`

**Description:** Stops the most recent Excel add-in launched by [`excel_launch_addin`](#excel_launch_addin) (or a specific manifest). Runs office-addin-debugging stop and kills the process if it does not exit cleanly.

**Parameters:**

- **manifestPath** (string) _(optional)_

---
