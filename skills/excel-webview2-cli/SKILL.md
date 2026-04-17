---
name: excel-webview2-cli
description: Use this skill to write shell scripts or run shell commands to automate tasks in the browser or otherwise use Excel WebView2 via CLI.
---

The `excel-webview2-mcp` CLI lets you interact with the browser from your terminal.

## Setup

_Note: If this is your very first time using the CLI, see [references/installation.md](references/installation.md) for setup. Installation is a one-time prerequisite and is **not** part of the regular AI workflow._

## AI Workflow

1. **Execute**: Run tools directly (e.g., `excel-webview2 list_pages`). The background server starts implicitly; **do not** run `start`/`status`/`stop` before each use.
2. **Inspect**: Use `take_snapshot` to get an element `<uid>`.
3. **Act**: Use `click`, `fill`, etc. State persists across commands.

Snapshot example:

```
uid=1_0 RootWebArea "Example Domain" url="https://example.com/"
  uid=1_1 heading "Example Domain" level="1"
```

## Command Usage

```sh
excel-webview2 <tool> [arguments] [flags]
```

Use `--help` on any command. Output defaults to Markdown, use `--output-format=json` for JSON.

## Input Automation (<uid> from snapshot)

```bash
excel-webview2 take_snapshot --help # Help message for commands, works for any command.
excel-webview2 take_snapshot # Take a text snapshot of the page to get UIDs for elements
excel-webview2 click "id" # Clicks on the provided element
excel-webview2 click "id" --dblClick true --includeSnapshot true # Double clicks and returns a snapshot
excel-webview2 drag "src" "dst" # Drag an element onto another element
excel-webview2 drag "src" "dst" --includeSnapshot true # Drag an element and return a snapshot
excel-webview2 fill "id" "text" # Type text into an input or select an option
excel-webview2 fill "id" "text" --includeSnapshot true # Fill an element and return a snapshot
excel-webview2 handle_dialog accept # Handle a browser dialog
excel-webview2 handle_dialog dismiss --promptText "hi" # Dismiss a dialog with prompt text
excel-webview2 hover "id" # Hover over the provided element
excel-webview2 hover "id" --includeSnapshot true # Hover over an element and return a snapshot
excel-webview2 press_key "Enter" # Press a key or key combination
excel-webview2 press_key "Control+A" --includeSnapshot true # Press a key and return a snapshot
excel-webview2 type_text "hello" # Type text using keyboard into a focused input
excel-webview2 type_text "hello" --submitKey "Enter" # Type text and press a submit key
excel-webview2 upload_file "id" "file.txt" # Upload a file through a provided element
excel-webview2 upload_file "id" "file.txt" --includeSnapshot true # Upload a file and return a snapshot
```

## Navigation

```bash
excel-webview2 close_page 1 # Closes the page by its index
excel-webview2 list_pages # Get a list of pages open in the browser
excel-webview2 navigate_page --url "https://example.com" # Navigates the currently selected page to a URL
excel-webview2 navigate_page --type "reload" --ignoreCache true # Reload page ignoring cache
excel-webview2 navigate_page --url "https://example.com" --timeout 5000 # Navigate with a timeout
excel-webview2 navigate_page --handleBeforeUnload "accept" # Handle before unload dialog
excel-webview2 navigate_page --type "back" --initScript "foo()" # Navigate back and run an init script
excel-webview2 new_page "https://example.com" # Creates a new page
excel-webview2 new_page "https://example.com" --background true --timeout 5000 # Create new page in background
excel-webview2 new_page "https://example.com" --isolatedContext "ctx" # Create new page with isolated context
excel-webview2 select_page 1 # Select a page as a context for future tool calls
excel-webview2 select_page 1 --bringToFront true # Select a page and bring it to front
```

## Emulation

```bash
excel-webview2 emulate --networkConditions "Offline" # Emulate network conditions
excel-webview2 emulate --cpuThrottlingRate 4 --geolocation "0x0" # Emulate CPU throttling and geolocation
excel-webview2 emulate --colorScheme "dark" --viewport "1920x1080" # Emulate color scheme and viewport
excel-webview2 emulate --userAgent "Mozilla/5.0..." # Emulate user agent
excel-webview2 resize_page 1920 1080 # Resizes the selected page's window
```

## Performance

```bash
excel-webview2 performance_analyze_insight "1" "LCPBreakdown" # Get more details on a specific Performance Insight
excel-webview2 performance_start_trace true false # Starts a performance trace recording
excel-webview2 performance_start_trace true true --filePath t.gz # Start trace and save to a file
excel-webview2 performance_stop_trace # Stops the active performance trace
excel-webview2 performance_stop_trace --filePath "t.json" # Stop trace and save to a file
excel-webview2 take_memory_snapshot "./snap.heapsnapshot" # Capture a memory heapsnapshot
```

## Network

```bash
excel-webview2 get_network_request # Get the currently selected network request
excel-webview2 get_network_request --reqid 1 --requestFilePath req.md # Get request by id and save to file
excel-webview2 get_network_request --responseFilePath res.md # Save response body to file
excel-webview2 list_network_requests # List all network requests
excel-webview2 list_network_requests --pageSize 50 --pageIdx 0 # List network requests with pagination
excel-webview2 list_network_requests --resourceTypes Fetch # Filter requests by resource type
excel-webview2 list_network_requests --includePreservedRequests true # Include preserved requests
```

## Debugging & Inspection

```bash
excel-webview2 evaluate_script "() => document.title" # Evaluate a JavaScript function on the page
excel-webview2 evaluate_script "(a) => a.innerText" --args 1_4 # Evaluate JS with UID arguments
excel-webview2 get_console_message 1 # Gets a console message by its ID
excel-webview2 lighthouse_audit --mode "navigation" # Run Lighthouse audit for navigation
excel-webview2 lighthouse_audit --mode "snapshot" --device "mobile" # Run Lighthouse audit for a snapshot on mobile
excel-webview2 lighthouse_audit --outputDirPath ./out # Run Lighthouse audit and save reports
excel-webview2 list_console_messages # List all console messages
excel-webview2 list_console_messages --pageSize 20 --pageIdx 1 # List console messages with pagination
excel-webview2 list_console_messages --types error --types info # Filter console messages by type
excel-webview2 list_console_messages --includePreservedMessages true # Include preserved messages
excel-webview2 take_screenshot # Take a screenshot of the page viewport
excel-webview2 take_screenshot --fullPage true --format "jpeg" --quality 80 # Take a full page screenshot as JPEG with quality
excel-webview2 take_screenshot --uid "id" --filePath "s.png" # Take a screenshot of an element
excel-webview2 take_snapshot # Take a text snapshot of the page from the a11y tree
excel-webview2 take_snapshot --verbose true --filePath "s.txt" # Take a verbose snapshot and save to file
```

## Service Management

```bash
excel-webview2 start   # Start or restart excel-webview2-mcp
excel-webview2 status  # Checks if excel-webview2-mcp is running
excel-webview2 stop    # Stop excel-webview2-mcp if any
```
