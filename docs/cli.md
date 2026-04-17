# Excel WebView2 CLI

The `@dsbissett/excel-webview2-mcp` package includes an **experimental** CLI interface that allows you to interact with the browser directly from your terminal. This is particularly useful for debugging or when you want an agent to generate scripts that automate browser actions.

## Getting started

Install the package globally to make the `excel-webview2` command available:

```sh
npm i @dsbissett/excel-webview2-mcp@latest -g
excel-webview2 status # check if install worked.
```

## How it works

The CLI acts as a client to a background `excel-webview2-mcp` daemon (uses Unix sockets on Linux/Mac and named pipes on Windows).

- **Automatic Start**: The first time you call a tool (e.g., `list_pages`), the CLI automatically starts the MCP server and the browser in the background if they aren't already running.
- **Persistence**: The same background instance is reused for subsequent commands, preserving the browser state (open pages, cookies, etc.).
- **Manual Control**: You can explicitly manage the background process using `start`, `stop`, and `status`. The `start` command forwards all subsequent arguments to the underlying MCP server (e.g., `--headless`, `--userDataDir`) but not all args are supported. Run `excel-webview2 start --help` for supported args. Headless is enabled by default. Isolated is enabled by default unless `--userDataDir` is provided.

```sh
# Check if the daemon is running
excel-webview2 status

# Navigate the current page to a URL
excel-webview2 navigate_page "https://google.com"

# Take a screenshot and save it to a file
excel-webview2 take_screenshot --filePath screenshot.png

# Stop the background daemon when finished
excel-webview2 stop
```

## Command Usage

The CLI supports all tools available in the [Tool reference](./tool-reference.md).

```sh
excel-webview2 <tool> [arguments] [flags]
```

- **Required Arguments**: Passed as positional arguments.
- **Optional Arguments**: Passed as flags (e.g., `--filePath`, `--fullPage`).

### Examples

**New Page and Navigation:**

```sh
excel-webview2 new_page "https://example.com"
excel-webview2 navigate_page "https://web.dev" --type url
```

**Interaction:**

```sh
# Click an element by its UID from a snapshot
excel-webview2 click "element-uid-123"

# Fill a form field
excel-webview2 fill "input-uid-456" "search query"
```

**Analysis:**

```sh
# Run a Lighthouse audit (defaults to navigation mode)
excel-webview2 lighthouse_audit --mode snapshot
```

## Output format

By default, the CLI outputs a human-readable summary of the tool's result. For programmatic use, you can request raw JSON:

```sh
excel-webview2 list_pages --output-format=json
```

## Troubleshooting

If the CLI hangs or fails to connect, try stopping the background process:

```sh
excel-webview2 stop
```

For more verbose logs, set the `DEBUG` environment variable:

```sh
DEBUG=* excel-webview2 list_pages
```

## CLI generation

Implemented in `scripts/generate-cli.ts`. Some commands are excluded from CLI
generation such as `wait_for` and `fill_form`.

`excel-webview2-mcp` args are also filtered in `src/bin/excel-webview2.ts`
because not all args make sense in a CLI interface.
