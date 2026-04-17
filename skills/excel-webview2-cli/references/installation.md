# Installation

Install the package globally to make the `excel-webview2` command available. You only need to do this the first time you use it.

```sh
npm i @dsbissett/excel-webview2-mcp@latest -g
excel-webview2 status # check if install worked.
```

## Troubleshooting

- **Command not found:** If `excel-webview2` is not recognized, ensure your global npm `bin` directory is in your system's `PATH`. Restart your terminal or source your shell configuration file (e.g., `.bashrc`, `.zshrc`).
- **Permission errors:** If you encounter `EACCES` or permission errors during installation, avoid using `sudo`. Instead, use a node version manager like `nvm`, or configure npm to use a different global directory.
- **Old version running:** Run `excel-webview2 stop && npm uninstall -g @dsbissett/excel-webview2-mcp` before reinstalling, or ensure the latest version is being picked up by your path.
