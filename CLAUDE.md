# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build and Development

- **Build**: `npm run build` (TypeScript compilation to `build/`)
- **Format & lint**: `npm run format` (ESLint + Prettier with caching)
- **Check format**: `npm run check-format` (runs checks without modifications)
- **Test**: `npm run test` (custom test runner at `scripts/test.mjs`, rebuilds on each run)
- **Full generation**: `npm run gen` (builds, generates docs, CLI, metrics, and formats)

## Type Safety

TypeScript strict mode is enforced:

- `noImplicitReturns`: every code path must return a value
- `noImplicitOverride`: subclass methods must explicitly use `override` keyword
- `noFallthroughCasesInSwitch`: switch statements must have breaks or return
- `forceConsistentCasingInFileNames`: file names must match imports exactly

## Project Context

This is a fork of [`ChromeDevTools/chrome-devtools-mcp`](https://github.com/ChromeDevTools/chrome-devtools-mcp) adapted for Excel add-ins running in WebView2. The server bridges Claude Code to a locally running Excel add-in via the Chrome DevTools Protocol on port 9222. It does not launch or manage Excel — the add-in must already be running with WebView2 remote debugging enabled.

## Node Version

Requires Node 20.19+, 22.12+, or 23+.

## Local Setup

The project has a format-on-write hook configured in `.claude/settings.local.json` that runs Prettier after edits. The `.claude/skills/verify` skill runs the full build, check, and test workflow.
