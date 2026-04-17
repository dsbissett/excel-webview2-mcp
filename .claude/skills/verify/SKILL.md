---
name: verify
description: Run local presubmit checks (format + tests) to mirror CI before committing or opening a PR.
---

Run the following commands in sequence and report any failures:

```bash
npm run check-format && npm run test:only
```

- `check-format` runs ESLint + Prettier in check mode (no writes) — mirrors the CI presubmit.
- `test:only` runs tests without rebuilding — use when you've already built.

If `check-format` fails, run `npm run format` to auto-fix, then re-run `check-format` to confirm.
If tests fail, report the failing test names and output. Do not attempt to fix failures automatically unless asked.
