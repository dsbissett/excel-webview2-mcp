# Office.js Cheatsheet

Verified snippets for Excel add-in debugging via `evaluate_script`. Every member here is sourced from [docs/plans/office-js-api-reference.md](../../../docs/plans/office-js-api-reference.md). Do not add entries that are not in that reference.

## 1. Global detection

```js
const hasOffice =
  typeof Office !== 'undefined' && !!Office.context && !!Office.context.host;
const hasExcel = typeof Excel !== 'undefined';
```

## 2. Office context snapshot

```js
const info = {
  host: Office.context.diagnostics.host,
  platform: Office.context.diagnostics.platform,
  version: Office.context.diagnostics.version,
  contentLanguage: Office.context.contentLanguage,
  displayLanguage: Office.context.displayLanguage,
};
```

`Office.HostType` members to compare against: `Excel`, `Word`, `PowerPoint`, `Outlook`, `OneNote`, `Project`.
`Office.PlatformType` members: `PC`, `Mac`, `OfficeOnline`, `iOS`, `Android`.

## 3. Requirement set probe

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
const supported = sets
  .filter(([name, ver]) =>
    Office.context.requirements.isSetSupported(name, ver),
  )
  .map(([name, ver]) => `${name} ${ver}`);
```

`isSetSupported` never throws — returns `false` for unknown names.

## 4. Active worksheet name

```js
await Excel.run(async ctx => {
  const sheet = ctx.workbook.worksheets.getActiveWorksheet();
  sheet.load('name');
  await ctx.sync();
  return sheet.name;
});
```

## 5. Read the selected range

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

Loadable `Excel.Range` properties: `address`, `values`, `formulas`, `formulasLocal`, `numberFormat`, `rowCount`, `columnCount`, `rowIndex`, `text`, `valueTypes`, `worksheet`.

## 6. Read a named range

```js
await Excel.run(async ctx => {
  const named = ctx.workbook.names.getItem('MyNamedRange');
  const range = named.getRange();
  range.load(['address', 'values']);
  await ctx.sync();
  return {address: range.address, values: range.values};
});
```

`names.getItem` throws if the name is missing. With `ExcelApi 1.4`+, use `getItemOrNullObject` and check `named.isNullObject` after sync.

## 7. Dialog API

```js
Office.context.ui.displayDialogAsync(
  'https://localhost:3000/dialog.html',
  {height: 40, width: 30},
  result => {
    const dialog = result.value;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, arg => {
      console.log('from dialog:', arg.message);
    });
  },
);
```

Inside the dialog page:

```js
Office.context.ui.messageParent(JSON.stringify({ok: true}), {
  targetOrigin: 'https://localhost:3000',
});
```

`startAddress` must be HTTPS and same-origin with the taskpane. On Excel desktop WebView2 the dialog is a separate CDP target — re-run `list_pages` after `displayDialogAsync`.

## 8. Readiness

```js
Office.onReady(info => {
  // info.host, info.platform
});
```

Never rely on `Office.initialize` alone in new tooling — prefer `Office.onReady`. Time-box any `await Office.onReady()` because if the library fails to load, the promise never resolves.

## Anti-patterns

- No `range.load('*')` — enumerate names.
- No property reads before `await ctx.sync()`.
- No assumption that `Office` / `Excel` globals exist — guard with `typeof`.
- No reliance on Excel-Online-only behavior when attached to desktop WebView2 (or vice versa) — branch on `Office.context.platform` or `isSetSupported('ExcelApiOnline', '1.1')`.
