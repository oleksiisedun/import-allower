# import-allower

A standalone Google Apps Script **library** project (not container-bound to any spreadsheet). It's added as a dependency by other Apps Script projects that need to silently grant `IMPORTRANGE` access instead of making the user click "Allow access" manually — see README.md for the consumer-facing API and usage examples.

## Invariant: never touches spreadsheet content

This library only ever reads formulas/values (to detect and resolve `IMPORTRANGE` sources) and makes one `UrlFetchApp` call per source to record an access grant. It never calls a mutating `SpreadsheetApp` method. Any change to this library must preserve that — if a future feature needs to write to the spreadsheet, that's a signal to reconsider the change, not just add the write.

## Why the endpoint is undocumented and what that means

`PermissionGranter.js` POSTs to `docs.google.com/.../externaldata/addimportrangepermissions` — the same internal endpoint the Sheets UI's "Allow access" button calls, discovered by inspection rather than published as a public API. There's no supported alternative for granting `IMPORTRANGE` access programmatically. Because it's undocumented, it can change without notice; `grantImportRangeAccess_` returns the raw HTTP status and response body for every source specifically so a consumer (or a future debugging session) can tell a real failure (bad source ID, no access) apart from the endpoint itself having changed shape.

## Why detection walks the whole formula, not just the first match

A formula can combine multiple `IMPORTRANGE` calls (inside `VLOOKUP`, joined with `&`, as separate array items). `ImportRangeScanner.js`'s `extractImportRangeFirstArgs_` finds every `IMPORTRANGE(` occurrence in the formula text via a global regex, then `parseFirstArg_` hand-parses each call's first argument (respecting quoted-string escaping and nested parens) rather than relying on a single regex match — a single match would silently miss every `IMPORTRANGE` after the first one in a combined formula. Don't collapse this back to a single-match regex; that regressed exactly the "multiple IMPORTRANGE per formula" case this was built to handle.

## Why cell discovery reads getFormulas() directly instead of using TextFinder

`forEachImportRangeFormula_` walks every sheet's `getDataRange().getFormulas()` grid rather than using `ss.createTextFinder('IMPORTRANGE').findAll()`. TextFinder was tried first and works for most formulas, but for an array-literal formula that spills across cells (e.g. `={IMPORTRANGE(...);IMPORTRANGE(...)}`), it can attribute a match to one of the spilled result cells instead of the anchor cell that actually holds the formula text — and `getFormula()` on a spilled cell returns `''`, silently dropping whichever `IMPORTRANGE` call landed there. Reading `getFormulas()` straight from each sheet's data range sidesteps this: only the anchor cell of any formula (spilled or not) ever has non-empty formula text, so nothing gets attributed to the wrong cell. Don't reintroduce TextFinder for this; it regresses the array-literal case.

## Why the trigger installer takes a function name, not a function reference

`installAutoApproveTrigger` in `Main.js` takes `handlerFunctionName: string`, not a function value. Apps Script's installable-trigger service resolves the handler by name against the *calling* project's global scope at trigger-fire time — a trigger can never point directly at a library-qualified function like `ImportAllower.autoApproveImportRanges`. The consuming project must define its own top-level wrapper function and pass that name in. Don't "simplify" this to accept a callback; it would work when tested inline (same project) and silently break for every real consumer (a different project, calling through the library).

## Why grant requests are sent in chunked batches, not all at once or one at a time

`grantImportRangeAccess_` in `PermissionGranter.js` sends requests via `UrlFetchApp.fetchAll` in batches of `GRANT_BATCH_SIZE_` (10) rather than one single `fetchAll` for every source, or one `UrlFetchApp.fetch` per source. A single `fetchAll` is fastest but gives no progress feedback until every source is done — `autoApproveImportRanges` can take a while with many `IMPORTRANGE` sources, and the caller wants to see results in the log as they happen. One-at-a-time `fetch` calls log immediately but lose the parallelism `fetchAll` provides within a batch, making the whole scan noticeably slower. Chunking keeps most of the speed of `fetchAll` while still surfacing a log line every `GRANT_BATCH_SIZE_` sources. Don't collapse this back to either extreme without re-confirming the speed/feedback tradeoff still favors it.

## File layout

Deployable code lives in `src/` (`.clasp.json` has `"rootDir": "src"`, so only `src/` is pushed; tooling and docs stay at the root). `.clasp.json` is tracked in git, so `rootDir` persists across checkouts — re-add it if `clasp clone`/`clasp create` ever regenerates the file.

- `src/Main.js` — public API: `autoApproveImportRanges`, `installAutoApproveTrigger`, `logImportRangeFormulas`.
- `src/ImportRangeScanner.js` — `IMPORTRANGE` detection: `findImportRangeSourceIds` (public) plus the formula-walking internals (`findImportRangeSourceIdsForSpreadsheet_`, `forEachImportRangeFormula_`, `extractImportRangeFirstArgs_`, `parseFirstArg_`, `resolveImportRangeSourceId_`).
- `src/PermissionGranter.js` — `grantImportRangeAccess_`, the only file that calls `UrlFetchApp`.
- `src/SpreadsheetResolver.js` — `resolveSpreadsheet_`, the shared "given ID or active spreadsheet" lookup every public entry point uses.
- `src/appsscript.json` — manifest.

## Testing

Run `npm run check` after every edit — it runs `npm run lint` (ESLint, configured for the Apps Script runtime in `eslint.config.mjs`), `npm run typecheck` (`tsc` with `checkJs` over `src/`, see `jsconfig.json`) and `npm run test`. All are fast and offline. Because all files share one global scope, ESLint's `no-undef` is off and the type check is what catches a wrong cross-file call.

`npm run test` uses the built-in `node --test` on `test/*.test.mjs` (kept outside `src/` so clasp never pushes them). It covers only the pure formula-parsing logic in `ImportRangeScanner.js` — `extractImportRangeFirstArgs_`, `parseFirstArg_`, `resolveImportRangeSourceId_`. The source files have no exports, so the tests load them into a `node:vm` context. Anything touching `SpreadsheetApp` / `UrlFetchApp` (the rest of the library) is tested manually from the Apps Script editor — see README.md's "Testing" section for the checklist. Don't add mock-everything tests for those wrappers.

## Deploying

`clasp push` and cutting a new library deployment both require explicit user confirmation before running — never do either without being asked in that moment (per the user's standing instruction not to auto-push Apps Script code).
