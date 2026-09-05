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

## File layout

- `Main.js` — public API: `autoApproveImportRanges`, `installAutoApproveTrigger`, `logImportRangeFormulas`.
- `ImportRangeScanner.js` — `IMPORTRANGE` detection: `findImportRangeSourceIds` (public) plus the formula-walking internals (`forEachImportRangeFormula_`, `extractImportRangeFirstArgs_`, `parseFirstArg_`, `resolveImportRangeSourceId_`).
- `PermissionGranter.js` — `grantImportRangeAccess_`, the only file that calls `UrlFetchApp`.

## Testing

No automated test framework in Apps Script. Test manually from the Apps Script editor — see README.md's "Testing" section for the checklist.

## Deploying

`clasp push` and cutting a new library deployment both require explicit user confirmation before running — never do either without being asked in that moment (per the user's standing instruction not to auto-push Apps Script code).
