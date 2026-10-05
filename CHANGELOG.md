# Changelog

## 0.2.0 (2026-10-05)

First public release, as `formula-engine`. Version 0.1.0 was an internal release named `@mike/formula` and was never published to npm.

Breaking changes:

- Every values `hash` changes: the configuration version, which is part of the hash, is now `formula-engine/1;compiled-js;formulajs@4.6.1;canon15` (it was `app-formula/1;compiled-js;formulajs@4.6.1;canon15`). Values and `inputHash` do not change. The synthetic book (seed 1, 3,400 rows, clock `2026-10-04T12:00:00Z` with seed 7) hashes to `18be025cb4581543…`; with the 0.1.0 version string the same values still hash to the 0.1.0 `9d7414efbda34a63…` (pinned by a test).
- The subpaths `./internal` and `./synth` move to `./unstable/internal` and `./unstable/synth` and are outside semver.
- Input of the wrong shape throws `InputError` instead of being accepted silently or failing with an unrelated error: cells given as bare values, unknown fields in cells, names and options, cell addresses outside `A1:XFD1048576` or not in upper-case A1 style, non-string formulas, malformed dates and clocks.
- Node 20.19 or newer (0.1.0 asked for Node 24). Installing from git no longer builds `dist` (there is no `prepare` script).

Added:

- `compute(book)` and `openSession(book)` work without a clock: the current moment and a random seed, recorded in `result.clock` / `session.clock`; passing that clock again reproduces the result. Calls with a clock work as before.
- `InputError` (a `TypeError` with `path`) with messages that name the place and the problem.
- When `new Function` is blocked (a Content Security Policy without `'unsafe-eval'`), `compute` and sessions throw an `EvalError` that says so.
- The library runs outside Node: SHA-256 in plain JS (digests identical to `node:crypto`), no `node:*` imports, `performance` and `crypto` used only when present. A test bundles it for the browser and runs it in an empty JavaScript context.

Changed:

- Documentation, comments, the contract's JSDoc, messages and test names are in English. `debugGeneratedCode` marks cells in a cycle with `(cycle)`.
- `exports` use the `types` and `default` conditions, so `require()` works on Node versions that can require ES modules.
- The LibreOffice oracle's answers are stored in `test/oracle/libreoffice.json`; `npm test` no longer needs LibreOffice, `npm run test:oracle` recalculates them.
- The README's code examples run as a test. CI for GitLab and GitHub. MIT license; `THIRD_PARTY.md` carries the full license texts of formula.js, jStat and bessel.

## 0.1.0 (2026-10-05)

Internal release as `@mike/formula`, extracted from a private application together with its history. Contract: `compute`, `validateFormula`, `validateName`, `openSession` (`setCells`, `setNames`, `setClock`, `explain`, `result`, `close`), check codes, cell errors with reasons, volatile cells, values and input hashes, external functions through `externals`. Whitelist: SUM, SUMIF, SUMIFS, COUNTIFS, IF, IFERROR, XLOOKUP, ROUND, EOMONTH, DATE, TEXT, TRUE, FALSE, TODAY, NOW, RAND, RANDBETWEEN; ROUND, EOMONTH, DATE and SUM from formulajs 4.6.1. Configuration version `app-formula/1;compiled-js;formulajs@4.6.1;canon15`; built with `tsc` to `dist` with subpaths `.`, `./contract`, `./synth`, `./internal`.
