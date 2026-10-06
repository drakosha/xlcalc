# Changelog

## Unreleased

Breaking changes:

- Every values `hash` changes: the configuration version is now `formula-engine/2;compiled-js;canon15` (it was `formula-engine/1;compiled-js;formulajs@4.6.1;canon15`). The synthetic book (seed 1, 3,400 rows, clock `2026-10-04T12:00:00Z` with seed 7) keeps its values and hashes to `8d63bad2c247ae26…`; with the 0.2.0 and 0.1.0 version strings the same values still hash to `18be025cb4581543…` and `9d7414efbda34a63…` (pinned by a test).
- No runtime dependencies: `@formulajs/formulajs` is no longer a dependency. Its code, and jStat's, is ported into `src/functions/` with attribution (`THIRD_PARTY.md`).
- `SUMIF`, `SUMIFS` and `COUNTIFS` compare the referenced sizes of their ranges: ranges of different sizes are `#VALUE!` (a whole column next to a shorter range was computed over the occupied rows), whole columns of sheets with different numbers of rows work (`SUMIFS('S2'!A:A,'S3'!B:B,">0")` was `#VALUE!`), and `COUNTIFS` counts the blank cells of whole columns when every criterion matches a blank (`COUNTIFS(A:A,"")` is 1,048,576 minus the filled cells).
- Values that change because they now match Excel en-US: a number criterion matches text that reads as that number (`COUNTIFS(A:A,2)` counts `"2"`); `DATE(1900,2,29)` is 60, `EOMONTH` reads serial 0 as January 1900 and 60 as February 1900 (`EOMONTH(0,1)` is 59, `EOMONTH(60,1)` is 91) and reads date text (`EOMONTH("2025-01-31",1)`); a month out of range in `DATE` and `EOMONTH` is `#NUM!` (was `#VALUE!`); `ROUND` with a huge digit count returns the number (was `#NUM!`); `TEXT` formats fractions (`TEXT(4.34,"# ?/?")` is `"4 1/3"`) and gives `""` for an empty format; `0^0` is `#NUM!` and `0^-1` `#DIV/0!`.

Added:

- 90 functions, ported from formula.js and matching Excel en-US on all of the coverage measurement's cases: `ABS`, `ACOS`, `ACOSH`, `ACOTH`, `AND`, `ASIN`, `ASINH`, `ATAN`, `ATANH`, `AVEDEV`, `AVERAGE`, `BINOM.DIST.RANGE`, `BINOMDIST`, `CEILING`, `CEILING.MATH`, `CEILING.PRECISE`, `CHOOSE`, `CLEAN`, `COLUMNS`, `CONCAT`, `CONCATENATE`, `COSH`, `COUNTA`, `COUNTBLANK`, `CSCH`, `DEGREES`, `ERROR.TYPE`, `EVEN`, `EXACT`, `EXP`, `F.DIST.RT`, `FIND`, `FISHER`, `FIXED`, `FLOOR`, `FLOOR.MATH`, `FLOOR.PRECISE`, `GAMMALN`, `GAMMALN.PRECISE`, `GAUSS`, `IFS`, `INT`, `ISBLANK`, `ISERR`, `ISERROR`, `ISLOGICAL`, `ISNA`, `ISNONTEXT`, `ISNUMBER`, `ISPMT`, `ISTEXT`, `LEFT`, `LEN`, `LN`, `LOG`, `LOG10`, `LOWER`, `MAX`, `MEDIAN`, `MIN`, `N`, `NA`, `NORM.DIST`, `NORM.S.DIST`, `NORMDIST`, `NOT`, `NPV`, `ODD`, `OR`, `PDURATION`, `PHI`, `PI`, `POWER`, `PROB`, `PRODUCT`, `RADIANS`, `REPT`, `RIGHT`, `ROUNDDOWN`, `ROUNDUP`, `ROWS`, `SECH`, `SIGN`, `SINH`, `SQRT`, `T`, `TRIM`, `TRUNC`, `UNICODE`, `XOR`. The whitelist has 107 functions.
- `test/excel.test.ts`: every whitelisted function is computed on the coverage measurement's cases and compared with Excel en-US (answers stored in `docs/coverage-formulajs/excel-en.json`); a function is in the whitelist only if all of its cases match.

Changed:

- Functions are organised by category in `src/functions/`, on a shared argument layer (`src/args.ts`) that coerces arguments by kind for every function of a group: numbers, text, booleans, dates with en-US date text, lists by the SUM, AVERAGEA and strict rules, values with their origin, grids, paired arrays, holidays.
- `XLOOKUP` compares the referenced sizes of its ranges: whole columns of sheets with different numbers of rows work (they gave `#VALUE!`).

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
