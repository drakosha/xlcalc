# formula-engine

A formula engine for spreadsheet-like workbooks in JavaScript. You give it a book (sheets of cells with values and Excel formulas, plus named ranges); it gives back the computed values, the errors with their reasons, and a deterministic hash of the values. Formulas use Excel syntax: English function names, commas between arguments, a decimal point.

What it does differently:

- **Compiles the book to JavaScript.** Each formula becomes a JavaScript function (one `new Function` call per book), and cells are evaluated over the reference graph in topological order. A session keeps the compiled book and recomputes only the cells that depend on an edit. On a synthetic book with 3,400 data rows and 3,907 formulas, mostly SUMIFS over whole columns, a cold compute takes 0.2–0.3 s and an edit of a criterion cell under 10 ms (p95) on an Apple M5 Pro with Node 26; `test/budgets.test.ts` enforces 1.5 s and 100 ms.
- **Deterministic values and hashes.** The same book with the same clock gives the same values and the same `hash` (SHA-256 of a canonical text of all values plus the engine configuration version), regardless of input order, process or time zone. `inputHash` identifies the input.
- **A clock for volatile functions.** `TODAY` and `NOW` come from `clock.now` (UTC), `RAND` and `RANDBETWEEN` from `clock.seed` and the cell address. Without a clock the engine takes the current moment and a random seed, and `result.clock` reproduces the result.
- **A whitelist of functions, checked against Microsoft Excel.** A function is admitted only when it gives Excel's answer on every one of its test cases, typical calls and edge cases (text and booleans in numeric arguments, blanks, errors, date text, ranges with mixed content), computed in Excel 16 with the en-US locale; [Functions](#functions) has the list. Any other function is reported as `unknown_function` (`#NAME?`) rather than computed with different semantics.
- **Errors are values.** An invalid formula never throws: its cell gets an Excel error, and `issues` says why, with a check code and the position in the formula.
- **No Node APIs.** It runs in Node 20.19 or newer and in browsers (see [Limitations](#limitations) about Content Security Policy).

Function implementations are ported from [formula.js](https://github.com/formulajs/formulajs) (MIT), [jStat](https://github.com/jstat/jstat) (MIT) and [bessel](https://github.com/SheetJS/bessel) (Apache-2.0), with fixes verified against Microsoft Excel; the rest (criteria and database functions, the lookups, `TEXT` and `VALUE`, dates, `IF`, the volatile functions) are written here.

## Install

```sh
npm install formula-engine
```

The package is ESM with TypeScript types included. From CommonJS use `await import('formula-engine')`, or `require()` on Node versions that can require ES modules (20.19+ and 22.12+). It has no dependencies.

## Quick start

```js
import { compute } from 'formula-engine'

const cells = {
  A1: { value: 'North' }, B1: { value: 120 },
  A2: { value: 'South' }, B2: { value: 80 },
  A3: { value: 'North' }, B3: { value: 45.5 },
  D1: { formula: 'SUMIFS(B:B,A:A,"North")' },
  D2: { formula: 'XLOOKUP("South",A1:A3,B1:B3)' },
  D3: { formula: 'INDIRECT("B2")' },
}
const { sheets, issues } = compute({ sheets: [{ name: 'Sales', cells }] })
console.log(sheets[0].cells.D1.value) // → 165.5
console.log(sheets[0].cells.D2.value) // → 80
console.log(sheets[0].cells.D3.value, issues[0].reason) // → { error: 'NAME' } unknown function INDIRECT
```

A formula is stored without the leading `=` in `formula`, separately from `value`, so a string from your data never turns into a formula. Dates are passed explicitly as `{ date: '2026-01-31' }` and become Excel serial numbers; `format` and `meta` pass through to the result untouched.

## Usage

### Sessions

A session holds a computed book. `setCells` applies edits atomically, recomputes only what depends on them and returns the cells whose values changed; `explain` shows where a value comes from.

```js
import { openSession } from 'formula-engine'

const session = openSession({
  sheets: [{
    name: 'Plan',
    cells: {
      A1: { value: 100 }, A2: { value: 250 },
      B1: { formula: 'SUM(A1:A2)' },
      B2: { formula: 'ROUND(B1*1.2,0)' },
    },
  }],
})

console.log(session.setCells([{ sheet: 'Plan', cell: 'A2', input: { value: 300 } }]))
// → [
//     { sheet: 'Plan', cell: 'A2', value: 300 },
//     { sheet: 'Plan', cell: 'B1', value: 400 },
//     { sheet: 'Plan', cell: 'B2', value: 480 }
//   ]
console.log(session.explain('Plan', 'B2').text)
// → 'Plan'!B2 = 480  ← ROUND(B1*1.2,0)
//     'Plan'!B1 = 400  ← SUM(A1:A2)
//       'Plan'!A1:A2 (2 cells)
console.log(session.result().hash === session.result().hash) // → true
session.close()
```

`setNames` replaces the named ranges, `result()` returns the same shape as `compute`, and `close()` releases the book. A full `compute` of the edited book with `session.clock` gives the same hash as `session.result()`; the test suite checks this after each of 200 random edits.

### The clock and volatile functions

```js
import { compute } from 'formula-engine'

const book = { sheets: [{ name: 'S', cells: { A1: { formula: 'TODAY()' }, A2: { formula: 'NOW()' }, A3: { formula: 'RANDBETWEEN(1,6)' } } }] }

const fixed = compute(book, { clock: { now: '2026-01-31T18:00:00Z', seed: 42 } })
console.log(fixed.sheets[0].cells.A1.value, fixed.sheets[0].cells.A2.value) // → 46053 46053.75

const now = compute(book) // the current moment and a random seed
console.log(compute(book, { clock: now.clock }).hash === now.hash) // → true
console.log(now.volatile.map((c) => c.cell).join(',')) // → A1,A2,A3
```

`volatile` lists the cells that depend on volatile functions, directly or through other cells. A session opened without a clock keeps the moment it was opened until `setClock`; `openSession(book, { live: true })` takes a fresh clock on every recomputation instead.

### Checking formulas

`validateFormula` checks a formula without a book: it returns the canonical text (sheet names quoted, function names upper-case) and what the formula uses, or a check code with a message and a position.

```js
import { validateFormula, validateName } from 'formula-engine'

console.log(validateFormula("sum(Sales!b1:b3, 'Q1 data'!A1)", { sheets: ['Sales', 'Q1 data'] }).canonical)
// → SUM('Sales'!B1:B3,'Q1 data'!A1)
const bad = validateFormula('SUM(B1:B3', { sheets: ['Sales'] })
console.log(bad.ok, bad.code, bad.message, bad.at) // → false syntax expected ), found end of formula 9
console.log(validateName('Q1').ok, validateName('tax_rate').ok) // → false true
```

Check codes: `empty`, `syntax`, `too_long` (over 8,192 characters), `too_deep` (over 64 levels of nesting), `external_ref` (a reference to another workbook), `unknown_function`, `unknown_sheet`, `unknown_name`, `unsupported` (array arithmetic, or an external function called with non-literal arguments). In `compute`, an invalid formula gives `#NAME?` for an unknown function or name, `#REF!` for an unknown sheet or an external reference, and `#ERROR!` otherwise.

### Errors

Cell errors come out as `{ error: code }` with codes `DIV0`, `NA`, `NAME`, `NUM`, `REF`, `VALUE`, `CYCLE`, `SPILL` and `ERROR`. Every cell with an error is listed in `issues` with a reason, which names the cell an error came from (`division by zero, from 'S'!A2`). A circular reference gives `#CYCLE!` to the cells in the cycle and to everything that depends on them; the rest of the book is computed.

Input of the wrong shape is a programming error and throws `InputError` (a `TypeError` with `path`), for example `book.sheets[0].cells["b2"]: invalid cell address "b2": expected A1 style such as "B12" (column letters must be upper-case: "B2")`. This covers wrong types, malformed or out-of-grid addresses, duplicate sheet names (compared case-insensitively, as in Excel), unknown fields, bad dates and clocks, and unknown sheets in session edits.

### Named ranges

```js
import { compute } from 'formula-engine'

const result = compute({
  sheets: [
    { name: 'Data', cells: { A1: { value: 'tea' }, B1: { value: 3 }, A2: { value: 'coffee' }, B2: { value: 5 } } },
    { name: 'Report', cells: { A1: { formula: 'SUMIF(items,"tea",amounts)' } } },
  ],
  names: [
    { name: 'items', ref: "'Data'!$A:$A" },
    { name: 'amounts', ref: "'Data'!$B:$B" },
  ],
})
console.log(result.sheets[1].cells.A1.value) // → 3
```

A name with `scope: 'Sheet'` is visible on that sheet only and shadows a global name. Invalid names are reported in `nameIssues`.

### External functions

Host functions that are not part of the book (a balance from a database, say) can appear in formulas. `validateFormula(text, { externalFunctions: ['BALANCE'] })` returns the canonical text of each call in `externals`; the host resolves the values in advance and passes them to `compute`:

```js
import { compute } from 'formula-engine'

const book = { sheets: [{ name: 'S', cells: { A1: { formula: 'BALANCE("2026-01")*2' } } }] }
const result = compute(book, { externals: { functions: ['BALANCE'], values: { 'BALANCE("2026-01")': 21 } } })
console.log(result.sheets[0].cells.A1.value) // → 42
```

Arguments of an external function must be literals; a call without a value gives `#N/A`.

## API

From `formula-engine`:

| export | |
|---|---|
| `compute(book, opts?)` | computes the whole book; `opts.clock`, `opts.externals`. Returns `ComputeResult`: `sheets` (values, canonical formulas, `format`, `meta`), `issues`, `nameIssues`, `volatile`, `clock`, `hash`, `inputHash`, `configVersion`, `ms` |
| `openSession(book, opts?)` | a `Session`: `setCells(edits)`, `setNames(names)`, `setClock(clock)`, `explain(sheet, cell, { depth, maxNodes, expand })`, `result()`, `close()`, `clock`. Options: `clock`, or `live: true`, and `externals` |
| `validateFormula(text, ctx?)` | `{ ok: true, canonical, functions, sheets, names, externals }` or `{ ok: false, code, message, at }`; `ctx`: `sheets`, `names`, `externalFunctions` (without them sheets and names are not checked) |
| `validateName(name)` | `{ ok: true }` or `{ ok: false, message }` |
| `engine` | the four functions above as one `FormulaEngine` object |
| `FUNCTION_NAMES`, `CONFIG_VERSION` | the whitelist; the configuration version that goes into every hash |
| `InputError` | thrown on input of the wrong shape |
| `debugGeneratedCode(book, opts?)` | the generated JavaScript in evaluation order, for debugging; not part of the contract |

`formula-engine/contract` exports the types and two constants (`MAX_FORMULA_LENGTH`, `MAX_FORMULA_DEPTH`) without any engine code, for code that only builds books or reads results.

`formula-engine/unstable/internal` (the parser, the criteria index, engine state) and `formula-engine/unstable/synth` (a seeded synthetic book for tests and benchmarks) exist for benchmark harnesses and adapters of other engines. They are outside semver and change in any release.

## Functions

The whitelist, by category (`FUNCTION_NAMES` has it in code):

- Math and trigonometry: `ABS`, `ACOS`, `ACOSH`, `ACOT`, `ACOTH`, `ARABIC`, `ASIN`, `ASINH`, `ATAN`, `ATAN2`, `ATANH`, `BASE`, `CEILING`, `CEILING.MATH`, `CEILING.PRECISE`, `COMBIN`, `COMBINA`, `COS`, `COSH`, `COT`, `COTH`, `CSC`, `CSCH`, `DECIMAL`, `DEGREES`, `EVEN`, `EXP`, `FACT`, `FACTDOUBLE`, `FLOOR`, `FLOOR.MATH`, `FLOOR.PRECISE`, `GCD`, `INT`, `ISO.CEILING`, `LCM`, `LN`, `LOG`, `LOG10`, `MDETERM`, `MOD`, `MROUND`, `MULTINOMIAL`, `ODD`, `PERCENTOF`, `PI`, `POWER`, `PRODUCT`, `QUOTIENT`, `RADIANS`, `RAND`, `RANDBETWEEN`, `ROMAN`, `ROUND`, `ROUNDDOWN`, `ROUNDUP`, `SEC`, `SECH`, `SERIESSUM`, `SIGN`, `SIN`, `SINH`, `SQRT`, `SQRTPI`, `SUM`, `SUMIF`, `SUMIFS`, `SUMPRODUCT`, `SUMSQ`, `SUMX2MY2`, `SUMX2PY2`, `SUMXMY2`, `TAN`, `TANH`, `TRUNC`

- Statistical: `AVEDEV`, `AVERAGE`, `BINOM.DIST.RANGE`, `BINOMDIST`, `COUNTA`, `COUNTBLANK`, `COUNTIFS`, `F.DIST.RT`, `FISHER`, `GAMMALN`, `GAMMALN.PRECISE`, `GAUSS`, `MAX`, `MEDIAN`, `MIN`, `NORM.DIST`, `NORM.S.DIST`, `NORMDIST`, `PHI`, `PROB`

- Logical: `AND`, `FALSE`, `IF`, `IFERROR`, `IFNA`, `IFS`, `NOT`, `OR`, `SWITCH`, `TRUE`, `XOR`

- Information: `ERROR.TYPE`, `ISBLANK`, `ISERR`, `ISERROR`, `ISEVEN`, `ISLOGICAL`, `ISNA`, `ISNONTEXT`, `ISNUMBER`, `ISODD`, `ISTEXT`, `N`, `NA`, `TYPE`

- Text: `ARRAYTOTEXT`, `BAHTTEXT`, `CLEAN`, `CONCAT`, `CONCATENATE`, `DOLLAR`, `EXACT`, `FIND`, `FINDB`, `FIXED`, `LEFT`, `LEFTB`, `LEN`, `LENB`, `LOWER`, `MID`, `MIDB`, `NUMBERVALUE`, `PROPER`, `REPLACE`, `REPLACEB`, `REPT`, `RIGHT`, `RIGHTB`, `SEARCH`, `SEARCHB`, `SUBSTITUTE`, `T`, `TEXT`, `TEXTAFTER`, `TEXTBEFORE`, `TEXTJOIN`, `TRIM`, `UNICHAR`, `UNICODE`, `UPPER`, `VALUE`, `VALUETOTEXT`

- Date and time: `DATE`, `DATEDIF`, `DATEVALUE`, `DAY`, `DAYS`, `DAYS360`, `EDATE`, `EOMONTH`, `HOUR`, `ISOWEEKNUM`, `MINUTE`, `MONTH`, `NETWORKDAYS`, `NETWORKDAYS.INTL`, `NOW`, `SECOND`, `TIME`, `TIMEVALUE`, `TODAY`, `WEEKDAY`, `WEEKNUM`, `WORKDAY`, `WORKDAY.INTL`, `YEAR`, `YEARFRAC`

- Financial: `ACCRINT`, `ACCRINTM`, `AMORDEGRC`, `AMORLINC`, `COUPDAYBS`, `COUPDAYSNC`, `COUPNCD`, `COUPNUM`, `COUPPCD`, `CUMIPMT`, `CUMPRINC`, `DB`, `DDB`, `DISC`, `DOLLARDE`, `DOLLARFR`, `DURATION`, `EFFECT`, `FV`, `FVSCHEDULE`, `INTRATE`, `IPMT`, `IRR`, `ISPMT`, `MDURATION`, `MIRR`, `NOMINAL`, `NPER`, `NPV`, `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`, `PDURATION`, `PMT`, `PPMT`, `PRICE`, `PRICEDISC`, `PRICEMAT`, `PV`, `RATE`, `RECEIVED`, `RRI`, `SLN`, `SYD`, `TBILLEQ`, `TBILLPRICE`, `TBILLYIELD`, `XIRR`, `XNPV`, `YIELD`, `YIELDDISC`, `YIELDMAT`

- Lookup and reference: `ADDRESS`, `CHOOSE`, `COLUMN`, `COLUMNS`, `HLOOKUP`, `HYPERLINK`, `INDEX`, `LOOKUP`, `MATCH`, `ROW`, `ROWS`, `VLOOKUP`, `XLOOKUP`, `XMATCH`

- Database: `DAVERAGE`, `DCOUNT`, `DCOUNTA`, `DGET`, `DMAX`, `DMIN`, `DPRODUCT`, `DSTDEV`, `DSTDEVP`, `DSUM`, `DVAR`, `DVARP`

- Engineering: `BESSELI`, `BESSELJ`, `BESSELK`, `BESSELY`, `BIN2DEC`, `BIN2HEX`, `BIN2OCT`, `BITAND`, `BITLSHIFT`, `BITOR`, `BITRSHIFT`, `BITXOR`, `COMPLEX`, `CONVERT`, `DEC2BIN`, `DEC2HEX`, `DEC2OCT`, `DELTA`, `ERF`, `ERF.PRECISE`, `ERFC`, `ERFC.PRECISE`, `GESTEP`, `HEX2BIN`, `HEX2DEC`, `HEX2OCT`, `IMABS`, `IMAGINARY`, `IMARGUMENT`, `IMCONJUGATE`, `IMCOS`, `IMCOT`, `IMCSCH`, `IMDIV`, `IMEXP`, `IMLN`, `IMLOG10`, `IMLOG2`, `IMPOWER`, `IMPRODUCT`, `IMREAL`, `IMSEC`, `IMSECH`, `IMSIN`, `IMSINH`, `IMSQRT`, `IMSUB`, `IMSUM`, `IMTAN`, `OCT2BIN`, `OCT2DEC`, `OCT2HEX`

Status: every function in the list matches Excel en-US on all of its cases in the coverage measurement (`docs/coverage-formulajs.md`) and in the Excel case sets (`docs/excel-cases/`), which also cover the functions formula.js lacks (`XLOOKUP`, `XMATCH`, `TEXTBEFORE`, `ADDRESS` and others). `TODAY`, `NOW`, `RAND` and `RANDBETWEEN` (volatile) have their own semantics and clock tests instead.

How the acceptance works: the measurement generated 5,275 cases for 402 Excel functions, typical calls (mostly the examples of Excel's function reference) and edge cases, and computed them in Excel 16.113 for Mac with the en-US locale and English UI; the answers are stored in `docs/coverage-formulajs/excel-en.json`. `test/excel.test.ts` computes the cases of every whitelisted function with this engine, in one book with the measurement's data sheet, and compares: numbers within |a−b| ≤ 1e-9·max(1,|a|), text, booleans and error codes exactly. The test needs neither Excel nor LibreOffice. Rules that are not one function's (numbers as text, long number literals, criteria over whole columns) are checked the same way on Excel's answers to 959 probes in `docs/excel-probes.json` (`test/excel-probes.test.ts`). The older cases of `test/semantics-cases.ts` (expected answers from Excel's documentation, cross-checked with LibreOffice) still run too.

Notes on semantics:

- Arguments are coerced by kind, the same way for every function of a group: numeric arguments read numeric text and booleans, a range in `SUM`, `AVERAGE`, `MAX` and the like contributes only its numbers while a direct argument is coerced, date arguments read date text as Excel en-US does (`2025-01-31`, `1/31/2025`, `31-Jan-2025`), and `"TRUE"`/`"FALSE"` text is a boolean where a boolean is expected.
- Criteria (`SUMIF`, `SUMIFS`, `COUNTIFS`) compare text case-insensitively but accent-sensitively, support `*` and `?` (where `~` takes the next character as it is; text without `*` or `?` is compared as it is, `~` included), read numbers and ISO dates inside criterion text as numbers (`">=2026-01-01"`), and a number criterion also matches text that reads as that number (`2` matches `"2"`), as in Excel. Their ranges must have the same referenced size (else `#VALUE!`); whole columns of sheets with different numbers of rows fit, and blank cells of a whole column count where every criterion matches a blank (`COUNTIFS(A:A,"")`).
- Numbers become text (`&`, `CONCAT`, text arguments) as in Excel: rounded to 15 significant digits, in positional notation when that takes at most 20 characters, the sign aside (`0.0000000745`, `10000000000000000000`), else in scientific notation with an exponent of two digits at least (`7.4500104300146E-06`, `1E+20`). `TEXT` with the `General` format has Excel's other rule, at most 11 characters (`0.333333333`, `1.23457E-05`).
- Number literals and numeric text keep 15 significant digits and drop the rest, as Excel reads them (`123456789012345678` is 123456789012345000); a result smaller than the smallest normal number (2.2250738585072E-308) is 0.
- `IF`, `IFERROR`, `IFNA` and `SWITCH` do not compute the branch not taken.
- Lookups compare within a type (numbers, text, booleans), text in any case, and never match blanks or errors in the data. An exact match in `VLOOKUP`, `HLOOKUP` and `MATCH` (and `XLOOKUP`, `XMATCH` in match mode 2) takes `*`, `?` and `~`, where `~` escapes any character. An approximate match is Excel's binary search over the referenced cells, so it gives Excel's answer on unsorted, mixed and whole-column data too (checked on about 200 such cases); `MATCH` with type −1 scans. A blank lookup value is 0 in `VLOOKUP`, `HLOOKUP`, `MATCH` and `LOOKUP` and finds a blank cell in `XLOOKUP` and `XMATCH`. An error in a column index or match type taken from a cell is `#REF!`, as in Excel.
- Database functions (`DSUM` and the rest) read records by rows under a header row. Criteria are written as in `SUMIFS`, except that text without an operator matches the values that begin with it (`App` selects `Apple`); criteria rows are alternatives, conditions in a row all hold. A condition under a header that is not a field name holds for every record when it is `TRUE` (Excel's computed criterion) and for none otherwise.
- `VALUE` reads currency, thousands groups, parentheses, `%`, exponents, `1 1/2`, dates and times as Excel en-US does, times past 24 hours too (`25:00`); `NUMBERVALUE` removes whitespace and the group separators, then reads the rest the same way.
- `CHAR` and `CODE` are not in the whitelist: Excel's answer depends on the platform and its language (Windows-1252 in Excel for Windows en-US, Mac Roman in Excel for Mac with the English interface, Mac Cyrillic with the Russian one). `UNICHAR` and `UNICODE` work with code points and are in the whitelist.
- Dates follow Excel's 1900 system, phantom 1900-02-29 included: `DATE(1900,2,29)` is 60.

## Limitations

- **No arrays.** Array arithmetic (`A1:A3*2`), array constants (`{1,2}`), dynamic arrays and spilling are not supported: such a formula gets the check code `unsupported` or `syntax` and `#ERROR!`. `#SPILL!` exists as an error code but never occurs. Functions with an array answer (`TRANSPOSE`, `MMULT`, `TAKE`, …) are not in the whitelist.
- **Content Security Policy.** Compiling uses `new Function`. In a browser with a CSP that does not allow `'unsafe-eval'`, `compute` and `openSession` throw an `EvalError` saying so; `validateFormula`, `validateName` and the types still work.
- **Known differences from Excel:**
  - `XLOOKUP` and `XMATCH` with binary search (search mode 2 or −2) scan linearly: the same answer on sorted data, where Excel's binary search is defined. Match mode 3 (regular expressions) is `#VALUE!`.
  - `INDEX` returns a value, not a reference: a whole row or column (`INDEX(A1:C3,0,2)`) is `#VALUE!` even inside `SUM`, which in Excel sums it. `ROW` and `COLUMN` of several rows or columns give the first one, as the cell shows it; Excel passes the whole list to an enclosing function (`SUM(ROW(A1:A3))` is 6 there).
  - `CONCAT` and `TEXTJOIN` longer than 32,767 characters are `#VALUE!`; Excel answers `#CALC!`, which this engine has no code for. `&` and `CONCATENATE` cut the text at 32,767 characters, as Excel does.
  - `TEXTBEFORE` and `TEXTAFTER` take one delimiter (there are no array constants).
  - `LOOKUP` with a result range shorter than the lookup range is `#REF!` past its end; Excel reads on past the end of the range.
  - Text order for `<` and `>` uses `Intl.Collator` (case-insensitive, accent-sensitive), which may differ from Excel for some characters; without `Intl` it falls back to comparing lower-cased code units.
  - `TEXT` knows en-US formats only; exponent formats (`0.00E+00`), colours and bracketed conditions give `#VALUE!`.
  - Date text without a year (`1/31`, which Excel completes with the current year) is not read as a date.
  - `#NULL!` is read as `#ERROR!`.
- **Syntax.** English function names and separators only (no `;` or decimal commas), no R1C1 references, no references to other workbooks, no 3D references or structured table references. Functions outside the whitelist (`INDIRECT`, `OFFSET`, `LET`, …) give `#NAME?`.
- **Output precision.** Values in the result and in the hash are rounded to 15 significant digits, as Excel shows them; inside, computation keeps full double precision.

## Development

```sh
npm ci
npm run typecheck
npm test                # build, then all tests against src
npm run test:dist       # build, then all tests against the built dist
npm run test:oracle     # recalculate the LibreOffice answers with soffice and check them
npm run check:package   # build, then publint and attw on the packed tarball
npm run pack:release    # build, then the tarball in release/
FORMULA_BENCH=1 npx vitest run test/budgets.test.ts   # budgets on the full synthetic book
```

The tests cover the whitelist against Excel's answers on the coverage measurement's cases (`test/excel.test.ts`) and on the probes (`test/excel-probes.test.ts`), the argument layer rule by rule, the semantics cases against Excel's documentation and the stored LibreOffice answers (`test/oracle/libreoffice.json`, refreshed by `npm run test:oracle`, which needs LibreOffice's `soffice` on `PATH`; the usual tests need no office suite); determinism (shuffled input, a child process, three time zones, cents over 3,400 amounts); references, errors, volatile functions, parsing and the canonical form; sessions (200 random edits against full computations) and `explain`; budgets; SHA-256 against test vectors and `node:crypto`; a browser bundle run in an empty JavaScript context and with code generation disabled; the published package (exports, tarball contents, a consumer importing it by name). The code blocks of this README run as a test too.

Functions live in `src/functions/<category>.ts` (`math`, `statistical`, `logical`, `information`, `text`, `date`, `financial`, `lookup`, `engineering`, `database`; old names such as `NORMDIST` sit next to their modern versions), each file with its own registry, a large category assembled from `src/functions/<category>/`; `src/functions.ts` assembles them. A function declares the kinds of its arguments (`args: 'n i?'`, see `src/args.ts`), and the shared argument layer coerces them, so an implementation gets numbers, text, lists or grids and never repeats the coercion rules. Ported code keeps its attribution in the file header. A function joins its registry when `test/excel.test.ts` passes for it; the coverage measurement (`bench/coverage`, `docs/coverage-formulajs.md`) lists the cases and the classes of the remaining formula.js functions.

## License

MIT. Function implementations are ported from [formula.js](https://github.com/formulajs/formulajs) (MIT, Copyright (c) 2014 Sutoiku, Inc.), [jStat](https://github.com/jstat/jstat) (MIT, Copyright (c) 2013 jStat) and [bessel](https://github.com/SheetJS/bessel) (Apache-2.0, Copyright (C) 2013-present SheetJS); [THIRD_PARTY.md](THIRD_PARTY.md) has their license texts and the files they are in.
