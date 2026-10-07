# xlcalc

A formula engine for spreadsheet-like workbooks in JavaScript. You give it a book (sheets of cells with values and Excel formulas, plus named ranges); it gives back the computed values, the errors with their reasons, and a deterministic hash of the values. Formulas use Excel syntax: English function names, commas between arguments, a decimal point; [Localization](#localization) shows them and reads them in another language, as Excel does.

What it does differently:

- **Compiles each formula once.** By default a formula becomes a tree of JavaScript closures, with no code generated, so a strict Content Security Policy does not stand in the way; with `execution: 'eval'` it becomes JavaScript text, and one `new Function` call compiles the book. Both modes give the same values and hashes, and the tests compute every case book both ways. Cells are evaluated over the reference graph in topological order. A session keeps the compiled book and recomputes only the cells that depend on an edit. On a synthetic book with 3,400 data rows and 3,907 formulas, mostly SUMIFS over whole columns, a cold compute takes 0.2–0.3 s and an edit of a criterion cell under 10 ms (p95) in either mode on an Apple M5 Pro with Node 26; `test/budgets.test.ts` enforces 1.5 s and 100 ms.
- **Deterministic values and hashes.** The same book with the same clock gives the same values and the same `hash` (SHA-256 of a canonical text of all values plus the engine configuration version), regardless of input order, process or time zone. `inputHash` identifies the input.
- **A clock for volatile functions.** `TODAY` and `NOW` come from `clock.now` (UTC), `RAND` and `RANDBETWEEN` from `clock.seed` and the cell address. Without a clock the engine takes the current moment and a random seed, and `result.clock` reproduces the result.
- **A whitelist of functions, checked against Microsoft Excel.** A function is admitted only when it gives Excel's answer on every one of its test cases, typical calls and edge cases (text and booleans in numeric arguments, blanks, errors, date text, ranges with mixed content), computed in Excel 16 with the en-US locale; [Functions](#functions) has the list. Any other function is reported as `unknown_function` (`#NAME?`) rather than computed with different semantics.
- **Arrays inside a formula.** Operators and functions work element by element over ranges and arrays with Excel 365's rules (`SUM(B:B*C:C)`, `SUMPRODUCT((A1:A9="x")*B1:B9)`, `SUM(IF(A1:A9>0,A1:A9))`), array constants `{1,2;3,4}` are read, and functions with an array answer (`SORT`, `FILTER`, `UNIQUE`, `SEQUENCE`, `MMULT`, …) can be used inside a formula. A formula whose result is an array spills into the cells to its right and below, as in Excel 365 (`#SPILL!` where they are taken), and `E2#` refers to the range it filled (see [Arrays](#arrays)).
- **Errors are values.** An invalid formula never throws: its cell gets an Excel error, and `issues` says why, with a check code and the position in the formula.
- **No Node APIs.** It runs in Node 20.19 or newer and in browsers (see [Limitations](#limitations) about Content Security Policy).

Function implementations are ported from [formula.js](https://github.com/formulajs/formulajs) (MIT), [jStat](https://github.com/jstat/jstat) (MIT) and [bessel](https://github.com/SheetJS/bessel) (Apache-2.0), with fixes verified against Microsoft Excel; the rest (criteria and database functions, the lookups, `TEXT` and `VALUE`, dates, `IF`, the volatile functions) are written here.

## Install

```sh
npm install xlcalc
```

The package is ESM with TypeScript types included. From CommonJS use `await import('xlcalc')`, or `require()` on Node versions that can require ES modules (20.19+ and 22.12+). It has no dependencies.

## Quick start

```js
import { compute } from 'xlcalc'

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
import { openSession } from 'xlcalc'

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

A session also restructures its book and keeps its styles, so a host can hold the book in the session alone and save `book()` as a version (the session's own objects: do not modify them).

```js
import { openSession } from 'xlcalc'

const session = openSession({
  sheets: [{ name: 'Plan', cells: { A1: { value: 100 }, A2: { value: 250 }, A3: { formula: 'SUM(A1:A2)', style: 0 } } }],
  styles: [{ font: { bold: true } }],
})
const r = session.restructure({ op: 'deleteRows', sheet: 'Plan', row: 1 })
console.log(r.mapCell({ sheet: 'Plan', cell: 'A3' }), r.changes)
// → { sheet: 'Plan', cell: 'A2' } [ { sheet: 'Plan', cell: 'A2', value: 250 } ]
session.setLayout([{ sheet: 'Plan', columns: 'A', set: { width: 14 } }])
console.log(session.book().sheets[0])
// → {
//     name: 'Plan',
//     cells: { A1: { value: 250 }, A2: { formula: 'SUM(A1:A1)', style: 0 } },
//     columns: { A: { width: 14 } }
//   }
```

`restructure(op)` takes the operations of [Structure operations](#structure-operations) and recomputes; besides `restructure`'s result it returns `changes`, the cells whose value differs from the one `mapCell` carried there, so an interface that moves its shown values by `mapCell` and applies `changes` is up to date. It builds the state again, at the cost of a cold computation (190 ms for the 3,400-row synthetic book). Styles change through the session and recompute nothing: a cell's `style` (or `format`, `meta`) through `setCells`, the style table through `setStyles`, column widths, row heights, row and column styles and formats and merged areas through `setLayout`; only a merged area added or removed over an array's spill range recomputes that array, and `setLayout` returns the changes.

### The clock and volatile functions

```js
import { compute } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { formula: 'TODAY()' }, A2: { formula: 'NOW()' }, A3: { formula: 'RANDBETWEEN(1,6)' } } }] }

const fixed = compute(book, { clock: { now: '2026-01-31T18:00:00Z', seed: 42 } })
console.log(fixed.sheets[0].cells.A1.value, fixed.sheets[0].cells.A2.value) // → 46053 46053.75

const now = compute(book) // the current moment and a random seed
console.log(compute(book, { clock: now.clock }).hash === now.hash) // → true
console.log(now.volatile.map((c) => c.cell).join(',')) // → A1,A2,A3
```

`volatile` lists the cells that depend on volatile functions, directly or through other cells, and on the clock's year through date text without a year ([Localization](#localization)). A session opened without a clock keeps the moment it was opened until `setClock`; `openSession(book, { live: true })` takes a fresh clock on every recomputation instead.

### Arrays

A range or an array constant where a function takes one value, or under an operator, is computed element by element, and functions that take ranges take arrays the same way. A formula whose result is an array of several values spills: its cell shows the first value and the cells to its right and below the others, as in Excel 365. In the result a spilled cell has `spillOf`, the formula's address, and the formula has `spill`, the range it fills; `E2#` refers to that range. Where a cell of the range holds a value or a formula, or a merged area meets it, the formula gives `#SPILL!` and its issue names the cell.

```js
import { compute } from 'xlcalc'

const cells = {
  A1: { value: 3 }, A2: { value: 5 }, A3: { value: 8 },
  B1: { formula: 'SUM(A1:A3*2)' },
  B2: { formula: 'SUMPRODUCT((A1:A3>4)*A1:A3)' },
  C1: { formula: 'SORT(A1:A3,1,-1)' },
  D1: { formula: 'SUM(C1#)' },
  E1: { formula: 'A1:A3*2' },
  E3: { value: 'x' },
}
const r = compute({ sheets: [{ name: 'S', cells }] })
const out = r.sheets[0].cells
console.log(out.B1.value, out.B2.value, out.D1.value) // → 32 13 16
console.log(out.C1, out.C2) // → { value: 8, formula: 'SORT(A1:A3,1,-1)', spill: 'C1:C3' } { value: 5, spillOf: 'C1' }
console.log(out.E1.value, r.issues[0].reason)
// → { error: 'SPILL' } the array of 3×1 values cannot spill: 'S'!E3 is not empty
```

A spilled cell is a value of the book: it is in `sheets[].cells` and in the hash. An array of more than 65,536 values, such as `A:A*2` (1,048,576 rows), lists its cells up to where the rest holds one value and gives that value as the formula's `spillFill` (0 below the data). The design is in `docs/arrays-incell.md`, spilling and the rules measured in Excel in `docs/spill.md`.

### Checking formulas

`validateFormula` checks a formula without a book: it returns the canonical text (sheet names quoted, function names upper-case) and what the formula uses, or a check code with a message and a position.

```js
import { validateFormula, validateName } from 'xlcalc'

console.log(validateFormula("sum(Sales!b1:b3, 'Q1 data'!A1)", { sheets: ['Sales', 'Q1 data'] }).canonical)
// → SUM('Sales'!B1:B3,'Q1 data'!A1)
const bad = validateFormula('SUM(B1:B3', { sheets: ['Sales'] })
console.log(bad.ok, bad.code, bad.message, bad.at) // → false syntax expected ), found end of formula 9
console.log(validateName('Q1').ok, validateName('tax_rate').ok) // → false true
```

Check codes: `empty`, `syntax`, `too_long` (over 8,192 characters), `too_deep` (over 64 levels of nesting), `external_ref` (a reference to another workbook), `unknown_function`, `unknown_sheet`, `unknown_name`, `unsupported` (an external function called with non-literal arguments). In `compute`, an invalid formula gives `#NAME?` for an unknown function or name, `#REF!` for an unknown sheet or an external reference, and `#ERROR!` otherwise. A formula whose result is an array is valid: it spills.

### Errors

Cell errors come out as `{ error: code }` with codes `DIV0`, `NA`, `NAME`, `NUM`, `REF`, `VALUE`, `CYCLE`, `SPILL`, `CALC` and `ERROR`. Every cell with an error is listed in `issues` with a reason, which names the cell an error came from (`division by zero, from 'S'!A2`). A circular reference gives `#CYCLE!` to the cells in the cycle and to everything that depends on them; the rest of the book is computed.

Input of the wrong shape is a programming error and throws `InputError` (a `TypeError` with `path`), for example `book.sheets[0].cells["b2"]: invalid cell address "b2": expected A1 style such as "B12" (column letters must be upper-case: "B2")`. This covers wrong types, malformed or out-of-grid addresses, duplicate sheet names (compared case-insensitively, as in Excel), unknown fields, bad dates and clocks, and unknown sheets in session edits.

### Named ranges

```js
import { compute } from 'xlcalc'

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
import { compute } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { formula: 'BALANCE("2026-01")*2' } } }] }
const result = compute(book, { externals: { functions: ['BALANCE'], values: { 'BALANCE("2026-01")': 21 } } })
console.log(result.sheets[0].cells.A1.value) // → 42
```

Arguments of an external function must be literals; a call without a value gives `#N/A`.

### Localization

Books store formulas the way xlsx files do: English function names, `,` between arguments, `.` as the decimal separator. A locale changes how formulas are typed and shown, as Excel shows them in another language. The package has `ruRU` (Excel with the Russian interface) and `enUS`; a locale is a plain object (the `Locale` type), so another language is another object of the same shape.

```js
import { ruRU, toInvariant, toLocal, validateLocalFormula } from 'xlcalc'

console.log(toLocal('IF(B1>0.5,SUM({1,2;3,4}),#N/A)', ruRU).formula) // → ЕСЛИ(B1>0,5;СУММ({1;2:3;4});#Н/Д)
console.log(toInvariant('=СУММ(A1:A3;1,5)', ruRU).formula) // → SUM(A1:A3,1.5)
const bad = validateLocalFormula('=СУММ(A1;1.5)', ruRU)
console.log(bad.ok, bad.code, bad.message, bad.at) // → false syntax unexpected "." 10
console.log(validateLocalFormula('=ОКРУГЛ(Данные!B2*1,2;0)', ruRU, { sheets: ['Данные'] }).canonical) // → ROUND('Данные'!B2*1.2,0)
```

`toLocal` and `toInvariant` return `{ ok: true, formula }` or `{ ok: false, code, message, at }` with the position in the given text. They translate function names, TRUE and FALSE, error literals, the decimal separator and the separators of arguments and of array constants (`{1;2:3;4}` in ru-RU); strings, sheet names, defined names, references and whitespace stay as they are. `toInvariant` drops a leading `=` and also reads English function names, TRUE/FALSE and error texts; a function the locale does not name (an external one, say) passes through. Text that would read differently once translated (`1.5` typed in ru-RU, where `.` is not a decimal separator) is a `syntax` error at its place. `validateLocalFormula(text, locale, ctx?)` is `validateFormula` for typed text: the same result with the canonical stored text, and positions in `text`.

The ru-RU data come from Excel itself. `bench/locale/snapshot.ts` opens books with 6,381 formulas (the coverage measurement's cases, the probes, a call of every function on Microsoft's list that the measurement lacks, literals) in Excel 16.113 for Mac with the Russian interface and reads each formula back in English and in Russian; `bench/locale/generate.ts` writes `src/locale/ru-RU.ts` from that text (`test/excel/locale/ru-RU.json`). `test/locale-excel.test.ts` checks that both translations reproduce Excel's text of all 6,381 formulas, and `test/locale-roundtrip.test.ts` that the 15,266 formulas of the measurement, the case sets and the probes come back unchanged from ru-RU. Excel 16.113 for Mac does not know `ENCODEURL`, `FILTERXML`, `WEBSERVICE`, `REGEXEXTRACT`, `REGEXREPLACE` and `REGEXTEST`, and deletes `CALL` and `REGISTER.ID` from the books it loads, so these keep their English names; so do the error literals newer than `#GETTING_DATA` (`#SPILL!`, `#CALC!` and the rest), which it does not take as literals in a formula.

A book is computed in its own locale, `BookInput.locale` (`'en-US'` by default, or `'ru-RU'`), as Excel computes with those regional settings, whoever opens the book; a viewer still sees the formulas in their own language through `toLocal`. The locale changes how text reads as a number, a date or a boolean, how numbers and booleans become text, `TEXT`'s format codes, `DOLLAR` and `FIXED`, error texts, the words and numbers in criteria and the order of text:

```js
import { compute } from 'xlcalc'

const cells = { A1: { formula: 'VALUE("1,5")+VALUE("1 234,5")' }, A2: { formula: '"a"&1.5&TRUE' }, A3: { formula: 'TEXT(45688.75,"ДД ММММ ГГГГ ч:мм")' }, A4: { formula: 'DOLLAR(-1234.567)' }, A5: { formula: 'VALUE("1.5")' } }
const r = compute({ locale: 'ru-RU', sheets: [{ name: 'S', cells }] }, { clock: { now: '2026-10-05T12:00:00Z', seed: 1 } })
const v = r.sheets[0].cells
console.log(v.A1.value, v.A2.value, v.A3.value, v.A4.value, v.A5.value) // → 1236 a1,5ИСТИНА 31 январь 2025 18:00 -1 234,57 ₽ 46143
console.log(r.volatile.map((c) => c.cell).join(',')) // → A5
```

In ru-RU, as Excel 16.113 for Mac with the Russian interface and locale computes: numbers in text take a decimal comma and a space between groups (`1 234,5`) and the ruble after the amount (`1 234,5 ₽`); dates are day-month-year (`31.01.2025`, `31/1/25`, `2025-01-31`, `31 янв 2025` with a month named by three letters or more of its nominative name; English month names are not read); `ИСТИНА` and `ЛОЖЬ` are the booleans in text (`"TRUE"` is no boolean there); numbers become text with a decimal comma (`a1,5`), complex numbers too; `TEXT` reads the Russian codes (`ДД.ММ.ГГГГ`, `ч:мм:сс`, `# ##0,00`, `Основной`) and treats English letters as literals; `DOLLAR` is `1 234,57 ₽`; text sorts Cyrillic before Latin. Date text without a year (`1.5` is 1 May, `15 янв`) takes the year of the book's clock, as Excel takes the current year, so a cell that reads one depends on the clock as `TODAY` does: it is in `volatile` and a session recomputes it on `setClock`. The rules were measured case by case in Excel ([docs/locale.md](docs/locale.md)).

### Structure operations

`restructure(book, op)` inserts or deletes rows and columns, moves (cut and paste) or copies (copy and paste, fill) a range, renames or deletes a sheet, and adjusts every reference in formulas and names on every sheet as Excel adjusts it: ranges grow and shrink, lost references become `#REF!`, absolute references move with inserted rows. It returns the new book and `mapCell`, which says where a cell of the old book is now; the input book is not changed.

```js
import { restructure } from 'xlcalc'

const book = {
  sheets: [{
    name: 'Sales',
    cells: {
      A1: { value: 120 }, A2: { value: 80 }, A3: { value: 45 },
      C1: { formula: 'SUM(A1:A3)' }, C2: { formula: 'A3*2' },
    },
  }],
}

const ins = restructure(book, { op: 'insertRows', sheet: 'Sales', row: 2 })
console.log(ins.book.sheets[0].cells.C1.formula, ins.book.sheets[0].cells.C3.formula) // → SUM(A1:A4) A4*2
console.log(ins.mapCell({ sheet: 'Sales', cell: 'A3' })) // → { sheet: 'Sales', cell: 'A4' }

const del = restructure(book, { op: 'deleteRows', sheet: 'Sales', row: 3 })
console.log(del.book.sheets[0].cells.C1.formula, del.book.sheets[0].cells.C2.formula) // → SUM(A1:A2) #REF!*2
const last = restructure(book, { op: 'deleteSheet', sheet: 'Sales' })
console.log(last.ok, last.code) // → false last_sheet
```

The operations are `insertRows` and `deleteRows` (`row`, a row number, and `count`), `insertColumns` and `deleteColumns` (`column`, letters, and `count`), `moveRange` (`from: { sheet, range }`, `to: { sheet, cell }`), `copyRange` (`from` and `to` as `{ sheet, range }`; a target that is a whole multiple of the source is tiled, as in fill down; copies have no `meta` unless `withMeta: true`), `renameSheet` (`name`, by Excel's rules for sheet names) and `deleteSheet`. Only the references that change are rewritten, in place; a formula that does not parse moves with its cell as written. An operation Excel refuses returns `{ ok: false, code, message }` with `off_sheet` (cells would leave the sheet), `sheet_name` (a name Excel would not take, or one another sheet has), `last_sheet` or `merged` (a move cutting a merged area, a copy landing on part of one). A name whose reference is lost gives `#REF!` where it is used, as in Excel; the parser reads the lost references Excel writes (`'S'!#REF!`, `#REF!A1`) as `#REF!`.

A book keeps styles as xlsx does: `styles`, a table of cell styles (font, solid fill, borders, alignment), `style` in a cell, an index into it, and per sheet `columns` and `rows` (by line `C`, `5` or span `B:D`, `5:9`: a width or height, a style, a number format) and `merges` (`B2:D3`). A cell without its own style or format shows its row's, else its column's. Computation ignores them, but a merged area stops an array spilling over it, as in Excel; the operations change them as Excel does: inserted rows take the format of the row above (its height, style and format, its cells' styles and formats) and columns of the column on the left, merged areas grow, shrink and move like ranges, moved and copied cells keep their look, whole rows and columns take their sizes and styles along. The rules, and how they were checked against Excel on 194 cases (64 with styles, 12 with spill references), are in [docs/structure.md](docs/structure.md).

## API

From `xlcalc`:

| export | |
|---|---|
| `compute(book, opts?)` | computes the whole book; `opts.clock`, `opts.externals`, `opts.execution` (`'closures'`, the default, or `'eval'`). Returns `ComputeResult`: `sheets` (values, canonical formulas, `format`, `meta`), `issues`, `nameIssues`, `volatile`, `clock`, `hash`, `inputHash`, `configVersion`, `ms` |
| `openSession(book, opts?)` | a `Session`: `setCells(edits)`, `setNames(names)`, `setClock(clock)`, `restructure(op)`, `setStyles(styles)`, `setLayout(edits)` (its changes), `book()`, `explain(sheet, cell, { depth, maxNodes, expand })`, `result()`, `close()`, `clock`. Options: `clock`, or `live: true`, `externals` and `execution`; edits are compiled in the session's execution mode |
| `validateFormula(text, ctx?)` | `{ ok: true, canonical, functions, sheets, names, externals }` or `{ ok: false, code, message, at }`; `ctx`: `sheets`, `names`, `externalFunctions` (without them sheets and names are not checked) |
| `validateName(name)` | `{ ok: true }` or `{ ok: false, message }` |
| `toLocal(formula, locale)`, `toInvariant(text, locale)` | a stored formula as the locale shows it, typed text as a stored formula: `{ ok: true, formula }` or `{ ok: false, code, message, at }` ([Localization](#localization)) |
| `validateLocalFormula(text, locale, ctx?)` | `validateFormula` for text typed in the locale, positions in `text` |
| `enUS`, `ruRU` | the locales of the package (`Locale` objects) |
| `engine` | `compute`, `validateFormula`, `validateName`, `openSession` and `restructure` as one `FormulaEngine` object |
| `restructure(book, op)` | one structure operation (`StructureOp`): `{ ok: true, book, mapCell }` or `{ ok: false, code, message }`; [docs/structure.md](docs/structure.md) |
| `FUNCTION_NAMES`, `CONFIG_VERSION` | the whitelist; the configuration version that goes into every hash |
| `InputError` | thrown on input of the wrong shape |
| `debugGeneratedCode(book, opts?)` | the JavaScript that the `eval` mode generates, in evaluation order, for debugging; not part of the contract |

`xlcalc/contract` exports the types and two constants (`MAX_FORMULA_LENGTH`, `MAX_FORMULA_DEPTH`) without any engine code, for code that only builds books or reads results.

`xlcalc/unstable/internal` (the parser, the criteria index, engine state) and `xlcalc/unstable/synth` (a seeded synthetic book for tests and benchmarks) exist for benchmark harnesses and adapters of other engines. They are outside semver and change in any release.

## Functions

The whitelist, by category (`FUNCTION_NAMES` has it in code):

- Math and trigonometry: `ABS`, `ACOS`, `ACOSH`, `ACOT`, `ACOTH`, `AGGREGATE`, `ARABIC`, `ASIN`, `ASINH`, `ATAN`, `ATAN2`, `ATANH`, `BASE`, `CEILING`, `CEILING.MATH`, `CEILING.PRECISE`, `COMBIN`, `COMBINA`, `COS`, `COSH`, `COT`, `COTH`, `CSC`, `CSCH`, `DECIMAL`, `DEGREES`, `EVEN`, `EXP`, `FACT`, `FACTDOUBLE`, `FLOOR`, `FLOOR.MATH`, `FLOOR.PRECISE`, `GCD`, `INT`, `ISO.CEILING`, `LCM`, `LN`, `LOG`, `LOG10`, `MDETERM`, `MINVERSE`, `MMULT`, `MOD`, `MROUND`, `MULTINOMIAL`, `MUNIT`, `ODD`, `PERCENTOF`, `PI`, `POWER`, `PRODUCT`, `QUOTIENT`, `RADIANS`, `RAND`, `RANDARRAY`, `RANDBETWEEN`, `ROMAN`, `ROUND`, `ROUNDDOWN`, `ROUNDUP`, `SEC`, `SECH`, `SEQUENCE`, `SERIESSUM`, `SIGN`, `SIN`, `SINH`, `SQRT`, `SQRTPI`, `SUBTOTAL`, `SUM`, `SUMIF`, `SUMIFS`, `SUMPRODUCT`, `SUMSQ`, `SUMX2MY2`, `SUMX2PY2`, `SUMXMY2`, `TAN`, `TANH`, `TRUNC`

- Statistical and compatibility: `AVEDEV`, `AVERAGE`, `AVERAGEA`, `AVERAGEIF`, `AVERAGEIFS`, `BETA.DIST`, `BETA.INV`, `BETADIST`, `BETAINV`, `BINOM.DIST`, `BINOM.DIST.RANGE`, `BINOM.INV`, `BINOMDIST`, `CHIDIST`, `CHIINV`, `CHISQ.DIST`, `CHISQ.DIST.RT`, `CHISQ.INV`, `CHISQ.INV.RT`, `CHISQ.TEST`, `CHITEST`, `CONFIDENCE`, `CONFIDENCE.NORM`, `CONFIDENCE.T`, `CORREL`, `COUNT`, `COUNTA`, `COUNTBLANK`, `COUNTIF`, `COUNTIFS`, `COVAR`, `COVARIANCE.P`, `COVARIANCE.S`, `CRITBINOM`, `DEVSQ`, `EXPON.DIST`, `EXPONDIST`, `F.DIST`, `F.DIST.RT`, `F.INV`, `F.INV.RT`, `F.TEST`, `FDIST`, `FINV`, `FISHER`, `FISHERINV`, `FORECAST`, `FORECAST.LINEAR`, `FREQUENCY`, `FTEST`, `GAMMA`, `GAMMA.DIST`, `GAMMA.INV`, `GAMMADIST`, `GAMMAINV`, `GAMMALN`, `GAMMALN.PRECISE`, `GAUSS`, `GEOMEAN`, `GROWTH`, `HARMEAN`, `HYPGEOM.DIST`, `HYPGEOMDIST`, `INTERCEPT`, `KURT`, `LARGE`, `LINEST`, `LOGEST`, `LOGINV`, `LOGNORM.DIST`, `LOGNORM.INV`, `LOGNORMDIST`, `MAX`, `MAXA`, `MAXIFS`, `MEDIAN`, `MIN`, `MINA`, `MINIFS`, `MODE`, `MODE.MULT`, `MODE.SNGL`, `NEGBINOM.DIST`, `NEGBINOMDIST`, `NORM.DIST`, `NORM.INV`, `NORM.S.DIST`, `NORM.S.INV`, `NORMDIST`, `NORMINV`, `NORMSDIST`, `NORMSINV`, `PEARSON`, `PERCENTILE`, `PERCENTILE.EXC`, `PERCENTILE.INC`, `PERCENTRANK`, `PERCENTRANK.EXC`, `PERCENTRANK.INC`, `PERMUT`, `PERMUTATIONA`, `PHI`, `POISSON`, `POISSON.DIST`, `PROB`, `QUARTILE`, `QUARTILE.EXC`, `QUARTILE.INC`, `RANK`, `RANK.AVG`, `RANK.EQ`, `RSQ`, `SKEW`, `SKEW.P`, `SLOPE`, `SMALL`, `STANDARDIZE`, `STDEV`, `STDEV.P`, `STDEV.S`, `STDEVA`, `STDEVP`, `STDEVPA`, `STEYX`, `T.DIST`, `T.DIST.2T`, `T.DIST.RT`, `T.INV`, `T.INV.2T`, `T.TEST`, `TDIST`, `TINV`, `TREND`, `TRIMMEAN`, `TTEST`, `VAR`, `VAR.P`, `VAR.S`, `VARA`, `VARP`, `VARPA`, `WEIBULL`, `WEIBULL.DIST`, `Z.TEST`, `ZTEST`

- Logical: `AND`, `FALSE`, `IF`, `IFERROR`, `IFNA`, `IFS`, `LET`, `NOT`, `OR`, `SWITCH`, `TRUE`, `XOR`

- Information: `ERROR.TYPE`, `ISBLANK`, `ISERR`, `ISERROR`, `ISEVEN`, `ISLOGICAL`, `ISNA`, `ISNONTEXT`, `ISNUMBER`, `ISODD`, `ISTEXT`, `N`, `NA`, `TYPE`

- Text: `ARRAYTOTEXT`, `BAHTTEXT`, `CLEAN`, `CONCAT`, `CONCATENATE`, `DOLLAR`, `EXACT`, `FIND`, `FINDB`, `FIXED`, `LEFT`, `LEFTB`, `LEN`, `LENB`, `LOWER`, `MID`, `MIDB`, `NUMBERVALUE`, `PROPER`, `REPLACE`, `REPLACEB`, `REPT`, `RIGHT`, `RIGHTB`, `SEARCH`, `SEARCHB`, `SUBSTITUTE`, `T`, `TEXT`, `TEXTAFTER`, `TEXTBEFORE`, `TEXTJOIN`, `TEXTSPLIT`, `TRIM`, `UNICHAR`, `UNICODE`, `UPPER`, `VALUE`, `VALUETOTEXT`

- Date and time: `DATE`, `DATEDIF`, `DATEVALUE`, `DAY`, `DAYS`, `DAYS360`, `EDATE`, `EOMONTH`, `HOUR`, `ISOWEEKNUM`, `MINUTE`, `MONTH`, `NETWORKDAYS`, `NETWORKDAYS.INTL`, `NOW`, `SECOND`, `TIME`, `TIMEVALUE`, `TODAY`, `WEEKDAY`, `WEEKNUM`, `WORKDAY`, `WORKDAY.INTL`, `YEAR`, `YEARFRAC`

- Financial: `ACCRINT`, `ACCRINTM`, `AMORDEGRC`, `AMORLINC`, `COUPDAYBS`, `COUPDAYSNC`, `COUPNCD`, `COUPNUM`, `COUPPCD`, `CUMIPMT`, `CUMPRINC`, `DB`, `DDB`, `DISC`, `DOLLARDE`, `DOLLARFR`, `DURATION`, `EFFECT`, `FV`, `FVSCHEDULE`, `INTRATE`, `IPMT`, `IRR`, `ISPMT`, `MDURATION`, `MIRR`, `NOMINAL`, `NPER`, `NPV`, `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`, `PDURATION`, `PMT`, `PPMT`, `PRICE`, `PRICEDISC`, `PRICEMAT`, `PV`, `RATE`, `RECEIVED`, `RRI`, `SLN`, `SYD`, `TBILLEQ`, `TBILLPRICE`, `TBILLYIELD`, `XIRR`, `XNPV`, `YIELD`, `YIELDDISC`, `YIELDMAT`

- Lookup and reference: `ADDRESS`, `CHOOSE`, `CHOOSECOLS`, `CHOOSEROWS`, `COLUMN`, `COLUMNS`, `DROP`, `EXPAND`, `FILTER`, `HLOOKUP`, `HSTACK`, `HYPERLINK`, `INDEX`, `LOOKUP`, `MATCH`, `ROW`, `ROWS`, `SORT`, `SORTBY`, `TAKE`, `TOCOL`, `TOROW`, `TRANSPOSE`, `UNIQUE`, `VLOOKUP`, `VSTACK`, `WRAPCOLS`, `WRAPROWS`, `XLOOKUP`, `XMATCH`

- Database: `DAVERAGE`, `DCOUNT`, `DCOUNTA`, `DGET`, `DMAX`, `DMIN`, `DPRODUCT`, `DSTDEV`, `DSTDEVP`, `DSUM`, `DVAR`, `DVARP`

- Engineering: `BESSELI`, `BESSELJ`, `BESSELK`, `BESSELY`, `BIN2DEC`, `BIN2HEX`, `BIN2OCT`, `BITAND`, `BITLSHIFT`, `BITOR`, `BITRSHIFT`, `BITXOR`, `COMPLEX`, `CONVERT`, `DEC2BIN`, `DEC2HEX`, `DEC2OCT`, `DELTA`, `ERF`, `ERF.PRECISE`, `ERFC`, `ERFC.PRECISE`, `GESTEP`, `HEX2BIN`, `HEX2DEC`, `HEX2OCT`, `IMABS`, `IMAGINARY`, `IMARGUMENT`, `IMCONJUGATE`, `IMCOS`, `IMCOSH`, `IMCOT`, `IMCSC`, `IMCSCH`, `IMDIV`, `IMEXP`, `IMLN`, `IMLOG10`, `IMLOG2`, `IMPOWER`, `IMPRODUCT`, `IMREAL`, `IMSEC`, `IMSECH`, `IMSIN`, `IMSINH`, `IMSQRT`, `IMSUB`, `IMSUM`, `IMTAN`, `OCT2BIN`, `OCT2DEC`, `OCT2HEX`

Status: checked against Excel on 33,710 cases: in en-US the coverage measurement (5,275), the case sets (16,830), the probes (959) and the random check (4,000); in ru-RU the coverage measurement (5,275) and the locale case sets (1,371); all with Excel's answers stored in this repository; spilling on 88 more (`docs/spill.md`).

Every function in the list matches Excel en-US on all of its cases in the coverage measurement (`docs/coverage-formulajs.md`) and in the Excel case sets (`test/excel/cases/`), which also cover the functions formula.js lacks (`XLOOKUP`, `XMATCH`, `TEXTBEFORE`, `ADDRESS` and others), with two kinds of exceptions, both listed in `test/excel-cases.ts`: `XIRR` matches within 1e-6 relative (`SOFT`: Excel stops its iteration within 1e-8 of the root, the engine finds the root), and so do the t and F quantiles (`T.INV`, `T.INV.2T`, `TINV`, `F.INV`, `F.INV.RT`, `FINV`), where Excel loses up to 7 digits at 10^9 degrees of freedom and more, and a few cases differ knowingly (`KNOWN`: the engine's answer is pinned by the tests and the reason is in [Known differences](#limitations)); a difference the random check finds outside the functions (operators, the array layer) is pinned as open in `test/excel-open.ts` until it is fixed (none is open now). `TODAY`, `NOW`, `RAND` and `RANDBETWEEN` (volatile) have their own semantics and clock tests instead; `RANDARRAY` has Excel cases on what does not depend on the draw (sizes, bounds, whole numbers) and clock tests.

How the acceptance works: the measurement generated 5,275 cases for 402 Excel functions, typical calls (mostly the examples of Excel's function reference) and edge cases, and computed them in Excel 16.113 for Mac with the en-US locale and English UI, as dynamic array formulas the way Excel 365 saves them (no implicit intersection, an array answer spills); the answers are stored in `test/excel/coverage/excel-en.json`. `test/excel.test.ts` computes the cases of every whitelisted function with this engine, in one book with the measurement's data sheet, and compares: numbers within |a−b| ≤ 1e-9·max(1,|a|), text, booleans and error codes exactly. The test needs neither Excel nor LibreOffice. Rules that are not one function's (numbers as text, long number literals, criteria over whole columns) are checked the same way on Excel's answers to 959 probes in `test/excel/probes.json` (`test/excel-probes.test.ts`). Every case set and the probes are computed as dynamic array formulas too. The `core` set pins the rules of operators, arrays and references the random check pointed at, and `near-integer` how Excel reads a number near a whole number, on exact values from one unit in the last place to 10^-6 away. The random check (`fuzz`) generates formulas over the whitelist from a seed, nested calls, operators, ranges and arrays, number and date text and errors among the arguments, and compares them with Excel's answers the same way; the rules it found are pinned by `fuzz-rules`, and the differences it found outside the functions are listed in `test/excel-open.ts`. Arrays are checked by the `arrays` set: operators and constants, array results, aggregates and lookups over arrays, every whitelisted function's one-value arguments given a range and an array, and the functions with an array answer. The older cases of `test/semantics-cases.ts` (expected answers from Excel's documentation, cross-checked with LibreOffice) still run too.

Notes on semantics:

- Arguments are coerced by kind, the same way for every function of a group: numeric arguments read text as `VALUE` does (`"$1,000"`, `"50%"`, `"1 1/2"`, date and time text: `ABS("2025-01-31")` is 45688) and booleans, so do operators and `TEXT`, a range in `SUM`, `AVERAGE`, `MAX` and the like contributes only its numbers while a direct argument is coerced, date arguments read date text as Excel does in the book's locale (en-US: `2025-01-31`, `1/31/2025`, `31-Jan-2025`), and `"TRUE"`/`"FALSE"` text (`"ИСТИНА"`/`"ЛОЖЬ"` in ru-RU) is a boolean where a boolean is expected; `AND`, `OR` and `XOR` take it given directly too, and skip other text.
- Where several arguments are wrong, the error is the one Excel finds first: arguments given as values are read before references and arrays (`LARGE(A1:A9,"x")` is `#VALUE!` even with an error in A1:A9; `AND`, `CONCAT` and `TEXTJOIN` answer an error given as a value before one inside a range), except `GCD`, `LCM` and `MULTINOMIAL`, which read every argument in order, values inside ranges and arrays too, and check signs last; `XIRR` finds a boolean or other text among its dates before a blank date; the former Analysis ToolPak functions check as Excel reads: the discount and most bond functions read the basis first, the coupon functions check the dates, the frequency and the basis in turn, the radix functions read their places first, `DELTA` and `GESTEP` their second number, `COMPLEX` its suffix; paired statistics take a one-cell error first, then compare sizes (`#N/A`), then look for errors inside.
- `LINEST`, `LOGEST`, `TREND` and `GROWTH` leave out a variable that depends linearly on the others or on the constant, with coefficient and standard error 0, as Excel does (Excel takes variables by their largest norm left; one y value takes a row of x as its variables); sizes are checked before values (`#REF!` before `#VALUE!`), and an error among `TREND`'s or `GROWTH`'s new x is `#VALUE!`.
- The complex-number functions read a real number from text as `VALUE` does when no sign follows its first character (`IMABS("50%")` is 0.5, `IMREAL("1/2/2025")` is 45659, `IMREAL("2025-01-31")` is `#NUM!`), as Excel does.
- Criteria (`SUMIF`, `SUMIFS`, `COUNTIFS`) compare text case-insensitively but accent-sensitively, support `*` and `?` (where `~` takes the next character as it is; text without `*` or `?` is compared as it is, `~` included), read numbers, dates and times inside criterion text as `VALUE` does (`">=1/1/2026"`, `"$1,000"`, `"50%"`), and a number criterion's `=` also matches text that reads as that number (`2` matches `"2"`, `0.5` matches `"50%"`) while its `<>` matches all text, as in Excel. An error, or its text in any case (`"#N/A"`), matches the cells holding that error; `<` and `>` order errors as `ERROR.TYPE` does. Their ranges must have the same referenced size (else `#VALUE!`); whole columns of sheets with different numbers of rows fit, and blank cells of a whole column count where every criterion matches a blank (`COUNTIFS(A:A,"")`).
- Numbers become text (`&`, `CONCAT`, text arguments) as in Excel: rounded to 15 significant digits, in positional notation when that takes at most 20 characters, the sign aside (`0.0000000745`, `10000000000000000000`), else in scientific notation with an exponent of two digits at least (`7.4500104300146E-06`, `1E+20`). `TEXT` with the `General` format has Excel's other rule, at most 11 characters (`0.333333333`, `1.23457E-05`).
- Number literals and numeric text keep 15 significant digits and drop the rest, as Excel reads them (`123456789012345678` is 123456789012345000); a result smaller than the smallest normal number (2.2250738585072E-308) is 0.
- An operator converts its left operand before it looks at the right one: `"abc"-NA()` is `#VALUE!`, `NA()-"abc"` is `#N/A`; over arrays each element the same way, and a position outside a shorter array is `#N/A` whatever the other operand holds.
- Numbers compare as in Excel: `=`, `<` and the other comparison operators and criteria (`COUNTIF`, `SUMIFS` and the rest) take numbers equal to 15 significant digits as equal (`0.1+0.2=0.3` is TRUE), while lookups, `MATCH`, `XMATCH`, `UNIQUE`, `RANK`, `DELTA` and sorting compare exactly; `INT`, `TRUNC`, `ROUNDDOWN` and `ROUNDUP` read their number to 15 significant digits (`INT(1-2^-52)` is 1, and so is `INT(FTEST(A1:A8,A1:A8))`).
- A whole-number argument less than 2^-22 (about 2.4·10^-7) short of an integer is that integer (`LEFT("abc",0.9999999)` is `"a"`) in `LEFT`, `RIGHT`, `FIND`, `SEARCH`, `REPLACE`, `SUBSTITUTE`, `TEXTBEFORE`, `TEXTAFTER` (the instance), `FIXED`, `DOLLAR`, `ROUND` (the digits), `INDEX`, `ADDRESS`, `DATE`, `WEEKDAY` (the type), `FACT`, the column index of `VLOOKUP` and `HLOOKUP`, the field of the database functions, the function of `SUBTOTAL` and `AGGREGATE`, the tails and type of `T.TEST` and the tails of `TDIST`, the mode of `CEILING.MATH` and `FLOOR.MATH`, the month of `DB` and the format of `ARRAYTOTEXT` and `VALUETOTEXT`; the other whole-number arguments are truncated, as Excel truncates them (`MID`, `REPT`, `CHOOSE`, the digits of `ROUNDDOWN`, `ROUNDUP` and `TRUNC`, the counts of the statistical functions, `EDATE`'s months and the rest the `near-integer` set measured). `ISEVEN` and `ISODD` take a number less than 10^-10 short of an integer as that integer, `DATE` rounds a day below zero down, and `DAY`, `MONTH`, `YEAR` and `WEEKDAY` read a date and time to the nearest second, as `HOUR`, `MINUTE` and `SECOND` do (0.4 s before midnight is the next day).
- Argument counts are those Excel loads: `SORTBY` takes each further by-array with its order (2, 3, 5… arguments; Excel refuses a book with `SORTBY` of 4), `SWITCH` at most 126 pairs and a default; any other count is the check code `syntax`.
- `IF`, `IFERROR`, `IFNA` and `SWITCH` do not compute the branch not taken, unless their condition (value, expression or a value to match) is an array: then every argument is computed and the function works element by element.
- Arrays follow Excel 365's dynamic arrays. An operator, or a function argument that takes one value, given an array or a range of several cells works element by element: the result has the larger number of rows and of columns, a single row or column is repeated, a position outside a shorter array is `#N/A` (`{1,2,3}+{1,2}` is `{2,4,#N/A}`). The functions of the former Analysis ToolPak (`EOMONTH`, `BESSELJ`, the radix and complex-number functions, the bond functions and others) and the field of the database functions do not take a range there: `#VALUE!`, while an array (`EOMONTH(A1:A2+0,0)`) works; `N` and `T` take the top-left cell of a range. Functions that take ranges take arrays the same way (`SUM` adds an array's numbers), except that `MAXA`, `STDEVA` and their kin take only an array's numbers and `AVERAGEA` counts its text as 0. An array of one value stays an array under an operator or a lifted call (`TYPE({1}+0)` is 64, so `SUM({TRUE}=TRUE)` skips the boolean and is 0), while a one-cell reference is a value (`TYPE(A1:A1+0)` is 1). A whole column is an array of 1,048,576 rows (`SUM(B:B+1)` adds 1,048,576), computed on its occupied rows.
- `INDEX`, `TAKE`, `DROP`, `XLOOKUP` (its `if_not_found` too), `CHOOSE`, `IF`, `IFS` and `SWITCH` answer references when given references, one cell too, so `SUMIF(INDEX(A1:C9,0,1),">2")` works and `SUM` skips text there (`SUM(IF(TRUE,A2))` is 0 when A2 holds the text "2"); `IFERROR`, `IFNA`, `FILTER`'s `if_empty` and unary `+` give values. `HSTACK` and `VSTACK` answer an error given as a value (`VSTACK(1,1/0)` is `#DIV/0!`) but stack a reference to a cell holding an error; the other functions with an array answer give arrays, where a blank cell is an empty value (`COUNTA` counts it, `ISBLANK` is TRUE). `TYPE` of an array, or of a range of several cells, is 64.
- Sorting (`SORT`, `SORTBY`) is stable and orders numbers, text, booleans, errors, then blanks (blanks last in a descending sort too); text compares as `<` does.
- `SUBTOTAL` and `AGGREGATE` (options 0–3) skip the cells whose formula calls either of them anywhere (`=SUBTOTAL(9,A1:A2)*10` too, a plain reference to such a cell not), so nothing is counted twice; a book has no hidden rows, so codes 101–111 equal 1–11 and options 1, 3, 5, 7 equal 0, 2, 4, 6.
- Lookups compare within a type (numbers, text, booleans), text in any case, and never match blanks or errors in the data. An exact match in `VLOOKUP`, `HLOOKUP` and `MATCH` (and `XLOOKUP`, `XMATCH` in match mode 2) takes `*`, `?` and `~`, where `~` escapes any character. An approximate match is Excel's binary search over the referenced cells, so it gives Excel's answer on unsorted, mixed and whole-column data too (checked on about 200 such cases); `MATCH` with type −1 scans. A blank lookup value is 0 in `VLOOKUP`, `HLOOKUP`, `MATCH` and `LOOKUP` and finds a blank cell in `XLOOKUP` and `XMATCH`. An error in a column index or match type taken from a cell is `#REF!`, as in Excel. An array or a range of several cells as the column (row) index of `VLOOKUP` (`HLOOKUP`) answers the found row (column) at each number, shaped as the index; the first number in order that is not valid answers for all: `#VALUE!` below 1, `#REF!` past the table, for an error or for text that is not a number. A call lifted by an array key or match mode reads the index's top-left number in each call (`INDEX(VLOOKUP({5;8},D!L1:N6,{2;4},FALSE),2)` is `"eight"`). A range as `MATCH`'s type gives its top-left cell.
- Database functions (`DSUM` and the rest) read records by rows under a header row; a database without records is `#VALUE!`, and so is a field taken from a cell holding an error (an error given as the field is itself the answer). Criteria are written as in `SUMIFS`, except that text without an operator matches the values that begin with it (`App` selects `Apple`); criteria rows are alternatives, conditions in a row all hold. A condition under a header that is not a field name holds for every record when it is `TRUE` (Excel's computed criterion) and for none otherwise.
- Text longer than 32,767 characters: `&` and `CONCATENATE` cut it there, `CONCAT` and `TEXTJOIN` answer `#CALC!`, the other text functions `#VALUE!`, as Excel does. An empty array (`FILTER` that keeps nothing, `TAKE` of zero rows, `DROP` of everything) is `#CALC!` too.
- `VALUE` reads currency, thousands groups, parentheses, `%`, exponents, `1 1/2`, dates and times as Excel en-US does, times past 24 hours too (`25:00`); `NUMBERVALUE` removes whitespace and the group separators, then reads the rest the same way.
- `CHAR` and `CODE` are not in the whitelist: Excel's answer depends on the platform and its language (Windows-1252 in Excel for Windows en-US, Mac Roman in Excel for Mac with the English interface, Mac Cyrillic with the Russian one). `UNICHAR` and `UNICODE` work with code points and are in the whitelist.
- Dates follow Excel's 1900 system, phantom 1900-02-29 included: `DATE(1900,2,29)` is 60. `DATE` takes a day outside −32768…32767 as 32767 and answers `#NUM!` for a month of 32767 or more, as Excel does.

## Limitations

- **Spilling.** Of two arrays over one cell that do not depend on each other, the engine always lets the lower one spill; Excel's choice varies between runs (`docs/spill.md`). An array whose size depends on a volatile draw spills by the engine's clock (Excel can give `#SPILL!`). A cycle through a spill range is `#CYCLE!`. There are no `LET`, `LAMBDA` or tables.
- **Content Security Policy.** Only `execution: 'eval'` uses `new Function`: in a browser with a CSP that does not allow `'unsafe-eval'`, `compute` and `openSession` in that mode throw an `EvalError` saying so. The default `closures` mode generates no code and computes under such a policy.
- **Known differences from Excel:**
  - `XLOOKUP` and `XMATCH` with binary search (search mode 2 or −2) scan linearly: the same answer on sorted data, where Excel's binary search is defined. Match mode 3 (regular expressions) is `#VALUE!`.
  - `LOOKUP` with a result range shorter than the lookup range is `#REF!` past its end; Excel reads on past the end of the range.
  - Text order for `<`, `>` and sorting uses `Intl.Collator` with the book's locale (case-insensitive, accent-sensitive; Latin before Cyrillic in en-US, Cyrillic before Latin in ru-RU, as Excel sorts), which may differ from Excel for some characters; without `Intl` it falls back to comparing lower-cased code units.
  - `TEXT` reads the format codes of en-US and of ru-RU; exponent formats (`0.00E+00`, `0,00E+00`), colours, elapsed time and bracketed conditions (`[h]:mm`, `[ч]:мм`) give `#VALUE!`.
  - A second decimal separator among the decimals of a number format (`#,##0,00` in ru-RU) is a literal between digits; Excel shows one digit fewer before it in some numbers (`TEXT(-0.5,"#,##0,00")` is `-,50,00` in Excel and `-,500,00` here).
  - In en-US date text without a year (`1/31`, which Excel completes with the current year) is not read as a date; ru-RU reads it (`1.5`) with the clock's year.
  - `#NULL!` is read as `#ERROR!`.
  - `RATE` with a rate of exactly 0 as the answer gives 0 or a value within 1e-10 of it; Excel answers its own rounding noise up to about 2e-9 or, from some guesses, `#NUM!`.
  - `GAMMA.INV`, `CHISQ.INV`, `CHISQ.INV.RT` (and `GAMMAINV`, `CHIINV`) give the quantile where Excel's iteration fails with `#NUM!`: shapes from about 5·10^9 (10^10 degrees of freedom) and probabilities near 1 with shapes under 0.03; Excel's own `GAMMA.DIST` at the engine's answer gives the probability back.
  - `BINOM.INV` where alpha equals the probability of no success, (1 − p)^n, exactly: Excel decides by its own rounding of that power, the engine by exp(n·ln(1 − p)); they part on `BINOM.INV(3,0.25,0.421875)`, 0 in Excel and 1 here.
  - A number that Excel computes by another algorithm can differ from the engine's in the 15th significant digit, within the tests' relative tolerance of 1e-9; turned into text (`&`, `TEXTJOIN`) the last digit then differs (`TEXTJOIN(";",FALSE,BESSELK(10,3))` is `2.72527006358477E-05` in Excel and `2.72527006358476E-05` here).
  - `SORTBY` with an array or a range of several cells as a sort order is `#VALUE!`; Excel answers an array the size of the orders (the first sorted value, then blanks, or an error where a sort fails).
  - `LINEST` with a constant x variable besides the intercept: both give it coefficient 0 and the same fit, but Excel's statistics count it as a variable in some way the engine does not reproduce (`LINEST({1;2;3;5},{1,1,2;2,1,4;3,1,6;4,1,8},TRUE,TRUE)`: degrees of freedom 1 in Excel, 2 here, and the F statistic and the intercept's standard error differ).
  - `ROWS` and `COLUMNS` of an `IFS` that answers an error: Excel counts 1 (`ROWS(IFS("x",1,TRUE,2))`), the engine answers the error.
  - `FTEST` of a sample with itself is 1 − 3·2^-53 in Excel and 1 − 2^-52 here: the last bits of a number Excel computes by another algorithm can differ, which shows only when the difference is magnified (`(1-FTEST(A1:A8,A1:A8))*2^52` is 1.5 in Excel, 1 here).
  - A LET name that reads as a reference or a boolean (`A1`, `TRUE`): an xlsx file writes LET's names with the prefix `_xlpm.`, so Excel loads `LET(_xlpm.A1,1,_xlpm.A1)` (1); formula text has no prefix, there such a word is a reference or a boolean, and `LET(A1,1,A1)` is a syntax error here, as Excel refuses that text.
  - `VLOOKUP`, `HLOOKUP` and `MATCH` with an empty text as the table (`MATCH(5,"",0)`) are `#VALUE!` in Excel and `#N/A` here.
  - `XLOOKUP` with one value as the lookup array and a return array of several rows (`XLOOKUP(3,0,A1:A6)`) is `#N/A` in Excel and `#VALUE!` here (the sizes differ).
  - `WRAPROWS` and `WRAPCOLS` of one value with a count above 1 (`WRAPROWS(0,2)`) answer the value alone in Excel; here it is padded to the count.
  - `F.TEST`, `FTEST` and `MIRR` given an empty text as an array (`F.TEST("",A1:A8)`) are `#DIV/0!` in Excel (no numbers) and `#VALUE!` here.
  - `XIRR` and `XNPV` given one text or boolean as their values or dates (`XIRR("",A1:A5)`) are `#VALUE!` in Excel and `#NUM!` here (the sizes differ).
  - Serial date 0 in the bond functions: `ACCRINT` with an issue or first interest date of 0 and `PRICE` with a settlement date of 0 are `#NUM!` in Excel and computed here; `AMORDEGRC` with a purchase date of 0 is 113 in Excel and −47885 here.
  - `WEEKNUM` with an invalid return type is `#NUM!`, as Excel answers it alone; inside one generated book of the random check Excel 16.113 gave a number for `DAY(WEEKNUM(40977,1E+3)<>2.5)` (1, on two runs).
- **Syntax.** Stored formulas use English function names and separators (no `;` or decimal commas; text typed in another language goes through `toInvariant`), no R1C1 references, no references to other workbooks, no 3D references or structured table references. Functions outside the whitelist (`INDIRECT`, `OFFSET`, `LET`, …) give `#NAME?`.
- **Structure operations and styles.** Whole rows or columns copied onto part of a merged area are refused, where Excel pastes them. A book has no hidden rows or columns, outlines, conditional formats, validation or comments. Styles are kept and moved, not computed: `compute` does not return them.
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

The tests cover the whitelist against Excel's answers on the coverage measurement's cases (`test/excel.test.ts`) and on the probes (`test/excel-probes.test.ts`), the argument layer rule by rule, the semantics cases against Excel's documentation and the stored LibreOffice answers (`test/oracle/libreoffice.json`, refreshed by `npm run test:oracle`, which needs LibreOffice's `soffice` on `PATH`; the usual tests need no office suite); determinism (shuffled input, a child process, three time zones, cents over 3,400 amounts); references, errors, volatile functions, parsing and the canonical form; sessions (200 random edits against full computations) and `explain`; structure operations against Excel's answers in `test/excel/struct/`, on the synthetic book and on seeded books; budgets; SHA-256 against test vectors and `node:crypto`; a browser bundle run in an empty JavaScript context and with code generation disabled; the two execution modes against each other on every case book and over 200 session edits, the closures mode with code generation disabled; the published package (exports, tarball contents, a consumer importing it by name). The code blocks of this README run as a test too.

Excel's answers that the tests compare with are stored in `test/excel/`, so the tests need no Excel: the coverage measurement's cases and Excel columns in `coverage/`, the case sets in `cases/` (`cases/ru-RU/` for the sets computed in ru-RU), the probes in `probes.json`, structure operations in `struct/`, spilling in `spill/`, Excel's ru-RU formula text in `locale/`. They are recomputed only with Excel for Mac, by a step of `npx tsx bench/coverage/run.ts`: `excel en` and `excel ru` for the measurement, `probes` for the probes, `probes <set>` and `probes <set> ru` for a case set, `struct <set>`, `spill <set>`; `npx tsx bench/locale/snapshot.ts ru` takes Excel's ru-RU formula text. A run keeps the stored answers of unchanged cases and computes the new and changed ones. `bench/coverage/README.md` has the layout, the steps and the rules for driving Excel.

Functions live in `src/functions/<category>.ts` (`math`, `statistical`, `logical`, `information`, `text`, `date`, `financial`, `lookup`, `engineering`, `database`; old names such as `NORMDIST` sit next to their modern versions), each file with its own registry, a large category assembled from `src/functions/<category>/`; `src/functions.ts` assembles them. A function declares the kinds of its arguments (`args: 'n i?'`, see `src/args.ts`), and the shared argument layer coerces them, so an implementation gets numbers, text, lists or grids and never repeats the coercion rules. Ported code keeps its attribution in the file header. A function joins its registry when `test/excel.test.ts` passes for it; the coverage measurement (`bench/coverage`, `docs/coverage-formulajs.md`) lists the cases and the classes of the remaining formula.js functions.

## License

MIT. Function implementations are ported from [formula.js](https://github.com/formulajs/formulajs) (MIT, Copyright (c) 2014 Sutoiku, Inc.), [jStat](https://github.com/jstat/jstat) (MIT, Copyright (c) 2013 jStat) and [bessel](https://github.com/SheetJS/bessel) (Apache-2.0, Copyright (C) 2013-present SheetJS); [THIRD_PARTY.md](THIRD_PARTY.md) has their license texts and the files they are in.
