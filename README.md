# xlcalc

A formula engine for spreadsheet-like workbooks in JavaScript. You give it a book (sheets of cells with values and Excel formulas, plus named ranges); it gives back the computed values, the errors with their reasons, and a deterministic hash of the values. Formulas use Excel syntax: English function names, commas between arguments, a decimal point; [Localization](#localization) shows them and reads them in another language, as Excel does.

What it does differently:

- **Compiles each formula once.** By default a formula becomes a tree of JavaScript closures, with no code generated, so a strict Content Security Policy does not stand in the way; with `execution: 'eval'` it becomes JavaScript text, and one `new Function` call compiles the book. A formula that exceeds the JS compiler's nesting limit falls back to closures, preserving evaluation order. Both modes give the same values and hashes, and the tests compute every case book both ways. Cells are evaluated over the reference graph in topological order. A session keeps the compiled book and recomputes only the cells that depend on an edit. On a synthetic book with 3,400 data rows and 3,907 formulas, mostly SUMIFS over whole columns, a cold compute takes about 0.3 s with the cells' shown text (0.28 s without) and an edit of a criterion cell under 10 ms (p95) in either mode on an Apple M5 Pro with Node 26; `test/budgets.test.ts` enforces 1.5 s and 100 ms.
- **Deterministic values and hashes.** The same book with the same clock gives the same values and the same `hash` (SHA-256 of a canonical text of all values plus the engine configuration version), regardless of input order, process or time zone. `inputHash` identifies the input.
- **A clock for volatile functions.** `TODAY` and `NOW` come from `clock.now` (UTC), `RAND` and `RANDBETWEEN` from `clock.seed` and the cell address. Without a clock the engine takes the current moment and a random seed, and `result.clock` reproduces the result.
- **A whitelist of functions, checked against Microsoft Excel.** A function is admitted only when it gives Excel's answer on every one of its test cases, typical calls and edge cases (text and booleans in numeric arguments, blanks, errors, date text, ranges with mixed content), computed in Excel 16 with the en-US locale; [Functions](#functions) has the list. Any other function is reported as `unknown_function` (`#NAME?`) rather than computed with different semantics.
- **Arrays inside a formula.** Operators and functions work element by element over ranges and arrays with Excel 365's rules (`SUM(B:B*C:C)`, `SUMPRODUCT((A1:A9="x")*B1:B9)`, `SUM(IF(A1:A9>0,A1:A9))`), array constants `{1,2;3,4}` are read, and functions with an array answer (`SORT`, `FILTER`, `UNIQUE`, `SEQUENCE`, `MMULT`, …) can be used inside a formula. A formula whose result is an array spills into the cells to its right and below, as in Excel 365 (`#SPILL!` where they are taken), and `E2#` refers to the range it filled (see [Arrays](#arrays)).
- **Cells shown as Excel shows them.** Each cell of a result carries, beside its value, the text Excel shows in it: its number format in the book's locale, General narrowed to the column's width, `#` across a column a number does not fit, the colour of the format's section (see [Showing values](#showing-values)). The value stays the number; `display: false` leaves the text out.
- **Errors are values.** An invalid formula never throws: its cell gets an Excel error, and `issues` says why, with a check code and the position in the formula.
- **xlsx in and out.** `xlcalc/xlsx` reads an xlsx file into a book, with the values Excel stored and a list of what the book cannot hold, and writes a book with its computed values back to xlsx. It has no dependencies and loads none of the engine's computation (see [Reading and writing xlsx](#reading-and-writing-xlsx)).
- **Computing without blocking.** `createInlineCalc` computes in the same thread in portions that give way to the event loop, `createCalcPool` from `xlcalc/worker` in Node worker threads or Web Workers, with soft and hard deadlines, cancellation, a thread's memory limit and limits on the size of a book; a session survives the loss of its thread (see [Computing without blocking](#computing-without-blocking)).
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

A formula is stored without the leading `=` in `formula`, separately from `value`, so a string from your data never turns into a formula. Dates are passed explicitly as `{ date: '2026-01-31' }` and become Excel serial numbers; `format` and `meta` pass through to the result untouched, and the result's `text` shows the value by its `format`.

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
//     { sheet: 'Plan', cell: 'A2', value: 300, text: '300' },
//     { sheet: 'Plan', cell: 'B1', value: 400, text: '400' },
//     { sheet: 'Plan', cell: 'B2', value: 480, text: '480' }
//   ]
console.log(session.explain('Plan', 'B2').text)
// → 'Plan'!B2 = 480  ← ROUND(B1*1.2,0)
//     'Plan'!B1 = 400  ← SUM(A1:A2)
//       'Plan'!A1:A2 (2 cells)
console.log(session.result().hash === session.result().hash) // → true
session.close()
```

`setNames` replaces the named ranges, `result()` returns the same shape as `compute`, `hash()` its values hash alone, `cells(sheet, range)` the cells of one rectangle in that shape (with each cell's own `style`), for a window over a sheet, and `close()` releases the book. A full `compute` of the edited book with `session.clock` gives the same hash as `session.result()`; the test suite checks this after each of 200 random edits.

A session also restructures its book and keeps its styles, so a host can hold the book in the session alone and save `book()` as a version (the session's own objects: do not modify them).

```js
import { openSession } from 'xlcalc'

const session = openSession({
  sheets: [{ name: 'Plan', cells: { A1: { value: 100 }, A2: { value: 250 }, A3: { formula: 'SUM(A1:A2)', style: 0 } } }],
  styles: [{ font: { bold: true } }],
})
const r = session.restructure({ op: 'deleteRows', sheet: 'Plan', row: 1 })
console.log(r.mapCell({ sheet: 'Plan', cell: 'A3' }), r.changes)
// → { sheet: 'Plan', cell: 'A2' } [ { sheet: 'Plan', cell: 'A2', value: 250, text: '250' } ]
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
console.log(out.C1, out.C2) // → { value: 8, text: '8', formula: 'SORT(A1:A3,1,-1)', spill: 'C1:C3' } { value: 5, text: '5', spillOf: 'C1' }
console.log(out.E1.value, r.issues[0].reason)
// → { error: 'SPILL' } the array of 3×1 values cannot spill: 'S'!E3 is not empty
```

A spilled cell is a value of the book: it is in `sheets[].cells` and in the hash. An array of more than 65,536 values, such as `A:A*2` (1,048,576 rows), lists its cells up to where the rest holds one value and gives that value as the formula's `spillFill` (0 below the data). The design is in `docs/arrays-incell.md`, spilling and the rules measured in Excel in `docs/spill.md`.

`@A1:A10` takes the cell of the range in the formula's row or column, and an array's top-left value. `SUM((A1:A3,C1:C3))` sums a union of ranges, and `SUM('Jan:Mar'!B2)` sums B2 on every sheet from Jan to Mar. A space between two ranges is their intersection (`SUM(A1:C3 B2:D4)` sums B2:C3, `#NULL!` where they do not meet), and `:` between references that are not one range joins them (`SUM(A1:INDEX(B1:B9,3))`). [docs/references.md](docs/references.md) has the rules.

An array formula of the old kind, entered with Ctrl+Shift+Enter, is a cell with `array`, its area:

```js
import { compute } from 'xlcalc'

const cells = { A1: { value: 1 }, A2: { value: 2 }, C1: { formula: 'A1:A2*10', array: 'C1:C3' } }
const out = compute({ sheets: [{ name: 'S', cells }] }).sheets[0].cells
console.log(out.C1.value, out.C2.value, out.C3.value) // → 10 20 { error: 'NA' }
```

The array fills its area and never spills. A session refuses to change part of the area, as Excel does. `fromLegacy` turns a formula of the old kind from a file into the dynamic formula Excel 365 shows for it, with `@` where Excel puts it, and `readXlsx` does so for every such formula of a file. [docs/legacy.md](docs/legacy.md) has both.

### Checking formulas

`validateFormula` checks a formula without a book: it returns the canonical text (sheet names quoted, function names upper-case) and what the formula uses, or a check code with a message and a position.

```js
import { validateFormula, validateName, validateNames } from 'xlcalc'

console.log(validateFormula("sum(Sales!b1:b3, 'Q1 data'!A1)", { sheets: ['Sales', 'Q1 data'] }).canonical)
// → SUM('Sales'!B1:B3,'Q1 data'!A1)
const bad = validateFormula('SUM(B1:B3', { sheets: ['Sales'] })
console.log(bad.ok, bad.code, bad.message, bad.at) // → false syntax expected ), found end of formula 9
console.log(validateName('Q1').ok, validateName('tax_rate').ok, validateName('Ставка').ok) // → false true true
console.log(validateName('ложь').ok, validateName('ложь', { locale: 'ru-RU' }).ok) // → true false
const names = [{ name: 'R', ref: '1' }, { name: 'rate', ref: "'Q1 data'!B1" }, { name: 'gone', ref: "'Other'!A1" }]
console.log(validateNames(names, { sheets: ['Sales', 'Q1 data'] }).map((x) => `${x.name} ${x.dropped}`)) // → [ 'R true', 'gone false' ]
```

`validateName` checks a name by Excel's rules for a book of `locale`, en-US without one (a ru-RU book also refuses `ИСТИНА` and `ЛОЖЬ`). `validateNames` gives the name issues a book with these sheets, tables and locale would give, without the book: `dropped` is true for a name the book leaves out (one Excel would not take, a duplicate, an unknown scope sheet, a formula that does not parse or a reference without its sheet), false for one it keeps that gives an error where it is used, so a host can refuse a name before it stores it.

Check codes: `empty`, `syntax`, `too_long` (over 8,192 characters), `too_deep` (over 64 levels of nesting), `external_ref` (a reference to another workbook), `unknown_function`, `unknown_sheet`, `unknown_name`, `unsupported` (an external function called with non-literal arguments). In `compute`, an invalid formula gives `#NAME?` for an unknown function or name, `#REF!` for an unknown sheet or an external reference, and `#ERROR!` otherwise. A formula whose result is an array is valid: it spills.

### Errors

Cell errors come out as `{ error: code }` with codes `DIV0`, `NA`, `NAME`, `NUM`, `REF`, `VALUE`, `CYCLE`, `SPILL`, `CALC` and `ERROR`. Every cell with an error is listed in `issues` with a reason, which names the cell an error came from (`division by zero, from 'S'!A2`). A circular reference gives `#CYCLE!` to the cells in the cycle and to everything that depends on them; the rest of the book is computed. A book with [iterative calculation](#iterative-calculation) computes its cycles instead.

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

A name follows Excel's rules for names, in letters of any script (`Ставка`, `НДС_2025`, `Став?ка`), and is compared in any case: `ставка` is the name `Ставка`, and a formula's text shows it as it is defined (`ставка*2` is `Ставка*2`). A name with `scope: 'Sheet'` is visible on that sheet only and shadows a global name. Invalid names are reported in `nameIssues`, `dropped` telling a name left out from one kept with an error.

A name can hold any formula: a constant (`0.2`), a calculation (`'Data'!$B$1*2`), a reference from `INDEX`, a union. A reference without `$` moves with the cell that uses the name. [docs/names.md](docs/names.md) has the rules.

### Tables

A sheet's `tables` are Excel's tables (list objects): a header row of column names, data rows and a totals row. Formulas refer to their parts by name, and the references follow a table as rows and columns are inserted and deleted in it.

```js
import { compute } from 'xlcalc'

const cells = {
  A1: { value: 'Item' }, B1: { value: 'Qty' }, C1: { value: 'Price' }, D1: { value: 'Total' },
  A2: { value: 'tea' }, B2: { value: 2 }, C2: { value: 10 }, D2: { formula: '[@Qty]*[@Price]' },
  A3: { value: 'coffee' }, B3: { value: 3 }, C3: { value: 20 }, D3: { formula: '[@Qty]*[@Price]' },
  F1: { formula: 'SUM(Orders[Total])' },
  F2: { formula: 'ROWS(Orders[#All])' },
  F3: { formula: 'Orders[[#Headers],[Price]]' },
}
const columns = [{ name: 'Item' }, { name: 'Qty' }, { name: 'Price' }, { name: 'Total', formula: '[@Qty]*[@Price]' }]
const r = compute({ sheets: [{ name: 'S', cells, tables: [{ name: 'Orders', range: 'A1:D3', columns }] }] })
const out = r.sheets[0].cells
console.log(out.F1.value, out.F2.value, out.F3.value) // → 80 3 Price
console.log(out.D2.formula) // → Orders[[#This Row],[Qty]]*Orders[[#This Row],[Price]]
```

A formula stores its references as an xlsx file does (`Orders[[#This Row],[Qty]]` for `[@Qty]`), and `toLocal` shows them as Excel does, in Russian too (`#Данные`). Structure operations grow, shrink and move tables as Excel does, an edit of a header cell renames its column, and `restructure` adds, renames, deletes and converts tables. [docs/tables.md](docs/tables.md) has the rules.

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

Books store formulas the way xlsx files do: English function names, `,` between arguments, `.` as the decimal separator. A locale changes how formulas are typed and shown, as Excel shows them in another language. The package has `ruRU` (Excel with the Russian interface) and `enUS`; a locale is a plain object (the `Locale` type), so formula text in another language is another object of the same shape. A book computes only in the package's locales; [docs/locale.md](docs/locale.md#adding-a-locale) says what adding one takes.

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

A book is computed in its own locale, `BookInput.locale` (`'en-US'` by default, or `'ru-RU'`), as Excel computes with those regional settings, whoever opens the book; a viewer still sees the formulas in their own language through `toLocal`. The locale changes how text reads as a number, a date or a boolean, how numbers and booleans become text, `TEXT`'s format codes and what cells show, `DOLLAR` and `FIXED`, error texts, the words and numbers in criteria and the order of text:

```js
import { compute } from 'xlcalc'

const cells = { A1: { formula: 'VALUE("1,5")+VALUE("1 234,5")' }, A2: { formula: '"a"&1.5&TRUE' }, A3: { formula: 'TEXT(45688.75,"ДД ММММ ГГГГ ч:мм")' }, A4: { formula: 'DOLLAR(-1234.567)' }, A5: { formula: 'VALUE("1.5")' } }
const r = compute({ locale: 'ru-RU', sheets: [{ name: 'S', cells }] }, { clock: { now: '2026-10-05T12:00:00Z', seed: 1 } })
const v = r.sheets[0].cells
console.log(v.A1.value, v.A2.value, v.A3.value, v.A4.value, v.A5.value) // → 1236 a1,5ИСТИНА 31 Январь 2025 18:00 -1 234,57 ₽ 46143
console.log(r.volatile.map((c) => c.cell).join(',')) // → A5
```

In ru-RU, as Excel 16.113 for Mac with the Russian interface and locale computes, and as Excel for Windows where the two part: numbers in text take a decimal comma and a space between groups (`1 234,5`) and the ruble after the amount (`1 234,5 ₽`); dates are day-month-year (`31.01.2025`, `31/1/25`, `2025-01-31`, `31 янв 2025` with a month named by three letters or more of its nominative name; English month names are not read); `ИСТИНА` and `ЛОЖЬ` are the booleans in text (`"TRUE"` is no boolean there); numbers become text with a decimal comma (`a1,5`), complex numbers too; `TEXT` reads the Russian codes (`ДД.ММ.ГГГГ`, `ч:мм:сс`, `# ##0,00`, `Основной`, `[Красный]`, `[Цвет10]`) and treats English letters as literals, while cells read their formats as xlsx files store them (`dd/mm/yyyy` shows `31.01.2025`); `DOLLAR` is `1 234,57 ₽`; month names have a capital (`Январь`) and short ones three letters (`фев`); text sorts Latin before Cyrillic, as Excel for Windows sorts (Excel for Mac puts Cyrillic first). Date text without a year (`1.5` is 1 May, `15 янв`) takes the year of the book's clock, as Excel takes the current year, so a cell that reads one depends on the clock as `TODAY` does: it is in `volatile` and a session recomputes it on `setClock`. The rules were measured case by case in Excel ([docs/locale.md](docs/locale.md)).

### Showing values

A result gives each cell, beside its `value`, the `text` Excel shows in it (VBA's `Range.Text`) and the `color` of its format's section. The format is the cell's own `format`, else its row's, else its column's: an Excel format code as an xlsx file stores it (`#,##0.00`, `d mmmm yyyy`, `[Blue]0;[Red]-0`), shown in the book's locale. The width is the column's (`columns: { A: { width: 16 } }`, as xlsx stores it; Excel's default without one), or a merged area's columns: General drops decimals or turns scientific to fit, a number that still does not fit shows `#` across the column, text is never cut (it runs over the next cells). The value stays the number whatever the text shows, and the text is not part of the values hash.

```js
import { compute, displayValue, openSession } from 'xlcalc'

const book = {
  sheets: [{
    name: 'S',
    cells: {
      A1: { value: 1234567.891, format: '#,##0.00' },
      A2: { value: -42, format: '[Blue]0;[Red]-0' },
      A3: { value: { date: '2025-01-31' }, format: 'd mmmm yyyy' },
      A4: { value: 1 / 3 },
      B1: { value: 123456789 },
    },
    columns: { A: { width: 16 }, B: { width: 6 } },
  }],
}
const cells = compute(book).sheets[0].cells
console.log(cells.A1, cells.A2.text, cells.A2.color, cells.A3.text, cells.A4.text, cells.B1.text)
// → { value: 1234567.891, text: '1,234,567.89', format: '#,##0.00' } -42 #FF0000 31 January 2025 0.333333333 1E+08
const ru = compute({ ...book, locale: 'ru-RU' }).sheets[0].cells
console.log(ru.A1.text, ru.A3.text, ru.A4.text) // → 1 234 567,89 31 Январь 2025 0,333333333
console.log(compute(book, { display: false }).sheets[0].cells.A1) // → { value: 1234567.891, format: '#,##0.00' }
const session = openSession(book)
console.log(session.setLayout([{ sheet: 'S', columns: 'A', set: { width: 8 } }]))
// → [
//     { sheet: 'S', cell: 'A1', value: 1234567.891, text: '########' },
//     { sheet: 'S', cell: 'A3', value: 45688, text: '########' },
//     { sheet: 'S', cell: 'A4', value: 0.333333333333333, text: '0.33333' }
//   ]
console.log(displayValue(0.4567, { format: '0.0%', locale: 'ru-RU' }), displayValue(1234.5, { width: 3 })) // → { text: '45,7%' } { text: '##' }
```

`compute(book, { display: false })` and `openSession(book, { display: false })` give values alone and spend nothing on formatting. A session's changes carry the text of the cells they list, and with the display on they also list the cells whose text or colour changed without a new value: a new `format` through `setCells`, a width, a row's or a column's format or a merged area through `setLayout`, a column moved by `restructure`. `displayValue(value, { format, locale, width })` shows a value without a book, for a cell being edited, say; the number itself is always the cell's `value`.

A format can also be the number of one of Excel's built-in formats, from 0 to 49, as an xlsx file stores a cell's format without a code. Excel shows some of them by the locale: `format: 14` is the short date, `1/31/2025` in en-US and `31.01.2025` in ru-RU, as Excel for Windows shows it. The table is in [docs/formats.md](docs/formats.md).

`TEXT` uses the same formats, read in the book's locale (`TEXT(x,"# ##0,00")` in ru-RU), without the width: it drops fills, has no colour and answers `#VALUE!` where a cell shows `#`. The rules, as measured in Excel 16.113 for Mac in en-US and ru-RU on 11,649 cells and 771 `TEXT` cases, are in [docs/formats.md](docs/formats.md), with the differences.

### Hidden rows and filters

A row or a column is hidden by `hidden: true` in `rows` or `columns`, as an xlsx file stores it, or by a height or width of 0, as Excel takes it. `SUBTOTAL` with codes 101 to 111 and `AGGREGATE` with options 1, 3, 5 and 7 leave out the cells of hidden rows. Hidden columns count, as in Excel.

A sheet whose autofilter has criteria is marked `filtered: true`. Excel takes every hidden row of such a sheet as one a filter hid, and `SUBTOTAL` with codes 1 to 11 leaves those out too.

```js
import { compute, openSession } from 'xlcalc'

const cells = { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 4 }, B1: { formula: 'SUBTOTAL(9,A1:A3)' }, B2: { formula: 'SUBTOTAL(109,A1:A3)' }, B3: { formula: 'AGGREGATE(9,5,A1:A3)' } }
const sheet = { name: 'S', cells, rows: { 2: { hidden: true } } }
const v = compute({ sheets: [sheet] }).sheets[0].cells
console.log(v.B1.value, v.B2.value, v.B3.value) // → 7 5 5
console.log(compute({ sheets: [{ ...sheet, filtered: true }] }).sheets[0].cells.B1.value) // → 5
const session = openSession({ sheets: [sheet] })
console.log(session.setLayout([{ sheet: 'S', rows: '2', set: { hidden: null } }]))
// → [
//     { sheet: 'S', cell: 'B3', value: 7, text: '7' },
//     { sheet: 'S', cell: 'B2', value: 7, text: '7' }
//   ]
```

A session hides and shows rows through `setLayout`, sets the filter mode with `{ sheet, filtered }`, and recomputes these formulas. Structure operations treat hidden rows as Excel does: on a filtered sheet they delete and copy the visible rows only ([docs/structure.md](docs/structure.md)).

### Date systems

A book is in Excel's 1900 date system unless it has `dateSystem: 1904`, as an xlsx file made on an old Mac may have. In the 1904 system serial 0 is 1 January 1904. Dates typed in, date text, `DATE`, `TODAY` and the date functions count from there, and cells and `TEXT` show a negative date or time with a minus.

```js
import { compute } from 'xlcalc'

const cells = { A1: { formula: 'DATE(2025,1,31)' }, A2: { formula: 'YEAR(A1)' }, A3: { value: -0.5, format: 'h:mm' } }
const c = compute({ dateSystem: 1904, sheets: [{ name: 'S', cells }] }).sheets[0].cells
console.log(c.A1.value, c.A2.value, c.A3.text) // → 44226 2025 -12:00
```

`displayValue` takes `dateSystem` too. A book in the 1904 system has its own values hash.

### Iterative calculation

A circular reference gives `#CYCLE!`, unless the book has `iteration`, as Excel's setting "Enable iterative calculation" stores it in a file. Then the cells of every cycle, and the cells that depend on them, are computed again and again from 0, as Excel computes a book it opens. The rounds stop after `maxIterations` (100 by default) or when no number changes by `maxChange` (0.001 by default) or more.

```js
import { compute } from 'xlcalc'

const cells = { B1: { value: 1000 }, B2: { value: 0.1 }, B3: { formula: 'B2*(B1+B4)/2' }, B4: { formula: 'B1+B3-100' } }
const r = compute({ sheets: [{ name: 'S', cells }], iteration: {} }).sheets[0].cells
console.log(r.B3.value, r.B4.value) // → 99.999984375 999.999984375
console.log(compute({ sheets: [{ name: 'S', cells }] }).sheets[0].cells.B3.value) // → { error: 'CYCLE' }
```

Cells are computed sheet by sheet in the order of the sheets' names, row by row. A change of text, a boolean or an error into another of them does not count as a change. Every computation starts from 0, so a session gives what a full computation gives. Excel goes on from the values it holds when it recalculates, which the engine does not do (see [Limitations](#limitations)).

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

The operations are `insertRows` and `deleteRows` (`row`, a row number, and `count`), `insertColumns` and `deleteColumns` (`column`, letters, and `count`), `moveRange` (`from: { sheet, range }`, `to: { sheet, cell }`), `copyRange` (`from` and `to` as `{ sheet, range }`; a target that is a whole multiple of the source is tiled, as in fill down; copies have no `meta` unless `withMeta: true`), `renameSheet` (`name`, by Excel's rules for sheet names), `deleteSheet`, `insertSheet` (`name`, `before`) and `moveSheet` (`sheet`, `before`), and over tables `addTable` (`sheet`, `table`), `renameTable` (`table`, `name`), `deleteTable` and `convertTable` (`table`; [docs/tables.md](docs/tables.md)). Only the references that change are rewritten, in place; a formula that does not parse moves with its cell as written. An operation Excel refuses returns `{ ok: false, code, message }` with `off_sheet` (cells would leave the sheet), `sheet_name` (a name Excel would not take, `History` among them, or one another sheet has), `last_sheet`, `merged` (a move cutting a merged area, a copy landing on part of one) `filtered` (a copy without the rows a filter hid, pasted over hidden rows of a filtered sheet) `array` (part of an old-kind array formula's area would change), `table` (rows or columns into two tables at once, or deleted from two tables unless both go whole, a header row deleted alone) or `table_name`. A name whose reference is lost gives `#REF!` where it is used, as in Excel; the parser reads the lost references Excel writes (`'S'!#REF!`, `#REF!A1`) as `#REF!`.

A book keeps styles as xlsx does: `styles`, a table of cell styles (font, solid fill, borders, alignment), `style` in a cell, an index into it, and per sheet `columns` and `rows` (by line `C`, `5` or span `B:D`, `5:9`: a width or height, a style, a number format) and `merges` (`B2:D3`). A cell without its own style or format shows its row's, else its column's. Computation ignores them, but a merged area stops an array spilling over it, as in Excel; the operations change them as Excel does: inserted rows take the format of the row above (its height, style and format, its cells' styles and formats) and columns of the column on the left, merged areas grow, shrink and move like ranges, moved and copied cells keep their look, whole rows and columns take their sizes and styles along. The rules, and how they were checked against Excel on 317 books, are in [docs/structure.md](docs/structure.md).

### Reading and writing xlsx

`readXlsx(data)` from `xlcalc/xlsx` reads a file's bytes into a book. Beside the book it gives `cached`, the values Excel stored for formulas and spilled cells, and `issues`, everything the book cannot hold, such as rich text or conditional formatting, with the sheet and cells. `writeXlsx(book, { result })` writes the book with its computed values. Formulas are written as dynamic array formulas with their spills, and `mode: 'values'` writes values only.

```js
import { compute } from 'xlcalc'
import { readXlsx, writeXlsx } from 'xlcalc/xlsx'

const book = { sheets: [{ name: 'Sales', cells: { A1: { value: 120 }, A2: { value: 80 }, B1: { formula: 'SUM(A1:A2)' }, C1: { formula: 'SEQUENCE(2)' } } }] }
const { data } = await writeXlsx(book, { result: compute(book) })

const read = await readXlsx(data)
console.log(read.book.sheets[0].cells.B1) // → { formula: 'SUM(A1:A2)' }
console.log(read.cached[0].cells) // → { B1: 200, C1: 1, C2: 2 }
console.log(read.issues) // → []

const values = await readXlsx((await writeXlsx(book, { result: compute(book), mode: 'values' })).data)
console.log(values.book.sheets[0].cells.C2) // → { value: 2 }
```

Both functions return promises: the archive is inflated and deflated by the platform's streams. Styles, number formats (built-in ones by their number), widths, heights, hidden rows and columns, merged areas, names, tables, a sheet's filter mode, the date system, iterative calculation and the book's locale go both ways. [docs/xlsx.md](docs/xlsx.md) lists what is read, written and reported, and how it was checked against Excel.

### Computing without blocking

`createInlineCalc()` computes in the same thread, in slices that let timers and input run between them; `createCalcPool()` from `xlcalc/worker` computes in a pool of threads, where a session's book lives in its thread. Both have one API: `compute`, `openSession` and the session's methods return promises, and every call takes `timeoutMs`, a soft deadline of computing time, and `signal`, an `AbortSignal`. A call that fails changes nothing. `limits` bound the formulas of a book, the values of one array and the cells of one spill, in the synchronous `compute` and `openSession` too; over them the call throws `CalcError` rather than computing.

```js
import { createInlineCalc } from 'xlcalc'
import { createCalcPool } from 'xlcalc/worker'

const book = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'A1*10' }, A3: { formula: 'SUM(SEQUENCE(1000,1000))' } } }] }

const calc = createInlineCalc({ limits: { cells: 100_000 } })
const e = await calc.compute(book).catch((e) => e)
console.log(e.code, e.limit, e.cell) // → limit cells { sheet: 'S', cell: 'A3' }

const pool = createCalcPool({ threads: 2, hardTimeoutMs: 10_000 })
const session = await pool.openSession(book, { clock: { now: '2026-10-07T12:00:00Z', seed: 1 } })
console.log(await session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }], { timeoutMs: 1_000 }))
// → [
//     { sheet: 'S', cell: 'A1', value: 3, text: '3' },
//     { sheet: 'S', cell: 'A2', value: 30, text: '30' }
//   ]
const stop = new AbortController()
const first = session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 4 } }])
const second = session.setCells([{ sheet: 'S', cell: 'A1', input: { value: 5 } }], { signal: stop.signal }).catch((e) => e)
stop.abort() // the second call waits for the first: it is refused at once
await first
console.log((await second).code, (await session.result()).sheets[0].cells.A2.value) // → aborted 40
await pool.close()
```

A task that runs past the pool's `hardTimeoutMs` terminates its thread, and in Node `maxHeapMb` bounds each thread's heap. The pool keeps a snapshot and a journal of each session's book, so a session whose thread was lost comes back on its next call as it was before the call that failed. [docs/async.md](docs/async.md) has the rules, the error codes and what this costs.

## API

From `xlcalc`:

| export | |
|---|---|
| `compute(book, opts?)` | computes the whole book; `opts.clock`, `opts.externals`, `opts.execution` (`'closures'`, the default, or `'eval'`), `opts.display` (`true`, the default, or `false`), `opts.limits` (`formulas`, `cells`, `spill`; over them `CalcError`). Returns `ComputeResult`: `sheets` (values, shown `text` and `color` with the display on, canonical formulas, `format`, `meta`), `issues`, `nameIssues`, `volatile`, `clock`, `hash`, `inputHash`, `configVersion`, `ms` |
| `openSession(book, opts?)` | a `Session`: `setCells(edits)`, `setNames(names)`, `setClock(clock)`, `restructure(op)`, `setStyles(styles)`, `setLayout(edits)` (its changes), `book()`, `explain(sheet, cell, { depth, maxNodes, expand })`, `cells(sheet, range)`, `result()`, `hash()` (the values hash without the result), `close()`, `clock`. Options: `clock`, or `live: true`, `externals`, `execution`, `display` and `limits`; edits are compiled in the session's execution mode |
| `displayValue(value, opts?)` | a value as a cell shows it, `{ text, color? }`; `opts`: `format`, `locale`, `width`, `dateSystem` ([Showing values](#showing-values)) |
| `validateFormula(text, ctx?)` | `{ ok: true, canonical, functions, sheets, names, externals }` or `{ ok: false, code, message, at }`; `ctx`: `sheets`, `names`, `externalFunctions` (without them sheets and names are not checked) |
| `validateName(name, opts?)` | `{ ok: true }` or `{ ok: false, message }` by the rules of a book of `opts.locale` |
| `validateNames(names, ctx)` | the name issues of a book with `ctx.sheets`, `ctx.tables`, `ctx.externalFunctions` and `ctx.locale`, as `compute` gives them, each with `dropped` |
| `toLocal(formula, locale, opts?)`, `toInvariant(text, locale)` | a stored formula as the locale shows it, typed text as a stored formula: `{ ok: true, formula }` or `{ ok: false, code, message, at }` ([Localization](#localization)); `opts.table`, the table of the formula's cell, shows its references as Excel does inside it (`[@Qty]`, [docs/tables.md](docs/tables.md)) |
| `validateLocalFormula(text, locale, ctx?)` | `validateFormula` for text typed in the locale, positions in `text` |
| `enUS`, `ruRU` | the locales of the package (`Locale` objects) |
| `engine` | `compute`, `validateFormula`, `validateName`, `validateNames`, `openSession`, `restructure` and `displayValue` as one `FormulaEngine` object |
| `restructure(book, op)` | one structure operation (`StructureOp`): `{ ok: true, book, mapCell }` or `{ ok: false, code, message }`; [docs/structure.md](docs/structure.md) |
| `copyFormula(text, from, to, book?)` | the text of a formula copied from cell `from` to `to` (`{ sheet, cell }`), as `copyRange` pastes it; with the book, a formula leaving its table names it ([docs/structure.md](docs/structure.md)) |
| `fromLegacy(formula, { names?, sheet? })` | a file's formula of the old kind as the dynamic formula Excel 365 shows for it, canonical text; [docs/legacy.md](docs/legacy.md) |
| `FUNCTION_NAMES`, `CONFIG_VERSION` | the whitelist; the configuration version that goes into every hash |
| `createInlineCalc(opts?)` | a `Calc` computing in the same thread in portions: `compute`, `openSession` (an `AsyncSession`) and `close`, as promises; `opts`: `sliceMs`, `timeoutMs`, `limits` ([docs/async.md](docs/async.md)) |
| `InputError` | thrown on input of the wrong shape |
| `CalcError` | an asynchronous call that failed, or a book over its limits; `code` says why ([docs/async.md](docs/async.md)) |
| `debugGeneratedCode(book, opts?)` | the JavaScript that the `eval` mode generates, in evaluation order, for debugging; not part of the contract |

`xlcalc/contract` exports the types and two constants (`MAX_FORMULA_LENGTH`, `MAX_FORMULA_DEPTH`) without any engine code, for code that only builds books or reads results.

`xlcalc/worker` exports `createCalcPool(opts?)`, a `Calc` over a pool of threads (`opts`: `threads`, `maxHeapMb`, `hardTimeoutMs`, `timeoutMs`, `maxQueue`, `maxSessions`, `maxSessionsPerThread`, `limits`, `snapshotEvery`, `sliceMs`, `spawn`), and `CalcError`.

`xlcalc/xlsx` exports `readXlsx(data)`, which resolves to `{ book, cached, issues }`, and `writeXlsx(book, opts?)`, which resolves to `{ data, issues }` (`opts`: `result`, `mode` `'formulas'` or `'values'`), with their types `XlsxRead`, `XlsxWritten`, `WriteXlsxOptions`, `XlsxIssue` and `XlsxIssueCode`.

`xlcalc/unstable/internal` (the parser, the criteria index, engine state) and `xlcalc/unstable/synth` (a seeded synthetic book for tests and benchmarks) exist for benchmark harnesses and adapters of other engines. They are outside semver and change in any release.

## Functions

The whitelist, by category (`FUNCTION_NAMES` has it in code):

- Math and trigonometry: `ABS`, `ACOS`, `ACOSH`, `ACOT`, `ACOTH`, `AGGREGATE`, `ARABIC`, `ASIN`, `ASINH`, `ATAN`, `ATAN2`, `ATANH`, `BASE`, `CEILING`, `CEILING.MATH`, `CEILING.PRECISE`, `COMBIN`, `COMBINA`, `COS`, `COSH`, `COT`, `COTH`, `CSC`, `CSCH`, `DECIMAL`, `DEGREES`, `EVEN`, `EXP`, `FACT`, `FACTDOUBLE`, `FLOOR`, `FLOOR.MATH`, `FLOOR.PRECISE`, `GCD`, `INT`, `ISO.CEILING`, `LCM`, `LN`, `LOG`, `LOG10`, `MDETERM`, `MINVERSE`, `MMULT`, `MOD`, `MROUND`, `MULTINOMIAL`, `MUNIT`, `ODD`, `PERCENTOF`, `PI`, `POWER`, `PRODUCT`, `QUOTIENT`, `RADIANS`, `RAND`, `RANDARRAY`, `RANDBETWEEN`, `ROMAN`, `ROUND`, `ROUNDDOWN`, `ROUNDUP`, `SEC`, `SECH`, `SEQUENCE`, `SERIESSUM`, `SIGN`, `SIN`, `SINH`, `SQRT`, `SQRTPI`, `SUBTOTAL`, `SUM`, `SUMIF`, `SUMIFS`, `SUMPRODUCT`, `SUMSQ`, `SUMX2MY2`, `SUMX2PY2`, `SUMXMY2`, `TAN`, `TANH`, `TRUNC`

- Statistical and compatibility: `AVEDEV`, `AVERAGE`, `AVERAGEA`, `AVERAGEIF`, `AVERAGEIFS`, `BETA.DIST`, `BETA.INV`, `BETADIST`, `BETAINV`, `BINOM.DIST`, `BINOM.DIST.RANGE`, `BINOM.INV`, `BINOMDIST`, `CHIDIST`, `CHIINV`, `CHISQ.DIST`, `CHISQ.DIST.RT`, `CHISQ.INV`, `CHISQ.INV.RT`, `CHISQ.TEST`, `CHITEST`, `CONFIDENCE`, `CONFIDENCE.NORM`, `CONFIDENCE.T`, `CORREL`, `COUNT`, `COUNTA`, `COUNTBLANK`, `COUNTIF`, `COUNTIFS`, `COVAR`, `COVARIANCE.P`, `COVARIANCE.S`, `CRITBINOM`, `DEVSQ`, `EXPON.DIST`, `EXPONDIST`, `F.DIST`, `F.DIST.RT`, `F.INV`, `F.INV.RT`, `F.TEST`, `FDIST`, `FINV`, `FISHER`, `FISHERINV`, `FORECAST`, `FORECAST.LINEAR`, `FREQUENCY`, `FTEST`, `GAMMA`, `GAMMA.DIST`, `GAMMA.INV`, `GAMMADIST`, `GAMMAINV`, `GAMMALN`, `GAMMALN.PRECISE`, `GAUSS`, `GEOMEAN`, `GROWTH`, `HARMEAN`, `HYPGEOM.DIST`, `HYPGEOMDIST`, `INTERCEPT`, `KURT`, `LARGE`, `LINEST`, `LOGEST`, `LOGINV`, `LOGNORM.DIST`, `LOGNORM.INV`, `LOGNORMDIST`, `MAX`, `MAXA`, `MAXIFS`, `MEDIAN`, `MIN`, `MINA`, `MINIFS`, `MODE`, `MODE.MULT`, `MODE.SNGL`, `NEGBINOM.DIST`, `NEGBINOMDIST`, `NORM.DIST`, `NORM.INV`, `NORM.S.DIST`, `NORM.S.INV`, `NORMDIST`, `NORMINV`, `NORMSDIST`, `NORMSINV`, `PEARSON`, `PERCENTILE`, `PERCENTILE.EXC`, `PERCENTILE.INC`, `PERCENTRANK`, `PERCENTRANK.EXC`, `PERCENTRANK.INC`, `PERMUT`, `PERMUTATIONA`, `PHI`, `POISSON`, `POISSON.DIST`, `PROB`, `QUARTILE`, `QUARTILE.EXC`, `QUARTILE.INC`, `RANK`, `RANK.AVG`, `RANK.EQ`, `RSQ`, `SKEW`, `SKEW.P`, `SLOPE`, `SMALL`, `STANDARDIZE`, `STDEV`, `STDEV.P`, `STDEV.S`, `STDEVA`, `STDEVP`, `STDEVPA`, `STEYX`, `T.DIST`, `T.DIST.2T`, `T.DIST.RT`, `T.INV`, `T.INV.2T`, `T.TEST`, `TDIST`, `TINV`, `TREND`, `TRIMMEAN`, `TTEST`, `VAR`, `VAR.P`, `VAR.S`, `VARA`, `VARP`, `VARPA`, `WEIBULL`, `WEIBULL.DIST`, `Z.TEST`, `ZTEST`

- Logical: `AND`, `FALSE`, `IF`, `IFERROR`, `IFNA`, `IFS`, `LET`, `NOT`, `OR`, `SWITCH`, `TRUE`, `XOR`

- Information: `ERROR.TYPE`, `ISBLANK`, `ISERR`, `ISERROR`, `ISEVEN`, `ISLOGICAL`, `ISNA`, `ISNONTEXT`, `ISNUMBER`, `ISODD`, `ISREF`, `ISTEXT`, `N`, `NA`, `TYPE`

- Text: `ARRAYTOTEXT`, `BAHTTEXT`, `CHAR`, `CLEAN`, `CODE`, `CONCAT`, `CONCATENATE`, `DOLLAR`, `EXACT`, `FIND`, `FINDB`, `FIXED`, `LEFT`, `LEFTB`, `LEN`, `LENB`, `LOWER`, `MID`, `MIDB`, `NUMBERVALUE`, `PROPER`, `REPLACE`, `REPLACEB`, `REPT`, `RIGHT`, `RIGHTB`, `SEARCH`, `SEARCHB`, `SUBSTITUTE`, `T`, `TEXT`, `TEXTAFTER`, `TEXTBEFORE`, `TEXTJOIN`, `TEXTSPLIT`, `TRIM`, `UNICHAR`, `UNICODE`, `UPPER`, `VALUE`, `VALUETOTEXT`

- Date and time: `DATE`, `DATEDIF`, `DATEVALUE`, `DAY`, `DAYS`, `DAYS360`, `EDATE`, `EOMONTH`, `HOUR`, `ISOWEEKNUM`, `MINUTE`, `MONTH`, `NETWORKDAYS`, `NETWORKDAYS.INTL`, `NOW`, `SECOND`, `TIME`, `TIMEVALUE`, `TODAY`, `WEEKDAY`, `WEEKNUM`, `WORKDAY`, `WORKDAY.INTL`, `YEAR`, `YEARFRAC`

- Financial: `ACCRINT`, `ACCRINTM`, `AMORDEGRC`, `AMORLINC`, `COUPDAYBS`, `COUPDAYSNC`, `COUPNCD`, `COUPNUM`, `COUPPCD`, `CUMIPMT`, `CUMPRINC`, `DB`, `DDB`, `DISC`, `DOLLARDE`, `DOLLARFR`, `DURATION`, `EFFECT`, `FV`, `FVSCHEDULE`, `INTRATE`, `IPMT`, `IRR`, `ISPMT`, `MDURATION`, `MIRR`, `NOMINAL`, `NPER`, `NPV`, `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE`, `ODDLYIELD`, `PDURATION`, `PMT`, `PPMT`, `PRICE`, `PRICEDISC`, `PRICEMAT`, `PV`, `RATE`, `RECEIVED`, `RRI`, `SLN`, `SYD`, `TBILLEQ`, `TBILLPRICE`, `TBILLYIELD`, `XIRR`, `XNPV`, `YIELD`, `YIELDDISC`, `YIELDMAT`

- Lookup and reference: `ADDRESS`, `AREAS`, `CHOOSE`, `CHOOSECOLS`, `CHOOSEROWS`, `COLUMN`, `COLUMNS`, `DROP`, `EXPAND`, `FILTER`, `HLOOKUP`, `HSTACK`, `HYPERLINK`, `INDEX`, `LOOKUP`, `MATCH`, `ROW`, `ROWS`, `SORT`, `SORTBY`, `TAKE`, `TOCOL`, `TOROW`, `TRANSPOSE`, `UNIQUE`, `VLOOKUP`, `VSTACK`, `WRAPCOLS`, `WRAPROWS`, `XLOOKUP`, `XMATCH`

- Database: `DAVERAGE`, `DCOUNT`, `DCOUNTA`, `DGET`, `DMAX`, `DMIN`, `DPRODUCT`, `DSTDEV`, `DSTDEVP`, `DSUM`, `DVAR`, `DVARP`

- Engineering: `BESSELI`, `BESSELJ`, `BESSELK`, `BESSELY`, `BIN2DEC`, `BIN2HEX`, `BIN2OCT`, `BITAND`, `BITLSHIFT`, `BITOR`, `BITRSHIFT`, `BITXOR`, `COMPLEX`, `CONVERT`, `DEC2BIN`, `DEC2HEX`, `DEC2OCT`, `DELTA`, `ERF`, `ERF.PRECISE`, `ERFC`, `ERFC.PRECISE`, `GESTEP`, `HEX2BIN`, `HEX2DEC`, `HEX2OCT`, `IMABS`, `IMAGINARY`, `IMARGUMENT`, `IMCONJUGATE`, `IMCOS`, `IMCOSH`, `IMCOT`, `IMCSC`, `IMCSCH`, `IMDIV`, `IMEXP`, `IMLN`, `IMLOG10`, `IMLOG2`, `IMPOWER`, `IMPRODUCT`, `IMREAL`, `IMSEC`, `IMSECH`, `IMSIN`, `IMSINH`, `IMSQRT`, `IMSUB`, `IMSUM`, `IMTAN`, `OCT2BIN`, `OCT2DEC`, `OCT2HEX`

Status: 42,115 cases are checked against Excel, with Excel's answers stored in this repository. In en-US they are the coverage measurement (5,275), the case sets (24,311), the probes (959) and the random check (4,000). In ru-RU they are the coverage measurement (5,275) and the case sets (2,295): the locale sets and six en-US sets run again in ru-RU. Spilling is checked on 104 more books (`docs/spill.md`), what cells show on 12,949 cells, 9,819 in en-US and 3,130 in ru-RU (`docs/formats.md`), structure operations on 317 books (`docs/structure.md`), tables on 171 books in en-US and ru-RU (`docs/tables.md`), and hidden rows, the 1904 date system and iterative calculation on 98 books (`bench/coverage/workbook/`).

Every function in the list matches Excel en-US on all of its cases in the coverage measurement (`docs/coverage-formulajs.md`) and in the Excel case sets (`test/excel/cases/`), which also cover the functions formula.js lacks (`XLOOKUP`, `XMATCH`, `TEXTBEFORE`, `ADDRESS` and others), with two kinds of exceptions, both listed in `test/excel-cases.ts`: `XIRR` matches within 1e-6 relative (`SOFT`: Excel stops its iteration within 1e-8 of the root, the engine finds the root), and so do the t and F quantiles (`T.INV`, `T.INV.2T`, `TINV`, `F.INV`, `F.INV.RT`, `FINV`), where Excel loses up to 7 digits at 10^9 degrees of freedom and more, and a few cases differ knowingly (`KNOWN`: the engine's answer is pinned by the tests and the reason is in [Known differences](#limitations)); a difference the random check finds outside the functions (operators, the array layer) is pinned as open in `test/excel-open.ts` until it is fixed (none is open now). `TODAY`, `NOW`, `RAND` and `RANDBETWEEN` (volatile) have their own semantics and clock tests instead; `RANDARRAY` has Excel cases on what does not depend on the draw (sizes, bounds, whole numbers) and clock tests.

How the acceptance works: the measurement generated 5,275 cases for 402 Excel functions, typical calls (mostly the examples of Excel's function reference) and edge cases, and computed them in Excel 16.113 for Mac with the en-US locale and English UI, as dynamic array formulas the way Excel 365 saves them (no implicit intersection, an array answer spills); the answers are stored in `test/excel/coverage/excel-en.json`. `test/excel.test.ts` computes the cases of every whitelisted function with this engine, in one book with the measurement's data sheet, and compares: numbers within |a−b| ≤ 1e-9·max(1,|a|), text, booleans and error codes exactly. The test needs neither Excel nor LibreOffice. Rules that are not one function's (numbers as text, long number literals, criteria over whole columns) are checked the same way on Excel's answers to 959 probes in `test/excel/probes.json` (`test/excel-probes.test.ts`). Every case set and the probes are computed as dynamic array formulas too. The `core` set pins the rules of operators, arrays and references the random check pointed at, and `near-integer` how Excel reads a number near a whole number, on exact values from one unit in the last place to 10^-6 away. The random check (`fuzz`) generates formulas over the whitelist from a seed, nested calls, operators, ranges and arrays, number and date text and errors among the arguments, and compares them with Excel's answers the same way; the rules it found are pinned by `fuzz-rules`, and the differences it found outside the functions are listed in `test/excel-open.ts`. Arrays are checked by the `arrays` set: operators and constants, array results, aggregates and lookups over arrays, every whitelisted function's one-value arguments given a range and an array, and the functions with an array answer. The older cases of `test/semantics-cases.ts` (expected answers from Excel's documentation, cross-checked with LibreOffice) still run too.

Notes on semantics:

- Arguments are coerced by kind, the same way for every function of a group: numeric arguments read text as `VALUE` does (`"$1,000"`, `"50%"`, `"1 1/2"`, date and time text: `ABS("2025-01-31")` is 45688) and booleans, so do operators and `TEXT`, a range in `SUM`, `AVERAGE`, `MAX` and the like contributes only its numbers while a direct argument is coerced, date arguments read date text as Excel does in the book's locale (en-US: `2025-01-31`, `1/31/2025`, `31-Jan-2025`; `1/31` and `Jan 31` in the clock's year), date and time text after a space is no date or time there (`DAY(" 1/31/2025")` is `#VALUE!`, while `VALUE`, `DATEVALUE`, `TIMEVALUE` and `DAYS` drop the spaces), and `"TRUE"`/`"FALSE"` text (`"ИСТИНА"`/`"ЛОЖЬ"` in ru-RU) is a boolean where a boolean is expected; `AND`, `OR` and `XOR` take it given directly too, and skip other text.
- Where several arguments are wrong, the error is the one Excel finds first: arguments given as values are read before references and arrays (`LARGE(A1:A9,"x")` is `#VALUE!` even with an error in A1:A9; `AND`, `CONCAT` and `TEXTJOIN` answer an error given as a value before one inside a range), except `GCD`, `LCM` and `MULTINOMIAL`, which read every argument in order, values inside ranges and arrays too, and check signs last; `XIRR` finds a boolean or other text among its dates before a blank date; the former Analysis ToolPak functions check as Excel reads: the discount and most bond functions read the basis first, the coupon functions check the dates, the frequency and the basis in turn, the radix functions read their places first, `DELTA` and `GESTEP` their second number, `COMPLEX` its suffix; paired statistics take a one-cell error first, then compare sizes (`#N/A`), then look for errors inside.
- `LINEST`, `LOGEST`, `TREND` and `GROWTH` leave out a variable that depends linearly on the others or on the constant, with coefficient and standard error 0, as Excel does (Excel takes variables by their largest norm left; one y value takes a row of x as its variables); sizes are checked before values (`#REF!` before `#VALUE!`), and an error among `TREND`'s or `GROWTH`'s new x is `#VALUE!`.
- The complex-number functions read each part of complex text as `VALUE` reads a number in the book's locale, the first one with its own sign (`IMREAL("50%+i")` is 0.5, `IMREAL("-(1)+i")` 1, `IMREAL("1 000+2i")` `#NUM!` in en-US), and a real number from text as `VALUE` does when no sign follows its first character (`IMABS("50%")` is 0.5, `IMREAL("1/2/2025")` is 45659, `IMREAL("2025-01-31")` is `#NUM!`), as Excel does.
- Criteria (`SUMIF`, `SUMIFS`, `COUNTIFS`) compare text case-insensitively but accent-sensitively, support `*` and `?` (where `~` takes the next character as it is; text without `*` or `?` is compared as it is, `~` included), read numbers, dates and times inside criterion text as `VALUE` does (`">=1/1/2026"`, `"$1,000"`, `"50%"`), and a number criterion's `=` also matches text that reads as that number (`2` matches `"2"`, `0.5` matches `"50%"`) while its `<>` matches all text, as in Excel. An error, or its text in any case (`"#N/A"`), matches the cells holding that error; `<` and `>` order errors as `ERROR.TYPE` does. Their ranges must have the same referenced size (else `#VALUE!`); whole columns of sheets with different numbers of rows fit, and blank cells of a whole column count where every criterion matches a blank (`COUNTIFS(A:A,"")`).
- Numbers become text (`&`, `CONCAT`, text arguments) as in Excel: rounded to 15 significant digits, in positional notation when that takes at most 20 characters, the sign aside (`0.0000000745`, `10000000000000000000`), else in scientific notation with an exponent of two digits at least (`7.4500104300146E-06`, `1E+20`). `TEXT` with the `General` format has Excel's other rule, at most 11 characters (`0.333333333`, `1.23457E-05`).
- Number literals and numeric text keep 15 significant digits and drop the rest, as Excel reads them (`123456789012345678` is 123456789012345000); a result smaller than the smallest normal number (2.2250738585072E-308) is 0.
- An operator converts its left operand before it looks at the right one: `"abc"-NA()` is `#VALUE!`, `NA()-"abc"` is `#N/A`; over arrays each element the same way, and a position outside a shorter array is `#N/A` whatever the other operand holds.
- Numbers compare as in Excel: `=`, `<` and the other comparison operators and criteria (`COUNTIF`, `SUMIFS` and the rest) take numbers equal to 15 significant digits as equal (`0.1+0.2=0.3` is TRUE), while lookups, `MATCH`, `XMATCH`, `UNIQUE`, `RANK`, `DELTA` and sorting compare exactly; `INT`, `TRUNC`, `ROUNDDOWN` and `ROUNDUP` read their number to 15 significant digits (`INT(1-2^-52)` is 1, and so is `INT(FTEST(A1:A8,A1:A8))`).
- A whole-number argument less than 2^-22 (about 2.4·10^-7) short of an integer is that integer (`LEFT("abc",0.9999999)` is `"a"`) in `LEFT`, `RIGHT`, `FIND`, `SEARCH`, `REPLACE`, `SUBSTITUTE`, `TEXTBEFORE`, `TEXTAFTER` (the instance), `FIXED`, `DOLLAR`, `ROUND` (the digits), `INDEX`, `ADDRESS`, `DATE`, `WEEKDAY` (the type), `FACT`, the column index of `VLOOKUP` and `HLOOKUP`, the field of the database functions, the function of `SUBTOTAL` and `AGGREGATE`, the tails and type of `T.TEST` and the tails of `TDIST`, the mode of `CEILING.MATH` and `FLOOR.MATH`, the month of `DB` and the format of `ARRAYTOTEXT` and `VALUETOTEXT`; the other whole-number arguments are truncated, as Excel truncates them (`MID`, `REPT`, `CHOOSE`, the digits of `ROUNDDOWN`, `ROUNDUP` and `TRUNC`, the counts of the statistical functions, `EDATE`'s months and the rest the `near-integer` set measured). `ISEVEN` and `ISODD` take a number less than 10^-10 short of an integer as that integer, `DATE` rounds a day below zero down, and `DAY`, `MONTH`, `YEAR` and `WEEKDAY` read a date and time to the nearest second, as `HOUR`, `MINUTE` and `SECOND` do (0.4 s before midnight is the next day).
- Argument counts are those Excel loads: `SORTBY` takes each further by-array with its order (2, 3, 5… arguments; Excel refuses a book with `SORTBY` of 4), `SWITCH` at most 126 pairs and a default, a function newer than Excel 2007 (`CONCAT`, `TEXTJOIN`, `STDEV.S`, `XOR`, `HSTACK`, `AGGREGATE` and the rest stored with the `_xlfn.` prefix) at most 254 arguments (253 for `SORTBY`, `MAXIFS` and `MINIFS`), the others 255; any other count is the check code `syntax`.
- `IF`, `IFERROR`, `IFNA` and `SWITCH` do not compute the branch not taken, unless their condition (value, expression or a value to match) is an array: then every argument is computed and the function works element by element.
- An argument left out between commas (`IF(A1,,1)`, `SUM(1,,2)`) is taken as Excel 16.113 takes it, measured on every whitelisted function's arguments (the `omit` sets): as a blank, read by the argument's kind as 0, `""` or FALSE (`ROUND(1.5,)` is 2, `LEFT("abc",)` is `""`, `VLOOKUP(4,A1:C6,2,)` an exact match, `IF(TRUE,)&"x"` is `"x"`, `COUNT(1,,2)` is 3); as `#VALUE!` where the function takes an array or a reference (`CORREL(,A1:A8)`, `INDEX(,1)`); as not given, with the function's default, where the function treats it so (the newer functions' optional arguments: `SEQUENCE(3,,5)`, `XLOOKUP(4,A1:A6,B1:B6,)` is `#N/A`; `FIXED`'s decimals, `ADDRESS`'s style, the significance of `CEILING.MATH` and its kin, `PRODUCT`'s values; `TEXTJOIN`'s left-out ignore_empty is TRUE); in the former Analysis ToolPak's functions as not given, a required one as `#N/A` (`EDATE(,1)`, `ERF(1,)`). A formula that leaves out a reference Excel requires (a criteria range, `RANK`'s list, `SUBTOTAL`'s references) is the check code `syntax`, as Excel does not load it. The table of the functions that differ from the blank is `src/functions/omitted.ts`.
- `LET` binds names to values, each seen from the end of its value to the end of its `LET` and in any case, an inner `LET` shadowing an outer one and both shadowing a defined name of the same word; a name used in its own value or past its `LET` is a defined name (`LET(x,x+1,x)` is `#NAME?` without one). A value is computed once, on its first use (`LET(x,RAND(),x-x)` is 0; an unused value's error does not count), and a reference stays a reference (`LET(r,A1:A9,SUM(r))` skips text, `SUMIF(LET(r,A1:A9,r),">0")` works); a left-out value is a blank. Excel loads a `LET` that ends with a value (an even number of arguments: `LET(x,1,y,2)` is 2) and at most 254 arguments, not a name declared twice or a word that is no name (`LET(A1,1,A1)`: the check code `syntax`). The canonical text spells every use as its name is declared (`LET(Abc,1,ABC+1)` is `LET(Abc,1,Abc+1)`), as Excel shows it; in ru-RU `LET` keeps its name.
- Arrays follow Excel 365's dynamic arrays. An operator, or a function argument that takes one value, given an array or a range of several cells works element by element: the result has the larger number of rows and of columns, a single row or column is repeated, a position outside a shorter array is `#N/A` (`{1,2,3}+{1,2}` is `{2,4,#N/A}`). The functions of the former Analysis ToolPak (`EOMONTH`, `BESSELJ`, the radix and complex-number functions, the bond functions and others) and the field of the database functions do not take a range there: `#VALUE!`, while an array (`EOMONTH(A1:A2+0,0)`) works; `N` and `T` take the top-left cell of a range. Functions that take ranges take arrays the same way (`SUM` adds an array's numbers), except that `MAXA`, `STDEVA` and their kin take only an array's numbers and `AVERAGEA` counts its text as 0. An array of one value stays an array under an operator or a lifted call (`TYPE({1}+0)` is 64, so `SUM({TRUE}=TRUE)` skips the boolean and is 0), while a one-cell reference is a value (`TYPE(A1:A1+0)` is 1). A whole column is an array of 1,048,576 rows (`SUM(B:B+1)` adds 1,048,576), computed on its occupied rows.
- `INDEX`, `TAKE`, `DROP`, `XLOOKUP` (its `if_not_found` too), `CHOOSE`, `IF`, `IFS` and `SWITCH` answer references when given references, one cell too, so `SUMIF(INDEX(A1:C9,0,1),">2")` works and `SUM` skips text there (`SUM(IF(TRUE,A2))` is 0 when A2 holds the text "2"); `IFERROR`, `IFNA`, `FILTER`'s `if_empty` and unary `+` give values. `HSTACK` and `VSTACK` answer an error given as a value (`VSTACK(1,1/0)` is `#DIV/0!`) but stack a reference to a cell holding an error; the other functions with an array answer give arrays, where a blank cell is an empty value (`COUNTA` counts it, `ISBLANK` is TRUE). `TYPE` of an array, or of a range of several cells, is 64. `IFS` answers the `#VALUE!` of a condition that is text as an array of that one error, so `ROWS(IFS("x",1,TRUE,2))` is 1, as in Excel, while an error given as a condition is the answer itself (`ROWS(IFS(1/0,1))` is `#DIV/0!`).
- Sorting (`SORT`, `SORTBY`) is stable and orders numbers, text, booleans, errors, then blanks (blanks last in a descending sort too); text compares as `<` does. `SORTBY` with an array or a range of several cells as a sort order answers as Excel does, an array as large as the orders, lifted as an operator lifts them (an order of one row or column repeats, past a shorter one `#N/A`): at each position, in rows, the error of the call with the orders there (an error given as an argument, `#VALUE!` for an order that is not 1 or −1 or has no value, or for a by-array of another size), the first value of the first call that sorts, and after it blanks, or `#VALUE!` when it has more than one by-array (`SORTBY(A1:A6,B1:B6,{1;-1})` is the first value and a blank, `SORTBY(A1:A6,B1:B6,{2;-1})` `#VALUE!` and the first value of the descending sort).
- `SUBTOTAL` and `AGGREGATE` (options 0–3) skip the cells whose formula calls either of them anywhere (`=SUBTOTAL(9,A1:A2)*10` too, a plain reference to such a cell not), so nothing is counted twice. `SUBTOTAL` 101–111 and `AGGREGATE` options 1, 3, 5 and 7 skip hidden rows, and `SUBTOTAL` 1–11 skips the hidden rows of a filtered sheet ([Hidden rows and filters](#hidden-rows-and-filters)). `AGGREGATE` 14–19 take an array as well as a reference (`AGGREGATE(14,6,10/A1:A9,1)`).
- Lookups compare within a type (numbers, text, booleans), text in any case, and never match blanks or errors in the data. An exact match in `VLOOKUP`, `HLOOKUP` and `MATCH` (and `XLOOKUP`, `XMATCH` in match mode 2) takes `*`, `?` and `~`, where `~` escapes any character. An approximate match is Excel's binary search over the referenced cells, so it gives Excel's answer on unsorted, mixed and whole-column data too (checked on about 200 such cases); `MATCH` with type −1 scans. A blank lookup value is 0 in `VLOOKUP`, `HLOOKUP`, `MATCH` and `LOOKUP` and finds a blank cell in `XLOOKUP` and `XMATCH`. `LOOKUP` reads on past the end of a result reference shorter than its lookup array, as Excel does, whatever answers the reference (a spill, `IF`, `INDEX`, `TAKE`): a column down, a row or a single cell to the right (`LOOKUP(4.19,L1:L6,B2:B3)` and `LOOKUP(4.19,L1:L6,IF(TRUE,B2:B3))` are B4 and `LOOKUP(4.19,L1:L6,B2)` D2 when 4.19 falls between L3 and L4), a blank past the sheet's end; the cells it may read count among the formula's references, past a spill as far as it spills. An error in a column index or match type taken from a cell is `#REF!`, as in Excel. An array or a range of several cells as the column (row) index of `VLOOKUP` (`HLOOKUP`) answers the found row (column) at each number, shaped as the index; the first number in order that is not valid answers for all: `#VALUE!` below 1, `#REF!` past the table, for an error or for text that is not a number. A call lifted by an array key or match mode reads the index's top-left number in each call (`INDEX(VLOOKUP({5;8},D!L1:N6,{2;4},FALSE),2)` is `"eight"`). A range as `MATCH`'s type gives its top-left cell. A table of `VLOOKUP`, `HLOOKUP` or `MATCH` given as one value is read as a number and finds nothing (`MATCH(5,5,0)` is `#N/A`, `MATCH(5,"x",0)` `#VALUE!`); one value as `XLOOKUP`'s lookup array looks along the return array's one row, else its one column (`XLOOKUP(0,0,A1:A6)` is A1:A6).
- Database functions (`DSUM` and the rest) read records by rows under a header row; a database without records is `#VALUE!`, and so is a field taken from a cell holding an error (an error given as the field is itself the answer). Criteria are written as in `SUMIFS`, except that text without an operator matches the values that begin with it (`App` selects `Apple`); criteria rows are alternatives, conditions in a row all hold. A condition under a header that is not a field name holds for every record when it is `TRUE` (Excel's computed criterion) and for none otherwise.
- Text longer than 32,767 characters: `&` and `CONCATENATE` cut it there, `CONCAT` and `TEXTJOIN` answer `#CALC!`, the other text functions `#VALUE!`, as Excel does. An empty array (`FILTER` that keeps nothing, `TAKE` of zero rows, `DROP` of everything) is `#CALC!` too.
- `VALUE` reads currency, thousands groups, parentheses, `%`, exponents, `1 1/2`, dates and times as Excel en-US does, times past 24 hours too (`25:00`), date text without a year in the clock's year (`1/31`; `2/29` in a year that is not leap is February 2029, as in Excel); time text follows one rule in `VALUE`, operators and date arguments (`6 PM` and `6 P` but not `6PM`, `6:30.5` as minutes and seconds); `NUMBERVALUE` removes whitespace and the group separators, then reads the rest the same way.
- `CHAR` and `CODE` work with the Windows code page of the book's locale, Windows-1252 in en-US and Windows-1251 in ru-RU, so a book gives the same characters as in Excel for Windows, where most books are made; Excel for Mac answers codes 128–255 by Mac Roman and Mac Cyrillic, a platform's own encoding rather than the book's. Codes 1–127, a code outside 1–255 (`#VALUE!`), the code of a character outside the page (95, `_`) and the arguments are as Excel computes them (the `char-code` set, in both locales); the tables are checked against Unicode's published ones. `UNICHAR` and `UNICODE` work with code points.
- `ACCRINT` counts each whole coupon period before the settlement as 1 and a part of one by its days over the period's length; on basis 1 that length is the period ending at the first interest date for every later period, and a period ending on the settlement counts as a part (`ACCRINT(DATE(2024,1,1),DATE(2024,7,1),DATE(2025,1,1),0.05,1000,2,1)` is 25·(1 + 184/182)). A coupon period that begins before 1900 leaves no coupons: `ACCRINT` is 0 when the issue's period does, `PRICE` `#NUM!` when the settlement's does, as `COUPNUM` finds none. Serial 0 as `ACCRINT`'s issue or first interest date and any date of 0 in `XNPV` are `#NUM!`, while `XIRR` takes it. `AMORDEGRC` and `AMORLINC` take no more than cost less salvage in the first period, and `AMORDEGRC`'s period that passes the salvage takes its whole amount, the later ones nothing.
- `WORKDAY`, `WORKDAY.INTL`, `NETWORKDAYS` and `NETWORKDAYS.INTL` read date and numeric text among their holidays, in cells and arrays too (`WORKDAY(A1,5,H1:H3)` with `"1/6/2025"` in H1); other text and booleans there are `#VALUE!`.
- Dates follow Excel's 1900 system, phantom 1900-02-29 included: `DATE(1900,2,29)` is 60. A book can be in the 1904 system instead ([Date systems](#date-systems)), where the former Analysis ToolPak's functions read date text given as a value as Excel does there: `EDATE("1/31/2025",1)` is 42793, while `EDATE(A1,1)` with that text in A1 is 44254. `DATE` takes a day outside −32768…32767 as 32767 and answers `#NUM!` for a month of 32767 or more, as Excel does.

## Limitations

- **Spilling.** Of two arrays over one cell that do not depend on each other, the engine always lets the lower one spill; Excel's choice varies between runs (`docs/spill.md`). An array whose size depends on a volatile draw spills by the engine's clock, where Excel gives `#SPILL!` when the size changes between its calculation passes (for most of ten such arrays in one book, for none of one alone). A cycle through a spill range is `#CYCLE!`. There are no `LAMBDA` (nor `MAP`, `REDUCE` and the other functions over it).
- **Content Security Policy.** Only `execution: 'eval'` uses `new Function`: in a browser with a CSP that does not allow `'unsafe-eval'`, `compute` and `openSession` in that mode throw an `EvalError` saying so. The default `closures` mode generates no code and computes under such a policy.
- **Known differences from Excel:**
  - `XLOOKUP` and `XMATCH` with binary search (search mode 2 or −2) scan linearly: the same answer on sorted data, where Excel's binary search is defined. Match mode 3 (regular expressions) is `#VALUE!`.
  - Text order for `<`, `>` and sorting uses `Intl.Collator` with the en-US order in both locales (case-insensitive, accent-sensitive, Latin before Cyrillic, as Excel for Windows sorts), which may differ from Excel for some characters; without `Intl` it falls back to comparing lower-cased code units.
  - Excel for Mac puts Cyrillic before Latin in ru-RU (`"a"<"а"` is FALSE there); the engine puts Latin first, as Excel for Windows does in every locale, in comparisons, criteria, lookups and SORT.
  - Excel for Mac writes month names in ru-RU in lower case (`TEXT` with `ММММ`, `МММММ` or `[$-419]mmmm`, and cells: `январь`, `я`); the engine writes them with a capital, as Excel for Windows does (`Январь`, `Я`).
  - Excel for Mac writes the short month names of ru-RU (`МММ`) `февр`, `март`, `июнь`, `июль`, `сент`, `нояб`; the engine takes three letters, as Excel for Windows does (`фев`, `мар`, `июн`, `июл`, `сен`, `ноя`).
  - A second decimal separator among the decimals of a number format (`#,##0,00` in ru-RU) is a literal between digits; Excel shows one digit fewer before it in some numbers (`TEXT(-0.5,"#,##0,00")` is `-,50,00` in Excel and `-,500,00` here).
  - `CHAR` and `CODE` take codes 128–255 from the Windows code page of the book's locale (Windows-1252 in en-US, Windows-1251 in ru-RU), as Excel for Windows does; Excel for Mac takes Mac Roman and Mac Cyrillic (`CHAR(200)` is `È` here and `»` in Excel for Mac en-US, `CODE("Ж")` in ru-RU 198 here and 134 there).
  - `CODE` of a character outside the code page of the book's locale is 63 (`?`), as Excel for Windows answers; Excel for Mac answers 95 (`_`): `CODE("中")`, and `CODE("Ж")` in en-US.
  - `RATE` with a rate of exactly 0 as the answer gives 0 or a value within 1e-10 of it; Excel answers its own rounding noise up to about 2e-9 or, from some guesses, `#NUM!`.
  - `GAMMA.INV`, `CHISQ.INV`, `CHISQ.INV.RT` (and `GAMMAINV`, `CHIINV`) give the quantile where Excel's iteration fails with `#NUM!`: shapes from about 5·10^9 (10^10 degrees of freedom) and probabilities near 1 with shapes under 0.03; Excel's own `GAMMA.DIST` at the engine's answer gives the probability back.
  - `BINOM.INV` where alpha equals the probability of no success, (1 − p)^n, exactly: Excel decides by its own rounding of that power, the engine by exp(n·ln(1 − p)); they part on `BINOM.INV(3,0.25,0.421875)`, 0 in Excel and 1 here.
  - A number that Excel computes by another algorithm can differ from the engine's in the 15th significant digit, within the tests' relative tolerance of 1e-9; turned into text (`&`, `TEXTJOIN`) the last digit then differs (`TEXTJOIN(";",FALSE,BESSELK(10,3))` is `2.72527006358477E-05` in Excel and `2.72527006358476E-05` here).
  - `LINEST` with a constant x variable besides the intercept: both give it coefficient 0 and the same fit, but Excel's statistics count it as a variable in some way the engine does not reproduce (`LINEST({1;2;3;5},{1,1,2;2,1,4;3,1,6;4,1,8},TRUE,TRUE)`: degrees of freedom 1 in Excel, 2 here, and the F statistic and the intercept's standard error differ).
  - `FTEST` of a sample with itself is 1 − 3·2^-53 in Excel and 1 − 2^-52 here: the last bits of a number Excel computes by another algorithm can differ, which shows only when the difference is magnified (`(1-FTEST(A1:A8,A1:A8))*2^52` is 1.5 in Excel, 1 here).
  - A LET name that reads as a reference or a boolean (`A1`, `TRUE`): an xlsx file writes LET's names with the prefix `_xlpm.`, so Excel loads `LET(_xlpm.A1,1,_xlpm.A1)` (1); formula text has no prefix, there such a word is a reference or a boolean, and `LET(A1,1,A1)` is a syntax error here, as Excel refuses that text.
  - Iterative calculation starts from 0 at every computation. Excel starts a book it opens from the values the file stores, empty text where it stores none, and goes on from the values it holds at every recalculation and edit. A cycle that converges differs from Excel's within the maximum change, and one that does not, such as a counter, keeps counting in Excel.
  - Excel's AutoCorrect extends a table when a value is typed right below or right of it and fills a calculated column when a formula is typed into an empty column of a table; `setCells` changes the cells typed alone ([docs/tables.md](docs/tables.md)).
  - A table renamed to a name Excel does not take, or to a defined name: Excel's object model makes the name one it takes (`AB12_`, `_1Ставка`, `_?Продажи`, `Rate_1`, `СТАВКА_1`), as its name box refuses it; `renameTable` refuses with `table_name`.
  - Excel ru-RU takes TRUE and FALSE as defined names, and Excel for Windows ru-RU as a table's name; the engine keeps formulas in en-US text, where they are the booleans, and refuses them in every locale. Excel for Mac ru-RU keeps a table named `ИСТИНА` or `ЛОЖЬ`, which Excel for Windows and the engine refuse, as both Excels refuse such a defined name.
  - A typed formula longer than 255 characters (one over a name of 255 letters) is not measured in Excel for Mac, whose AppleScript cuts what it sets to 255 characters; Excel for Windows computes it as the engine does.
  - Text typed into a cell is read by Excel as a number, a date or a boolean where it can be (`2025-01-31` in ru-RU is a date); `setCells` takes a value as given, so a header cell typed so names its column `2025-01-31` here and `31.01.2025` in Excel ru-RU.
  - Excel's Fill Right and fill handle move the column of a structured reference (`SUM(Sales[Qty])` filled right is `SUM(Sales[Price])`); `copyRange`, as copy and paste, keeps it.
  - `VAR`, `STDEV`, `DEVSQ`, `CORREL`, `RSQ`, `SLOPE`, `STEYX` and their kin, the database ones too, are computed here from exact sums; Excel computes them in double precision about a mean it sums plainly, and where the spread is small against the mean, or the residual against the variation, it loses the digits of the answer: of 10,000 values 1E12 + 0, 0.01, 0.02 `VAR.P` is 1.598E-04 in Excel and 6.680E-05 here, as exactly, `CORREL` of them with their steps 0, 1, 2 is 0.647 in Excel and 1 here, and `STEYX` of 1, 2.00000001, 3 on 1, 2, 3 is 0 in Excel and 8.16E-09 here.
  - Where the squares or products of the values leave the range of doubles, Excel answers what is left of them and the engine the exact answer: the squares of 1E-200 and 2E-200 are 0 in Excel, so `STDEV.P` of them is 0 (5E-201 here, within the tests' tolerance) and `CORREL` with 1, 2 and `SLOPE` of 1, 2 on them `#DIV/0!` (1 and 1E+200 here); a sum or a product past 1.8E+308 is `#NUM!` in Excel where the answer fits (`AVERAGE` of 1E+308 three times and its variance 0, the covariance of 1E+308, −1E+308, 1E-200 with 0, 0, 1E+308, 2.2E+107 here).
  - `LINEST` over values far from 0 against their spread loses digits here that Excel keeps (its least squares by Householder QR): of 10,000 values 1E12 + 0, 0.01, 0.02 on their steps 0, 1, 2 the slope is 0.0101547 here and 0.0100098 in Excel, 0.010009765625 exactly, as `SLOPE` gives it here.
  - `WEEKNUM` with an invalid return type is `#NUM!`, as Excel answers it alone; inside one generated book of the random check Excel 16.113 gave a number for `DAY(WEEKNUM(40977,1E+3)<>2.5)` (1, on two runs).
- **Shown text.** Widths are those of Excel for Mac with the default font, Calibri 11; a cell's own font, its indent, wrapping and rotation are not taken into account, characters outside Latin, Cyrillic, digits and common signs count as a digit, and Excel for Windows lays text out at other widths, so near the edge of a column it can show `#` where the engine shows the number or the other way round. Locale tags other than `[$-409]`, `[$-419]`, `[$-F800]` and `[$-F400]`, calendars and numeral systems (`[$-x-sysdate]`, `[DBNum1]`, `B2`) are not read. [docs/formats.md](docs/formats.md) has the details.
- **Syntax.** Stored formulas use English function names and separators (no `;` or decimal commas; text typed in another language goes through `toInvariant`), no R1C1 references, no references to other workbooks. Functions outside the whitelist (`INDIRECT`, `OFFSET`, `LAMBDA`, …) give `#NAME?`.
- **Structure operations and styles.** Whole rows or columns copied onto part of a merged area are refused, where Excel pastes them. Excel's fill handle fills the hidden rows of a filtered sheet, while `copyRange` fills only the visible ones, as Fill Down does. A book has no outlines, conditional formats, validation or comments. Styles are kept and moved, not computed: `compute` does not return them.
- **Output precision.** Values in the result and in the hash are rounded to 15 significant digits, as Excel shows them; inside, computation keeps full double precision.
- **xlsx.** A file's conditional formats, validation, comments, pictures, charts and pivot tables are not read, and each is reported in `issues`. Of a filter the book keeps the filter mode alone: an autofilter's range and criteria are reported, and a filtered sheet is written with criteria that every row meets. Formulas of the old kind are read as the dynamic formulas Excel 365 shows for them, with `@` where it takes one value of a range; the intersection operator is not read. Every formula is written as a dynamic array formula, which Excel 2019 and older show as an array formula. The East Asian date formats 50 to 58 are read as the code Excel for Mac en-US shows for them, `m/d/yy`, where Excel for Windows shows its short date in each locale. [docs/xlsx.md](docs/xlsx.md) has the full list.

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

The tests cover the whitelist against Excel's answers on the coverage measurement's cases (`test/excel.test.ts`) and on the probes (`test/excel-probes.test.ts`), the argument layer rule by rule, the semantics cases against Excel's documentation and the stored LibreOffice answers (`test/oracle/libreoffice.json`, refreshed by `npm run test:oracle`, which needs LibreOffice's `soffice` on `PATH`; the usual tests need no office suite); determinism (shuffled input, a child process, three time zones, cents over 3,400 amounts); references, errors, volatile functions, parsing and the canonical form; sessions (200 random edits against full computations) and `explain`; structure operations against Excel's answers in `test/excel/struct/`, on the synthetic book and on seeded books; budgets; SHA-256 against test vectors and `node:crypto`; a browser bundle run in an empty JavaScript context and with code generation disabled; the two execution modes against each other on every case book and over 200 session edits, the closures mode with code generation disabled; computation in portions and in threads against the synchronous engine, threads lost and sessions restored, the pool on Web Workers in headless Chromium, Firefox and WebKit (through `playwright-core`; each skipped where its engine is not found, `CHROME_PATH` names a Chromium); the published package (exports, tarball contents, a consumer importing it by name). The code blocks of this README run as a test too.

Excel's answers that the tests compare with are stored in `test/excel/`, so the tests need no Excel: the coverage measurement's cases and Excel columns in `coverage/`, the case sets in `cases/` (`cases/ru-RU/` for the sets computed in ru-RU), the probes in `probes.json`, structure operations in `struct/`, spilling in `spill/`, Excel's ru-RU formula text in `locale/`, the files Excel saved for the xlsx subpath in `xlsx/`. They are recomputed only with Excel for Mac, by a step of `npx tsx bench/coverage/run.ts`: `excel en` and `excel ru` for the measurement, `probes` for the probes, `probes <set>` and `probes <set> ru` for a case set, `struct <set>`, `spill <set>`, `xlsx <step>`; `npx tsx bench/locale/snapshot.ts ru` takes Excel's ru-RU formula text. A run keeps the stored answers of unchanged cases and computes the new and changed ones. `bench/coverage/README.md` has the layout, the steps and the rules for driving Excel. Excel for Windows's answers to the same books are in `test/excel/win/`, computed by `npx tsx bench/windows/run.ts` on a Windows machine over SSH (`bench/windows/README.md`).

Functions live in `src/functions/<category>.ts` (`math`, `statistical`, `logical`, `information`, `text`, `date`, `financial`, `lookup`, `engineering`, `database`; old names such as `NORMDIST` sit next to their modern versions), each file with its own registry, a large category assembled from `src/functions/<category>/`; `src/functions.ts` assembles them. A function declares the kinds of its arguments (`args: 'n i?'`, see `src/args.ts`), and the shared argument layer coerces them, so an implementation gets numbers, text, lists or grids and never repeats the coercion rules. Ported code keeps its attribution in the file header. A function joins its registry when `test/excel.test.ts` passes for it; the coverage measurement (`bench/coverage`, `docs/coverage-formulajs.md`) lists the cases and the classes of the remaining formula.js functions.

## License

MIT. Function implementations are ported from [formula.js](https://github.com/formulajs/formulajs) (MIT, Copyright (c) 2014 Sutoiku, Inc.), [jStat](https://github.com/jstat/jstat) (MIT, Copyright (c) 2013 jStat) and [bessel](https://github.com/SheetJS/bessel) (Apache-2.0, Copyright (C) 2013-present SheetJS); [THIRD_PARTY.md](THIRD_PARTY.md) has their license texts and the files they are in.
