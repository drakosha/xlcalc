# Formulas

What a formula can hold and how the engine reads it: arrays inside a formula, checking formula text before it is stored, errors, functions a host adds, hidden rows, the 1904 date system.

## Arrays

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
console.log(out.E1.value, r.diagnostics[0].reason)
// → { error: 'SPILL' } the array of 3×1 values cannot spill: 'S'!E3 is not empty
```

A spilled cell is a value of the book: it is in `sheets[].cells` and in the hash. An array of more than 65,536 values, such as `A:A*2` (1,048,576 rows), lists its cells up to where the rest holds one value and gives that value as the formula's `spillFill` (0 below the data). The design is in `docs/arrays-incell.md`, spilling and the rules measured in Excel in `docs/spill.md`.

`@A1:A10` takes the cell of the range in the formula's row or column, and an array's top-left value. `SUM((A1:A3,C1:C3))` sums a union of ranges, and `SUM('Jan:Mar'!B2)` sums B2 on every sheet from Jan to Mar. A space between two ranges is their intersection (`SUM(A1:C3 B2:D4)` sums B2:C3, `#NULL!` where they do not meet), and `:` between references that are not one range joins them (`SUM(A1:INDEX(B1:B9,3))`). [docs/references.md](references.md) has the rules.

An array formula of the old kind, entered with Ctrl+Shift+Enter, is a cell with `array`, its area:

```js
import { compute } from 'xlcalc'

const cells = { A1: { value: 1 }, A2: { value: 2 }, C1: { formula: 'A1:A2*10', array: 'C1:C3' } }
const out = compute({ sheets: [{ name: 'S', cells }] }).sheets[0].cells
console.log(out.C1.value, out.C2.value, out.C3.value) // → 10 20 { error: 'NA' }
```

The array fills its area and never spills. A session refuses to change part of the area, as Excel does. `fromLegacy` from `xlcalc/xlsx` turns a formula of the old kind from a file into the dynamic formula Excel 365 shows for it, with `@` where Excel puts it, and `readXlsx` does so for every such formula of a file. [docs/legacy.md](legacy.md) has both.

## Checking formulas

`validateFormula` checks a formula without a book: it returns the canonical text (sheet names quoted, function names upper-case) and what the formula uses, or a check code with a message in English and a position.

```js
import { validateFormula, validateName, validateNames } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

console.log(validateFormula("sum(Sales!b1:b3, 'Q1 data'!A1)", { sheets: ['Sales', 'Q1 data'] }).canonical)
// → SUM('Sales'!B1:B3,'Q1 data'!A1)
const bad = validateFormula('SUM(B1:B3', { sheets: ['Sales'] })
console.log(bad.ok, bad.code, bad.message, bad.at) // → false syntax expected ), found end of formula 9
console.log(validateName('Q1').ok, validateName('tax_rate').ok, validateName('Ставка').ok) // → false true true
console.log(validateName('ложь').ok, validateName('ложь', { locale: ruRU }).code) // → true boolean
const names = [{ name: 'R', ref: '1' }, { name: 'rate', ref: "'Q1 data'!B1" }, { name: 'gone', ref: "'Other'!A1" }]
console.log(validateNames(names, { sheets: ['Sales', 'Q1 data'] }).map((x) => `${x.name} ${x.dropped}`)) // → [ 'R true', 'gone false' ]
```

`validateName` checks a name by Excel's rules for a book of `locale`, a loaded locale's tag or a locale object (en-US without one; a ru-RU book also refuses `ИСТИНА` and `ЛОЖЬ`), and refuses with the rule the name breaks: `chars` (a character or a length Excel does not take), `reference` (it reads as a cell, `Q1`, or R1C1, `R2C3`) or `boolean`. `validateNames` gives the name issues a book with these sheets, tables and locale would give, without the book: `dropped` is true for a name the book leaves out (one Excel would not take, a duplicate, an unknown scope sheet, a formula that does not parse or a reference without its sheet), false for one it keeps that gives an error where it is used, so a host can refuse a name before it stores it.

Every check answers `{ ok: true, … }` or a refusal, the type `Refusal`: `{ ok: false, code, message, at? }`, a code a program reads, a message for people and, for text, the position in it. So do `validateFormula`, `validateLocalFormula`, `toLocal`, `toInvariant`, `parseEntry`, `validateName` and `editStructure`.

Check codes: `empty`, `syntax`, `too_long` (over 8,192 characters), `too_deep` (over 64 levels of nesting), `external_ref` (a reference to a workbook `ctx.externalBooks` holds no link to, [External references](links.md)), `unknown_function`, `unknown_sheet`, `unknown_name`, `unsupported` (an external function called with non-literal arguments). In `compute`, an invalid formula gives `#NAME?` for an unknown function or name, `#REF!` for an unknown sheet or an external reference, and `#ERROR!` otherwise. A formula whose result is an array is valid: it spills.

## Errors

Cell errors come out as `{ error: code }` with codes `DIV0`, `NA`, `NAME`, `NUM`, `REF`, `VALUE`, `CYCLE`, `SPILL`, `CALC` and `ERROR`. Every cell with an error is listed in the result's `diagnostics` with a reason (`kind: 'cell'`), which names the cell an error came from (`division by zero, from 'S'!A2`); a session's `getCellIssue` gives one cell's, with the cell the error began in ([A window for a grid](grid.md#a-window-for-a-grid)). A circular reference stops the cells in the cycle and everything that reads them: they hold their values ([Circular references](cycles.md)), each listed with `error: 'CYCLE'`; the rest of the book is computed. A book with [iterative calculation](cycles.md#iterative-calculation-in-short) computes its cycles instead. `#CYCLE!` itself is a name that uses itself, through other names or not.

Input of the wrong shape is a programming error and throws `InputError` (a `TypeError` with `path`), for example `book.sheets[0].cells["b2"]: invalid cell address "b2": expected A1 style such as "B12" (column letters must be upper-case: "B2")`. This covers wrong types, malformed or out-of-grid addresses, duplicate sheet names (compared case-insensitively, as in Excel), unknown fields, bad dates and clocks, and unknown sheets in session edits. A known field given as undefined is no field: `{ border: { top: undefined } }` is a border without a top, in a cell's style as in a conditional format's. A call never ignores what it does not take: an options object or a callback past its arguments, or an option it does not know, is an `InputError` that names the call taking that option, `getCells: takes 2 arguments, got an object as argument 3 (formulaLocal is an option of getViewport)`; an asynchronous call rejects with it. An array, a number or undefined past the arguments passes, so `['A1', 'B2'].map(parseCell)` works, as `map` gives a callback its index and array too.

## External functions

Host functions that are not part of the book (a balance from a database, say) can appear in formulas. `validateFormula(text, { externalFunctions: ['BALANCE'] })` returns the canonical text of each call in `externals`; the host resolves the values in advance and passes them to `compute`:

```js
import { compute } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { formula: 'BALANCE("2026-01")*2' } } }] }
const result = compute(book, { externals: { functions: ['BALANCE'], values: { 'BALANCE("2026-01")': 21 } } })
console.log(result.sheets[0].cells.A1.value) // → 42
```

Arguments of an external function must be literals; a call without a value gives `#N/A`.

## Hidden rows and filters

A row or a column is hidden by `hidden: true` in `rows` or `columns`, as an xlsx file stores it, or by a height or width of 0, as Excel takes it. `SUBTOTAL` with codes 101 to 111 and `AGGREGATE` with options 1, 3, 5 and 7 leave out the cells of hidden rows. Hidden columns count, as in Excel.

A sheet whose autofilter has criteria is marked `filtered: true`. Excel takes every hidden row of such a sheet as one a filter hid, and `SUBTOTAL` with codes 1 to 11 leaves those out too.

```js
import { compute } from 'xlcalc'

const cells = { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 4 }, B1: { formula: 'SUBTOTAL(9,A1:A3)' }, B2: { formula: 'SUBTOTAL(109,A1:A3)' }, B3: { formula: 'AGGREGATE(9,5,A1:A3)' } }
const sheet = { name: 'S', cells, rows: { 2: { hidden: true } } }
const v = compute({ sheets: [sheet] }).sheets[0].cells
console.log(v.B1.value, v.B2.value, v.B3.value) // → 7 5 5
console.log(compute({ sheets: [{ ...sheet, filtered: true }] }).sheets[0].cells.B1.value) // → 5
```

A session hides and shows rows and sets the filter mode through `setLayout`, and recomputes these formulas ([Layout](grid.md#layout)). Structure operations treat hidden rows as Excel does: on a filtered sheet they delete and copy the visible rows only ([docs/structure.md](structure.md)).

## Date systems

A book is in Excel's 1900 date system unless it has `dateSystem: 1904`, as an xlsx file made on an old Mac may have. In the 1904 system serial 0 is 1 January 1904. Dates typed in, date text, `DATE`, `TODAY` and the date functions count from there, and cells and `TEXT` show a negative date or time with a minus.

```js
import { compute } from 'xlcalc'

const cells = { A1: { formula: 'DATE(2025,1,31)' }, A2: { formula: 'YEAR(A1)' }, A3: { value: -0.5, numberFormat: 'h:mm' } }
const c = compute({ dateSystem: 1904, sheets: [{ name: 'S', cells }] }).sheets[0].cells
console.log(c.A1.value, c.A2.value, c.A3.text) // → 44226 2025 -12:00
```

`displayValue` takes `dateSystem` too. A book in the 1904 system has its own values hash.

`dateToSerial(text, { dateSystem })` gives the serial number of a date as `{ date }` takes it, and `serialToDate(serial, { dateSystem })` the date of a serial number in that form, without a time zone: the date alone, or with the time to the millisecond. Serial 60 of the 1900 system is Excel's phantom 1900-02-29. Neither makes a date before the system's day 0 or after 9999-12-31, which Excel does not have: they throw `InputError`. A `Date` object is not taken, as its time zone would move the day.

```js
import { dateToSerial, serialToDate } from 'xlcalc'

console.log(dateToSerial('2026-01-31'), dateToSerial('2026-01-31T18:00'), dateToSerial('2026-01-31', { dateSystem: 1904 })) // → 46053 46053.75 44591
console.log(serialToDate(46053.75), serialToDate(60), serialToDate(0, { dateSystem: 1904 })) // → 2026-01-31T18:00 1900-02-29 1904-01-01
```
