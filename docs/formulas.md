# Formulas

What a formula can contain and how the engine parses it: arrays inside a formula, functions of a workbook's own (LAMBDA), validating formula text before it is stored, errors, functions a host adds, hidden rows and the 1904 date system.

## Arrays

A range or an array constant used where a function expects a single value, or as an operand, is calculated element-wise, and functions that accept ranges accept arrays the same way. A formula whose result is an array of several values spills: its cell displays the first value, and the cells to the right and below display the rest, as in Excel 365. In the result, a spilled cell has `spillOf`, the address of the formula, and the formula has `spill`, the range it fills; `E2#` refers to that range. If a cell of the range contains a value or a formula, or a merged area overlaps it, the formula returns `#SPILL!` and its issue names that cell.

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

A spilled cell is a value of the workbook: it appears in `sheets[].cells` and in the hash. An array of more than 65,536 values, such as `A:A*2` (1,048,576 rows), lists its cells down to the point where the rest are all one value, and gives that value as the formula's `spillFill` (0 below the data). The design is in `docs/arrays-incell.md`, and spilling with the rules as tested in Excel in `docs/spill.md`.

`@A1:A10` returns the cell of the range in the formula's row or column, or an array's top-left value. `SUM((A1:A3,C1:C3))` sums a union of ranges, and `SUM('Jan:Mar'!B2)` sums B2 on every sheet from Jan to Mar. A space between two ranges is their intersection (`SUM(A1:C3 B2:D4)` sums B2:C3, and the result is `#NULL!` if they do not intersect), and `:` between references that are not a single range joins them (`SUM(A1:INDEX(B1:B9,3))`). [docs/references.md](references.md) has the rules.

A legacy array formula, entered with Ctrl+Shift+Enter, is a cell with `array`, its range:

```js
import { compute } from 'xlcalc'

const cells = { A1: { value: 1 }, A2: { value: 2 }, C1: { formula: 'A1:A2*10', array: 'C1:C3' } }
const out = compute({ sheets: [{ name: 'S', cells }] }).sheets[0].cells
console.log(out.C1.value, out.C2.value, out.C3.value) // → 10 20 { error: 'NA' }
```

The array fills its range and never spills. A session does not allow changing part of the range, just as Excel does not. `fromLegacy` from `xlcalc/xlsx` converts a legacy array formula from a file into the dynamic array formula Excel 365 displays for it, with `@` where Excel inserts it, and `readXlsx` does so for every such formula in a file. [docs/legacy.md](legacy.md) covers both.

## Functions of your own: LAMBDA

`LAMBDA` makes a function out of a calculation, as in Excel 365: `LAMBDA(x,y,x*y)` has the parameters `x` and `y`, seen in its calculation as LET's names are seen in LET's. A function is called where it is written (`LAMBDA(x,x*2)(3)` is 6), through a LET name (`LET(f,LAMBDA(x,x+1),f(2))`), or through a defined name whose formula is a `LAMBDA`, as Name Manager keeps one: a name `Twice` with `LAMBDA(x,x*2)` makes `=Twice(A1)` a call, shown with the name as it is defined. A name's `LAMBDA` may call itself and other names, and the cells that call it depend on the cells its calculation reads. A call by a built-in function's name is the function's, whatever name of that word the book defines, as Excel reads it (a name `Fact` called as `Fact(5)` is `FACT(5)`).

```js
import { compute } from 'xlcalc'

const book = {
  sheets: [{ name: 'S', cells: { A1: { value: 4 }, B1: { formula: 'twice(A1)' }, B2: { formula: 'Down(5)' }, B3: { formula: 'SUM(MAP(A1:A1,LAMBDA(x,x+1)))' }, B4: { formula: 'Twice' } } }],
  names: [
    { name: 'Twice', ref: 'LAMBDA(x,x*2)' },
    { name: 'Down', ref: 'LAMBDA(n,IF(n<=0,0,1+Down(n-1)))' },
  ],
}
const out = compute(book).sheets[0].cells
console.log(out.B1.formula, out.B1.value, out.B2.value, out.B3.value, out.B4.value) // → Twice(A1) 8 5 5 { error: 'CALC' }
```

A call gives each parameter its argument, a reference as a reference; an argument left out between commas is omitted (`LAMBDA(x,y,ISOMITTED(y))(1,)` is TRUE) and reads as a blank. A parameter in brackets is optional, `LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y))`: a call may leave it out, and it reads as a blank. A required parameter after an optional one is refused, as Excel refuses it. A call with fewer arguments than the required parameters or more than all of them is `#VALUE!`. A function is a value of its own: a cell that holds one shows `#CALC!`, an operator or another function given one answers `#VALUE!`, `TYPE` of it is 128, it is no error (`ISERROR` is FALSE, `IFERROR` keeps it), and `ARRAYTOTEXT` writes it `λ`.

A function's name alone is that function as a value, Excel's eta-reduced LAMBDA: `BYROW(A1:C3,SUM)` sums each row, and `LET(f,SUM,f(1,2))` is 3; called with a number of arguments the function does not take, it is `#CALC!`. Alone, a defined name of the same word comes first. `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL` and `MAKEARRAY` call a function for each value, row or column; `GROUPBY` and `PIVOTBY` summarize data by groups with one ([Functions](functions.md)).

Calls of LAMBDAs inside one another, a recursion, go about 200 calls deep for a short calculation, fewer for a longer one; past that a call is `#NUM!`, as Excel answers too many recursive calls, though Excel goes deeper ([Limitations](limitations.md)). A long recursion counts toward a calculation's `timeoutMs` and `signal` as any other work.

## Checking formulas

`validateFormula` checks a formula without a workbook: it returns the canonical text (sheet names quoted, function names in upper case, a range typed with its corners reversed in order, `B2:A1` as `A1:B2`, as Excel shows it) and what the formula uses, or an error code with a message in English and a position.

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

`validateName` checks a name by Excel's rules for a workbook in `locale`, given as a loaded locale's tag or a locale object (en-US if omitted; a ru-RU workbook also does not allow `ИСТИНА` and `ЛОЖЬ`), and rejects it with the rule the name breaks: `chars` (a character or a length Excel does not allow), `reference` (it would be parsed as a cell reference, `Q1`, or an R1C1 reference, `R2C3`) or `boolean`. `validateNames` returns the name issues that a workbook with these sheets, tables and locale would report, without the workbook itself: `dropped` is true for a name the workbook leaves out (one Excel would not accept, a duplicate, an unknown scope sheet, a formula that does not parse, or a reference without its sheet), and false for one it keeps that produces an error where it is used, so a host can reject a name before storing it.

Every check returns `{ ok: true, … }` or a rejection of type `Refusal`: `{ ok: false, code, message, at? }`, with a code for programs, a message for users and, for text, the position in it. This applies to `validateFormula`, `validateLocalFormula`, `toLocal`, `toInvariant`, `parseEntry`, `validateName` and `editStructure`.

A number literal is refused as Excel refuses it typed (`syntax`): from `1E308` on, and one that is not 0 below `1E-309`; one up to the smallest normal number (2.2250738585072014E-308) is 0, and its text `0` (`=1E-309+1` is stored as `0+1`).

Error codes: `empty`, `syntax`, `too_long` (over 8,192 characters), `too_deep` (over 64 levels of nesting, or over 1,024 operators in one formula), `external_ref` (a reference to a workbook that `ctx.externalBooks` has no link to, see [External references](links.md)), `unknown_function`, `unknown_sheet`, `unknown_name` and `unsupported` (an external function called with non-literal arguments). A rejection for `unknown_function` or `unsupported` names its function in `fn`, upper-case, and so does the cell's issue in `diagnostics` and `Session.getCellIssue`. In `compute`, an invalid formula evaluates to `#NAME?` for an unknown function or name, `#REF!` for an unknown sheet or an external reference, and `#ERROR!` otherwise. A formula whose result is an array is valid: it spills.

## Errors

Cell errors appear as `{ error: code }`, with the codes `DIV0`, `NA`, `NAME`, `NUM`, `REF`, `VALUE`, `CYCLE`, `SPILL`, `CALC` and `ERROR`, and those of calls outside the workbook: `CONNECT` and `BLOCKED` as Excel gives them where a connection or a service is not reached, `BUSY` and `GETTING_DATA` for an answer on its way ([External functions](#external-functions)). Every cell with an error is listed in the result's `diagnostics` with a reason (`kind: 'cell'`) that names the cell the error came from (`division by zero, from 'S'!A2`); a session's `getCellIssue` returns the issue of a single cell, with the cell where the error originated ([A viewport for a grid](grid.md#a-viewport-for-a-grid)). A circular reference stops the cells in the cycle and everything that depends on them: they keep their values ([Circular references](cycles.md)) and are each listed with `error: 'CYCLE'`, while the rest of the workbook is calculated. A workbook with [iterative calculation](cycles.md#iterative-calculation-in-short) calculates its cycles instead. `#CYCLE!` itself is the value of a name that refers to itself, directly or through other names.

Input of the wrong shape is a programming error and throws `InputError` (a `TypeError` with `path`), for example `book.sheets[0].cells["b2"]: invalid cell address "b2": expected A1 style such as "B12" (column letters must be upper-case: "B2")`. This covers wrong types, malformed or out-of-grid addresses, duplicate sheet names (compared case-insensitively, as in Excel), unknown fields, invalid dates and clocks, a cell's `meta` that holds a BigInt or a cycle, and unknown sheets in session edits. A known field set to undefined counts as absent: `{ border: { top: undefined } }` is a border without a top, in a cell's style and in a conditional format's alike. A call never silently ignores what it does not accept: an options object or a callback beyond its parameters, or an unknown option, throws an `InputError` that names the call that does accept that option, `getCells: takes 2 arguments, got an object as argument 3 (formulaLocal is an option of getViewport)`; an asynchronous call rejects with it. An array, a number or undefined beyond the parameters is allowed, so `['A1', 'B2'].map(parseCell)` works, since `map` also passes its callback the index and the array.

## External functions

Host functions that are not part of the workbook (for example, a balance from a database) can appear in formulas. `validateFormula(text, { externalFunctions: ['BALANCE'] })` returns the canonical text of each call in `externals`; the host resolves the values in advance and passes them to `compute`:

```js
import { compute } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { formula: 'BALANCE("2026-01")*2' } } }] }
const result = compute(book, { externals: { functions: ['BALANCE'], values: { 'BALANCE("2026-01")': 21 } } })
console.log(result.sheets[0].cells.A1.value) // → 42
```

Arguments of an external function must be literals; a call without a value evaluates to `#N/A`. A value is a cell's value: a number, text, a boolean, `null` or `{ error: code }`; anything else (an unknown error code, a date input, an array, `undefined`) throws `InputError` instead of computing as 0.

### A resolver

`externals.resolve` answers calls at the time of the calculation instead: a call of the host's `functions` without an entry in `values`, with any arguments, and a call of a function of connections (the cube functions, `RTD`, `WEBSERVICE`, `IMAGE`, `STOCKHISTORY`, `TRANSLATE`, `DETECTLANGUAGE`, `CALL`, `REGISTER.ID`, `EUROCONVERT`). It gets the function's name, the call's text with its arguments' values as `values` keys it (`BALANCE({2;3})`), the arguments (a range or an array as rows of values), the formula's cell and an `AbortSignal`. It returns a value, an array (spilled as a formula's array), a cube member `{ value, member }` or `undefined` for no answer: then a function of connections answers as Excel does where it reaches nothing, and a host function is `#N/A`. A thrown error or a rejected promise is no answer too.

```js
import { compute } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { formula: 'BALANCE(A1)*2' }, C1: { formula: 'CUBEMEMBER("Sales","[Time].[2025]")' }, D1: { formula: 'CUBEVALUE("Sales",C1)' } } }] }
const resolve = ({ name, args }) => (name === 'BALANCE' ? args[0] * 10 : name === 'CUBEMEMBER' ? { value: '2025', member: args[1] } : name === 'CUBEVALUE' ? 77 : undefined)
const r = compute(book, { externals: { functions: ['BALANCE'], resolve } })
console.log(r.sheets[0].cells.B1.value, r.sheets[0].cells.C1.value, r.sheets[0].cells.D1.value) // → 40 2025 77
```

A cell a cube function answered with a member shows its caption (`value`), and gives its `member` to the cube functions that read it: `CUBEVALUE("Sales",C1)` is asked as `CUBEVALUE("Sales","[Time].[2025]")`. A calculation asks once for each call text. A session keeps the answers until `refreshConnections(names?)`, which asks again for the calls of the functions named, or of every one, as Excel's Refresh All does, and calculates the formulas calling them, those keeping the value their file stores too ([Values a file stores](#values-a-file-stores)). The answers are values of cells, so the result's `hash` includes them; `inputHash` does not.

An answer may be a promise. Until it settles the cell is `#BUSY!` (a cube function's `#GETTING_DATA`), a call whose argument waits waits with it, and the result's `pending` lists the cells that asked. The synchronous `compute` and session return at once with them; a session's `settleConnections()` waits for the answers and calculates those cells, then the calls their answers lead to, in up to 8 rounds. `createInlineCalc` and a pool wait within the call's `timeoutMs`: past it the call succeeds with the cells still waiting, and the time spent waiting does not count toward the calculation's own deadline (`soft_timeout`). A call's `signal` ends the wait and aborts the signals the resolver got; an answer aborted so is asked again by the next calculation. In a pool the resolver runs in the calling thread: the pool sends it each call a thread makes and the answer back, and a task's hard deadline starts again when an answer arrives.

## Values a file stores

A workbook from a file can use functions the engine cannot calculate: an add-in's (`MYADDIN(A1)`), or one that reaches outside the workbook, a cube's connection (`CUBEVALUE`), a web service (`WEBSERVICE`), Microsoft's services (`STOCKHISTORY`). Excel shows the value the file stores for such a formula until it calculates the cell; the engine does the same with the file's values (`readXlsx`'s `cached`):

- a formula over a function of connections keeps the value its file stores, always;
- a formula over a function the engine does not know keeps it with `unknownFunctions: 'held'`; by default it is `#NAME?`, as Excel answers without the add-in.

The cells that read such a formula are calculated from that value. The formula is listed in `diagnostics` with `fromFile: true` (`error` NAME for a function the engine does not know, CONNECT for one of connections), so a host can show that the value is the file's. Once the cell is edited, or a cell it reads changes, it is calculated: `#NAME?`, or the function's answer without its connection (`CUBEVALUE` `#N/A`, `WEBSERVICE` `#VALUE!`). Without a stored value it is calculated at once. A session's `getBook()` holds the value as the cell's `held`, so the workbook it gives calculates as the session shows.

```js
import { compute } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, B1: { formula: 'MYADDIN(A1)*2' }, C1: { formula: 'B1+1' }, D1: { formula: 'CUBEVALUE("Sales","[Measures].[Amount]")' } } }] }
const cached = { sheets: [{ name: 'S', cells: { B1: { value: 40 }, D1: { value: 7 } } }] }
const r = compute(book, { cached, unknownFunctions: 'held' })
console.log(r.sheets[0].cells.C1.value, r.sheets[0].cells.D1.value) // → 41 7
console.log(r.diagnostics.map((d) => `${d.cell} ${d.error} ${d.fromFile}`)) // → [ 'B1 NAME true', 'D1 CONNECT true' ]
```

## Hidden rows and filters

A row or a column is hidden by `hidden: true` in `rows` or `columns`, as an xlsx file stores it, or by a height or width of 0, as Excel treats it. `SUBTOTAL` with codes 101 to 111 and `AGGREGATE` with options 1, 3, 5 and 7 ignore cells in hidden rows. Hidden columns are included, as in Excel.

A sheet whose AutoFilter has criteria is marked `filtered: true`. Excel treats every hidden row of such a sheet as filtered out, and `SUBTOTAL` with codes 1 to 11 ignores those rows too.

```js
import { compute } from 'xlcalc'

const cells = { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 4 }, B1: { formula: 'SUBTOTAL(9,A1:A3)' }, B2: { formula: 'SUBTOTAL(109,A1:A3)' }, B3: { formula: 'AGGREGATE(9,5,A1:A3)' } }
const sheet = { name: 'S', cells, rows: { 2: { hidden: true } } }
const v = compute({ sheets: [sheet] }).sheets[0].cells
console.log(v.B1.value, v.B2.value, v.B3.value) // → 7 5 5
console.log(compute({ sheets: [{ ...sheet, filtered: true }] }).sheets[0].cells.B1.value) // → 5
```

A session hides and unhides rows and sets the filter mode through `setLayout`, and recalculates these formulas ([Layout](grid.md#layout)). Structure operations treat hidden rows as Excel does: on a filtered sheet they delete and copy only the visible rows ([docs/structure.md](structure.md)).

## Date systems

A workbook uses Excel's 1900 date system unless it has `dateSystem: 1904`, as an xlsx file created on an old Mac may have. In the 1904 system, serial 0 is 1 January 1904. Typed dates, date text, `DATE`, `TODAY` and the date functions count from there, and cells and `TEXT` display a negative date or time with a minus sign.

```js
import { compute } from 'xlcalc'

const cells = { A1: { formula: 'DATE(2025,1,31)' }, A2: { formula: 'YEAR(A1)' }, A3: { value: -0.5, numberFormat: 'h:mm' } }
const c = compute({ dateSystem: 1904, sheets: [{ name: 'S', cells }] }).sheets[0].cells
console.log(c.A1.value, c.A2.value, c.A3.text) // → 44226 2025 -12:00
```

`displayValue` accepts `dateSystem` too. A workbook in the 1904 system has a values hash of its own.

`dateToSerial(text, { dateSystem })` returns the serial number of a date in the form a cell's `value: { date }` accepts, and `serialToDate(serial, { dateSystem })` returns the date of a serial number in that form, without a time zone: the date only, or with the time to the millisecond. Serial 60 in the 1900 system is Excel's phantom 1900-02-29. Neither function accepts a date before the system's day 0 or after 9999-12-31, which Excel does not have: they throw `InputError`. A date or time that does not exist, such as `2025-02-30`, `2025-13-01` or `T25:00`, or text after it, also throws `InputError`, from `dateToSerial` and in a cell's `value: { date }`, instead of becoming another date (Excel's day 0, `1900-01-00`, and its 1900-02-29 are dates). A `Date` object is not accepted, because its time zone could shift the day.

```js
import { dateToSerial, serialToDate } from 'xlcalc'

console.log(dateToSerial('2026-01-31'), dateToSerial('2026-01-31T18:00'), dateToSerial('2026-01-31', { dateSystem: 1904 })) // → 46053 46053.75 44591
console.log(serialToDate(46053.75), serialToDate(60), serialToDate(0, { dateSystem: 1904 })) // → 2026-01-31T18:00 1900-02-29 1904-01-01
```
