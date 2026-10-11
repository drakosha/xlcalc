# Typed entry

`parseEntry(text, opts)` reads what a user typed into a cell the way Excel reads it in the book's locale, and gives the cell's new input: a number, a date or a time with the format Excel sets for it, a boolean, an error, text, or a formula in the stored form. An interface that edits a sheet calls it on the text of its cell editor and passes the input to `setCells`, or calls a session's `setCellText`, which does both. It needs no book, runs in the same thread whatever computes the book, and returns at once.

```js
import { parseEntry, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const now = '2026-10-10T12:00:00Z'
console.log(parseEntry('1 234,5', { locale: 'ru-RU' }).input) // → { value: 1234.5, numberFormat: 4 }
console.log(parseEntry('15.3', { locale: 'ru-RU', now }).input) // → { value: 46096, numberFormat: 16 }
console.log(parseEntry('$1,234.50').input) // → { value: 1234.5, numberFormat: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)' }
console.log(parseEntry('5', { cell: { numberFormat: '0%', style: 3 } }).input) // → { value: 0.05, numberFormat: '0%', style: 3 }
console.log(parseEntry('=СУММ(A1:A3;1,5)', { locale: 'ru-RU' }).input) // → { formula: 'SUM(A1:A3,1.5)' }
const bad = parseEntry('=СУММ(A1;', { locale: 'ru-RU' })
console.log(bad.ok, bad.code, bad.message, bad.at) // → false syntax Формула не закончена 9
```

## API

`parseEntry(text, opts?)` returns `{ ok: true, input }` or `{ ok: false, code, message, at }`.

- `input` is a `CellInput`: `value` or `formula`, `numberFormat` where the cell has one after the entry, and the `style`, `meta`, `note` and `link` of the cell typed into, which Excel keeps whatever is typed. It is `null` when nothing is left of the cell: empty text typed into a cell with none of these.
- `ok: false` is a formula Excel does not take as typed (its window says "There's a problem with this formula"): the check code `syntax`, `empty`, `too_long` or `too_deep`, a message and the position in `text`. Nothing else stops an entry.

The options:

- `locale`: the book's locale, the tag of a loaded locale (`LOCALES`) or a locale object such as `ruRU` of `xlcalc/locales/ru-RU`, which needs no loading here; en-US without it.
- `cell`: the cell typed into, as the book holds it: its `numberFormat`, `style`, `meta`, `note` and `link` (a `CellInput`, or a `CellView` from `Session.getCells`; other fields are not read). Pass the format the cell shows by: its row's or column's when it has none of its own.
- `now`: the moment whose year completes a date typed without one, ISO as `Clock.now`; the current moment without it.
- `dateSystem`: the book's date system, 1900 or 1904, for the serial numbers of dates.

## In a session

`session.setCellText(sheet, cell, text)` does it in a session: it reads the text in the book's locale and date system, by the session's clock and by the format the cell shows (its own, else its row's or column's, which the cell does not take as its own), keeps the rest of the cell, widens the column as Excel does and gives the rule of data validation the value breaks ([Editing a grid](edits.md#typing-text)). `previewCellText` gives the same without writing, and `setCellTexts` takes several texts at once, each read as typed into its cell, in one step of the undo. `pasteText` pastes text from the clipboard and `fillCellText` enters a text into a selection with Control+Return, each as Excel reads it there ([A paste and Control+Return](#a-paste-and-controlreturn)).

```js
import { createSession, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const session = createSession({ locale: 'ru-RU', sheets: [{ name: 'S', cells: { A1: { value: 1, numberFormat: '0.00', meta: { id: 7 } }, A2: { formula: 'A1*2' } } }] })
console.log(session.setCellText('S', 'A1', '1 234,5').changes)
// → [
//     { sheet: 'S', cell: 'A1', value: 1234.5, text: '1234,50' },
//     { sheet: 'S', cell: 'A2', value: 2469, text: '2469' }
//   ]
console.log(session.getBook().sheets[0].cells.A1) // → { meta: { id: 7 }, value: 1234.5, numberFormat: '0.00' }
session.close()
```

A value the cell's rule of data validation does not take is written, with its alert (`validation`). Excel asks first: its alert holds the entry, a `stop` rule never takes the value (Retry or Cancel), a `warning` takes it on Yes and an `information` on OK. With `{ refuseInvalid: true }`, `setCellText`, `setCellTexts`, `fillCellText` and their previews refuse such a value before anything is written: the answer is `{ ok: false, code: 'validation', sheet, cell, rule, severity, message }`, the book, the revision and the history stay, and the host shows the alert. Where the person keeps the value, the host makes the same call without the option, which writes it.

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: {}, validations: [{ range: 'A1:A9', type: 'whole', operator: 'greaterThan', formula1: '0' }] }] })
const r = session.setCellText('S', 'A1', '-2', { refuseInvalid: true })
console.log(r.code, r.cell, r.severity, session.revision) // → validation A1 stop 0
```

Several texts are checked with all of them written, and the first cell whose value its rule does not take refuses them all; an entry into a selection checks its active cell alone. A rule that shows no alert (`showError: false`) refuses nothing, and a paste checks no rule.

Without a session, a host reads the cell (its own copy, or `cells`), calls `parseEntry` with the format the cell shows and sends the input with `setCells`, which keeps the fields the input leaves out.

A number typed into a column too narrow for it widens the column in Excel when nobody set its width by hand. `setCellText` sets it; without it, `entryWidth(value, { format, locale, width, bestFit })` gives that width for the value the entry computes (the cell's change from `setCells`) and the column as it was, and the host sets it with `setLayout([{ sheet, columns: 'B', set: { width, bestFit: true } }])`. Only typing widens a column, the object model does not; the rules are in [Fitting columns](formats.md#fitting-columns).

## The rules

### Formulas

- Text that begins with `=` is a formula, typed in the locale's formula text (`=СУММ(A1;1,5)` in ru-RU, `=SUM(A1,1.5)` in en-US), translated by `toInvariant` and stored in its canonical form (`SUM(A1,1.5)`).
- A `+` or a `-` that begins no number begins a formula as well: `-C1` is the formula `-C1`, `-qq` the formula `-qq` (`#NAME?`), `+C1` the formula `+C1`. Excel drops a `+` before a number there: `+1+2` is the formula `1+2`, `++5` the formula `+5`. `-5`, `+5`, `- 5` and `-5%` are numbers, and `+`, `-` and `=` alone are text.
- An unknown function or name does not stop the entry: `=FOO(1)` is stored and computes `#NAME?`. A formula that does not parse is refused with its position: `=1+`, `==1`, `-1,5` in en-US (where `,` is no decimal separator), `=SUM(A1,1.5)` in ru-RU.
- A formula takes the format of what it computes: `TODAY()` and `DATE()` the short date (14), `NOW()` the date and time (22), `TIME()` the time with AM or PM (18), `PMT()`, `PV()`, `FV()`, `IPMT()`, `PPMT()` and `NPV()` the locale's money with two decimals. The format passes through a sign, parentheses, `SUM`, `MAX`, `MIN`, `AVERAGE` and `ROUND`, a sum (`1+TODAY()`) and a difference less a plain number (`TODAY()-1`); `YEAR(TODAY())`, `EDATE(…)`, `IF(…, TODAY(), …)`, `TODAY()*1` and `NOW()-TODAY()` have none.

### Text

- An apostrophe makes text of what follows it, and is not part of it: `'007` is the text `007`, `'=1+2` the text `=1+2`, `'` alone empty text.
- In a cell formatted as text (`@`) everything typed is text, formulas too.
- Text Excel reads as nothing else stays as typed, spaces too: ` abc `, `5 abc`, `#SPILL!`.

### Numbers

Numbers are read as `VALUE` reads them in the locale ([Localization](../README.md#localization)), with 15 significant digits (`1234567890123456789` is 1234567890123450000), and their form sets the format of a cell in General:

| typed | ru-RU | en-US | format |
|---|---|---|---|
| a plain number | `1,5`, `-1,5`, `(5)` | `1.5`, `-1.5`, `(5)` | none |
| groups of digits | `1 234`, `1 234,5`, `1 2345` | `1,234`, `1,234.5`, `1,234.` | 3 (`#,##0`), or 4 (`#,##0.00`) with a decimal separator |
| percent | `5%`, `1,5%`, `%5`, `(5%)` | `5%`, `1.5%` | 9 (`0%`), or 10 (`0.00%`) with a decimal separator or a fraction (`5 1/2%`) |
| money | `5 ₽`, `1 234,5 ₽`, `-5 ₽` | `$5`, `$1,234.50`, `-$5`, `($5)` | the locale's money with red negatives, two decimals with a decimal separator: `#,##0\ "₽";[Red]\-#,##0\ "₽"`, `"$"#,##0_);[Red]\("$"#,##0\)` |
| scientific | `1,5E+3`, `1e5` | `1.5E+3` | 11 (`0.00E+00`) |
| a fraction | `1 1/2`, `0 1/2`, `1 1/12` | the same | 12 (`# ?/?`), 13 (`# ??/??`) for a denominator of two digits |

A percent keeps the digits typed: `1,1%` in ru-RU (`1.1%` in en-US) is stored as 0.011, the number its decimal text reads as once the point moves two places, as Excel stores it, not 1.1 divided by 100 (0.011000000000000001). A fraction with percent is divided: `1 2/3%` is (1 + 2/3)/100. Money with percent is text in both locales (`$1.1%`, `1,1 ₽%`).

The money format is built from the locale's currency (`RegionalData.currency`): its symbol before or after the amount, the space between and the sign of a negative amount. A euro sign before or after a number is dropped in every locale, with no format (`€5`, `5 €`, `(€5)`). A number in scientific form that Excel cannot hold stays text, in the scientific format: `1E+308` and above, an exponent past 308 or below -324, a mantissa that comes to 0 (`1E-330`); `1E-309` is 0.

### Dates and times

Dates and times are read as a date argument reads them in the locale, with no space before them (` 15.03.2026` is text), and set the date and time formats, which Excel shows by the locale:

| typed | ru-RU | en-US | format |
|---|---|---|---|
| a full date | `15.03.2026`, `15.03.26`, `15/03/2026`, `2026-03-15`, `1-2-3` | `3/15/2026`, `3/15/26`, `2026-03-15`, `1-2-3` | 14 (`dd.mm.yyyy`, `m/d/yyyy`) |
| a day and a month | `15.3`, `15 мар`, `1.5` (1 May), `1/2` (1 February) | `3/15`, `15-Mar`, `1/2` (January 2) | 16 (`dd.mmm`, `d-mmm`), in the year of `now` |
| a month and a year | `мар 2026`, `1.13`, `12/31` (December 1931) | `Mar 2026`, `5/2026` | 17 (`mmm.yy`, `mmm-yy`) |
| a day, a month by name, a year | `15 мар 2026` | `15-Mar-2026`, `March 15, 2026` | 15 (`dd.mmm.yy`, `d-mmm-yy`) |
| a time | `18:30`, `18:30:15`, `6:30 PM`, `6 PM` | the same | 20 (`h:mm`), 21 (`h:mm:ss`), 18 and 19 with AM or PM |
| hours past a day, a fraction of a second | `25:00`, `18:30:15,5` | `25:00`, `18:30:15.5` | 46 (`[h]:mm:ss`), 47 (`mm:ss.0`) |
| a date and a time | `15.03.2026 18:30`, `18:30 15.03.2026` | `3/15/2026 6:30 PM` | 22 (`dd.mm.yyyy h:mm`, `m/d/yyyy h:mm`) |

A year of one digit is of the 2000s (`1-2-3` is 2003), in `VALUE` too.

### Booleans and errors

The locale's words for TRUE and FALSE in any case are booleans: `ИСТИНА`, `истина`, `Ложь` in ru-RU, where `TRUE` is text; `TRUE`, `true` in en-US. Case goes letter by letter, as in Excel for Windows: `falſe`, with the long s, is text (the entry set `name-case-values-held`; Excel for Mac reads it as FALSE). The locale's error texts in any case are errors: `#Н/Д`, `#н/д`, `#ДЕЛ/0!` in ru-RU, where `#N/A` is text. An error is stored as the formula of its literal (`{ formula: '#N/A' }`), as the xlsx reader reads an error a file stores as a value.

### A cell with a format

A format set on the cell before the entry stays, unless what is typed is of another kind than the format shows:

- A cell in General takes the format of the typed form.
- A built-in date, time, percent, money, scientific or fraction format gives way to a typed date, time, percent, money, scientific number or fraction of another kind: `18:30` typed into a cell of format 14 sets 20, `15.03.2026` typed into `0%` sets 14, `$5` typed into `0%` sets money. A plain or grouped number, a boolean, text and an error keep it. An accounting format (41 to 44) gives way to any of these forms.
- The built-in number formats (1 to 4, `0.00`, `#,##0.00`, and 37 to 40) and the book's own formats (`d mmmm yyyy`, `#,##0.00 "₽"`) are kept whatever is typed.
- A code counts as a built-in format where Excel takes a file's code so: where it is the code of that format in en-US and in the locale alike (`h:mm` is 20, `0%` is 9; `m/d/yyyy` is 14 in en-US and the book's own format in ru-RU). No code counts as a money or accounting format.
- A formula's format follows the same rules (`=TODAY()` typed into `0%` sets 14).

### A percent cell

As a number is typed into a cell whose format shows percent, Excel's editor adds `%` after it: `5` is 5% (0.05), `0.5` in en-US 0.5%, `1 1/2` 1.5%, `1/2` 0.5%, `-5` -5%. It adds it to text begun by a digit or a sign and a digit, so `5%` becomes the text `5%%` and `15.03.2026` the text `15.03.2026%`; `.5`, `(5)`, ` 5`, `abc` and formulas get none.

### Empty text

Empty text clears the value or formula and keeps the format, style, meta, note and link: `{ numberFormat: '0.00', meta: 1 }`, or `null` when nothing is left.

## A paste and Control+Return

`session.pasteText(sheet, cell, text)` pastes text from the clipboard as Excel pastes it, in both Excels and both locales alike:

- The text is a rectangle from the cell (`parseClipboard`): a row at each line break (`\r\n`, `\r` or `\n`), a field at each tab, every row as wide as the longest. An empty field, a short row's missing fields and the row after a line break that ends the text clear their cells, as empty text typed does (the cell keeps its format, style, note and `meta`).
- Each field is read as typed into its cell, with the cell's format, but for four things. An apostrophe is part of the text: `'007` holds `'007`. A formula Excel does not take is pasted as text (`=1+`, `=SUM(A1,1.5)` in ru-RU), and the other cells are written. A percent cell has no percent editor: `5` is 5, shown 500%, `5%` is 0.05, and `0.5` in ru-RU is text. `€5` is 5 in the format `[$€-2]\ #,##0;[Red]\-[$€-2]\ #,##0`.
- No column widens: a long number shows in scientific form, a date as `###`. No rule of data validation is checked, and no alert comes back.
- Part of an old-kind array formula and a formula into a table's header cell are refused, with the cell, as typing refuses them; a paste that covers an array's whole area replaces it. A rectangle past the sheet's last row or column is refused (`off_sheet`). Text pasted into a header cell renames its column.

The answer has the range filled (`B1:D4`); `previewPasteText` gives each of its cells' input and value and changes nothing.

`session.fillCellText(sheet, range, text, { active })` enters a text into every cell of a range, as Control+Return enters it from the selection's active cell, its top-left cell unless `active` names another:

- The text is read once, as typed into the active cell, by the active cell's format: with a percent cell active `5` is 5% (0.05) in every cell, in General too.
- Every cell takes what the active cell holds after it: the value, or the formula with its relative references moved as a copy from the active cell moves them (`=A1*2` entered into B1:B3 is `=A2*2` in B2 and `=A3*2` in B3, as `copyFormula` writes it), and the active cell's number format where it has one. With the active cell in General a cell keeps its own: `5` with a cell in General active is 5 in a percent cell, shown 500%.
- The columns widen as typing widens them (`widths`, unless `widen: false`).
- Only the active cell's rule of data validation is checked, once: its value is written and its alert comes back (`validation`), for the host to take the call back as Excel takes a stop alert's value back from every cell, or with `refuseInvalid` the call is refused with the active cell and nothing is written; a rule over another cell of the range gives no alert.
- A formula Excel does not take is refused with the active cell, as Excel's window refuses it; part of an array and a formula into a header cell are refused with their cell. A range holds 1,048,576 cells at most.

Both are one recomputation, one revision and one step of the undo.

## The text a cell opens with

`entryText(value, opts)` is the other way round: the text a cell opens with for editing, as Excel's formula bar shows the value it holds, so that the editor shows what Excel shows and the value reads back from it. `needsApostrophe(text, opts)` tells whether text typed into the cell would read as something else than itself.

```js
import { entryText, needsApostrophe, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)

console.log(entryText(1234.5), entryText(1e20), entryText(1e21), entryText(0.0000001)) // → 1234.5 100000000000000000000 1E+21 0.0000001
console.log(entryText(0.125, { numberFormat: '0.0%' }), entryText(45688.75, { numberFormat: 22 }), entryText(0.75, { numberFormat: 'h:mm', locale: 'ru-RU' }))
// → 12.5% 1/31/2025  6:00:00 PM 18:00:00
console.log(entryText('007'), entryText('abc'), entryText('1.5', { locale: 'ru-RU' }), entryText(true, { locale: 'ru-RU' })) // → '007 abc '1.5 ИСТИНА
console.log(needsApostrophe('5', { numberFormat: '0%' }), needsApostrophe('5', { numberFormat: '@' })) // → true false
```

The options are the cell's `numberFormat` (the one it shows by: its row's or column's without its own), the book's `locale` and `dateSystem`, and `now`, the moment whose year completes text read as a date without one.

A grid without the engine takes the text from its window: `getViewport` and `viewportOf` with `editText: true` give each cell's `editText`, its formula with `=` in the book's locale (`=СУММ(A1:A3)`, `=[@Qty]*2` inside its table), else `entryText` of its value by the format it shows, the session's clock or the result's as `now`, and empty text for a cell without a value or a formula of its own or of its old-kind array's area.

- A number is written whole: its 15 significant digits with the locale's decimal separator and no groups, in positional notation up to 21 characters (`100000000000000000000`), in scientific notation past that (`1E+21`, `1E-20`). Money, fractions, scientific and number formats change nothing (`$1,234.50` opens as `1234.5`, `1 1/2` as `1.5`).
- A number under a percent format is written as a percent: `12%`, `12.5%`, `-5%`.
- A number under a date or time format is written by the locale's short date (the built-in format 14: `1/31/2025`, `31.01.2025`) and long time (`6:00:00 PM`, `18:00:00`), two spaces between them, as the formula bar writes them: a date alone when its time is midnight, a time alone under a format with no day when it is under one day (`6:30:00 PM`), the date and the time otherwise (`1/1/1900  6:00:00 PM` for 1.75 under `h:mm`). Day 0 is `1/0/1900`; a negative number or one past 9999-12-31 is written as a number.
- A boolean and an error are written in the locale's words: `TRUE`, `ЛОЖЬ`, `#DIV/0!`, `#Н/Д`.
- A cell of an old-kind array formula's area opens with the area's formula, every cell of it, as F2 opens it (the formula bar shows it in braces, `{=SUM(C1:C3)}`). A cell a dynamic array spilled into opens empty; selected, Excel's formula bar shows its anchor's formula.
- Text is written as it is, with an apostrophe before it when typing it would read as something else (`needsApostrophe`, by `parseEntry`): a number, a date, a time, a boolean, an error or a formula in the locale (`'007`, `'1.5` in ru-RU where it is a date, `'=1+2`), text a percent cell's editor would make a number (`'5`), text that begins with an apostrophe (`''x`), empty text. Under the text format nothing reads otherwise (`007`).

## How it was checked

The rules were measured in Excel 16.113 for Mac and Excel 16.0.20430 for Windows, in en-US and ru-RU, by the step `entry` of the harness (`bench/coverage/entry.ts`, `bench/windows/entry.ts`). A case types its text into a cell of its own, in General or with a format set before, through the object model as Excel 365 takes typed text in the user's language (`formula2 local` on the Mac, `Formula2Local` on Windows); then the cell's shown text and format codes are read, and the book is saved and read back with `readXlsx` for what the cell holds. On the Mac a set may be typed on the keyboard instead (System Events, an ASCII keyboard layout): the window Excel shows for a formula it does not take ("We found a typo in your formula", then why it is not taken) is answered with No and OK, and Escape leaves the cell as it was. The answers are in `test/excel/entry/` and `test/excel/win/entry/`, and `test/entry-excel.test.ts` and `test/entry-win.test.ts` compare `parseEntry` with every case, on the day of the run.

| set | cases per locale | how | predicted before the run, en-US and ru-RU |
|---|---:|---|---|
| `typed` | 552 | object model, Mac and Windows | Mac 385 and 397, Windows 385 and 378 |
| `typed-keys` | 39 | keyboard, Mac | 30 and 30 |
| `typed-held` | 257 | object model, Mac and Windows | Mac 181 and 184, Windows 186 and 184 |
| `typed-keys-held` | 44 | keyboard, Mac | 36 and 36 |
| `typed-held-2` | 102 | object model, Mac and Windows | 85 and 84 on both |
| `typed-keys-held-2` | 29 | keyboard, Mac | 25 and 26 |
| `typed-pct-held` | 47 | object model, Mac and Windows | 45 and 46 on both |
| `names-held` | 28 (en-US) | object model, Mac and Windows | 21 on both: cells' words as LET's names (`XFD1`, `XFE1`, `A0`) and the words past the sheet alone, refused here and taken by Excel, and an unknown name's case |
| `let-grid-held` | 14 (en-US) | object model, Mac and Windows | 10 on both: `IV1`, `A65536` and `IV65536` are LET's names too, so the edge predicted, Excel 2003's sheet, is none, and a LET name read after its LET |
| `let-cells-held` | 25 (en-US) | object model, Mac and Windows | 11 on both: every cell is LET's name but those that read as R1C1's row or column too |

Each set's predictions (`bench/coverage/entry/<set>.predicted.json`) were committed before Excel ran it, and the misses became rules. The engine now gives Excel for Windows's answer on 1,812 of the 1,983 cases of the object model's sets, and the Mac's keyboard answer on 220 of 224; the other cases are the known differences below, pinned in `test/entry-known.ts`.

Values are compared exactly, as the saved file holds every digit. `typed-pct-held` types percents with decimals where moving the point and dividing by 100 give different numbers (`1.1%`, `0.7%`, `1.255%`, `99.9%`, with a sign, parentheses, a leading `%`, groups and an exponent), fractions and plain decimals; each cell's value is also read through the object model (`value` on the Mac, `Value2` on Windows), which equals the number the file holds in every case. Both Excels stored every decimal percent as its decimal text reads; the miss was `$1.1%`, which Excel en-US keeps as text, the other the sign difference of the object model.

Excel for Windows and Excel for Mac answer alike, but where their own built-in codes part: the Mac's en-US short date is `m/d/yy` and its ru-RU short time `hh:mm`, so it keeps the codes `m/d/yyyy` (en-US) and `h:mm` (ru-RU) as formats of the book's own, where Windows takes them as the formats 14 and 20. The engine follows Windows.

The text a cell opens with was measured by the same step on the sets `edit-held` (85 texts) and `edit-numbers-held` (38 numbers around a length of 21 characters), which read each cell's `Formula` and `FormulaLocal` and, on the Mac, its formula bar itself: each cell selected in turn with Excel in front, the bar's text read through the window's accessibility (`Formula Bar. A,1 1/31/2025  6:00:00 PM`). The bar is needed: both Excels' object model gives a date's, a time's and a percent's number there (`45688`, `0.12`), and under other formats the text the bar shows, the apostrophe being the cell's prefix character. `entryText`'s predictions were committed before Excel ran each set: `edit-held` came out as predicted on 72 of 85 cells in en-US and 74 in ru-RU, the misses being numbers of 21 characters written whole, the two spaces before a time, the apostrophe Excel keeps and ru-RU's times, and with the rules fixed `edit-numbers-held` came out as predicted on all 38 numbers in both locales, in Excel for Mac's formula bar and Excel for Windows's `FormulaLocal` alike. `test/entry-text-excel.test.ts` compares `entryText` with the formula bar and with the object model under the other formats, and reads every text back with `parseEntry` to the value Excel saved.

The object model takes some text otherwise than typing, and the keyboard sets settle these: through `Formula2Local` a text begun by `+` or `-` that is no number is text (`-C1`, `+1+2`), a number goes into a percent cell as it is (`5` is 500%), and empty text makes a cell formatted as text hold empty text. Excel for Windows has `AutoPercentEntry` on, as the answers' `meta` records, but takes `Formula2Local` as the Mac does.

Text pasted from the clipboard and a text entered into a selection with Control+Return were measured by the step `ui` (`bench/coverage/ui.ts`; the sets `paste-held`, 25 pastes on both Excels, and `fill-held`, 18 entries on the Mac's keyboard), and the text a cell of an array opens with by the set `arrays-held` (the formula bar selected and once F2 began the editing, on the Mac; `Formula2Local` on both, which is the text F2 opens). The predictions, `setCellTexts` of the texts and `editText`, were committed before Excel ran each set, and the misses became the rules of [A paste and Control+Return](#a-paste-and-controlreturn) and of the cells of an old-kind array's area; Excel for Mac and Excel for Windows pasted alike in every case. `test/ui-excel.test.ts` compares `pasteText`, `fillCellText` (from the selection's first cell, a call with a stop alert taken back) and `editText` with every case: the engine gives Excel's answer to the 100 pastes (25 in each locale on each Excel), to 33 of the 36 entries into a selection and to the 16 cases of `arrays-held`; the other three are known differences below.

Refusing a value a rule does not take (`refuseInvalid`) follows the buttons of Excel's alert of each style, which the harness does not measure: the alert is a window, and the object model writes a value without one. `test/session-refuse-invalid.test.ts` checks, in each kind of session, that a refused call leaves the book, the revision and the history as they were, and that a value the rules take is written as without the option.

## Known differences

- English function names and TRUE in a formula typed in ru-RU: Excel ru-RU takes them as unknown names (`=SUM(A1;1,5)` and `=TODAY()` compute `#ИМЯ?` and set no format, `=TRUE` is a name); `parseEntry` reads them as `toInvariant` does, as the function and the boolean.
- A formula's format from a reference: Excel gives `=C2`, `=C2+1`, `=SUM(C2,1)`, `=MAX(C2,C1)` and `=C1+C2` the date format of C2; `parseEntry` has no book, and a reference gives no format. So does Control+Return, and it widens the column for the date (`=A1+A2` over a date in A2, `fill-held` in both locales).
- Excel for Mac's own short date in en-US, `m/d/yy`, is narrower than Windows's `m/d/yyyy`, which the engine follows: a date entered into a selection widens the column here and not in Excel for Mac (`fill-held`).
- `#GETTING_DATA` (`#ОЖИДАНИЕ_ДАННЫХ` in ru-RU) typed is an error in Excel; the engine has no such error and keeps the text.
- An unknown name typed in another case than before in the book: Excel writes it in its first case (`=Foo*2` after `=foo` is `=foo*2`); `parseEntry` has no book and keeps the case typed.
- A cell's word used as LET's name and again after the LET in one formula: Excel reads the later one as a defined name (`=LET(IW1,5,IW1)+IW1` holds `_xlnm.IW1` and is `#NAME?`); `parseEntry` reads it as the cell.
- Typed on the keyboard, `+5%+1` takes the percent format of its first number, and Excel drops a `+` before a number further into a formula begun by `+` (`+-+5` is `+-5`); `parseEntry` gives `+5%+1` no format and drops only the first `+`.
- Not measured: typing on the keyboard on Windows (Excel over SSH has no screen), so the percent cell's editor and a leading `+` or `-` follow the Mac's keyboard; the ruble sign typed on the keyboard (the layout used has none); the order of entries (Excel's AutoComplete of text from the column above, AutoCorrect).
- Not measured for a paste: quoted fields on the clipboard (`parseClipboard` takes a quote as text), a euro sign in a form other than `€5` (decimals, groups, a minus, after the number), which a paste reads as typing does, a paste over a whole array area, past the sheet's edge, over merged cells or into a selection larger than the clipboard's rectangle. For Control+Return: an active cell in a format other than General or percent, and a cell formatted as text in the selection, where every cell taking the active cell's format is the reading of the measured rule; a selection over a whole array area; Control+Return on Windows (no screen over SSH).
- Excel's formula bar keeps the apostrophe a person typed before text that needed none (`'abc`, `'TRUE` in ru-RU); a book holds no such mark, and `entryText` puts the apostrophe only before text that would read as something else.
- Excel for Mac's formula bar writes a ru-RU time with a leading zero (`01:30:00`, `00:00:00`), as its own regional data have it; `entryText` writes the locale's long time without it (`1:30:00`), as Excel for Windows's regional data have it and the engine follows. Excel for Windows's formula bar was not read (it has no screen over SSH).
