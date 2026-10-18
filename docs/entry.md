# Cell entry

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
const bad = parseEntry('=СУММ(A1)+', { locale: 'ru-RU' })
console.log(bad.ok, bad.code, bad.message, bad.at) // → false syntax Формула не закончена 10
```

`xlcalc/entry` gives `parseEntry`, `needsApostrophe` and `entryText` without the engine (about 100 KB minified, no function's implementation), a locale given as an object, for a grid that writes what it types and opens cells for editing without loading the computation.

## In short

`parseEntry` reads text typed into a cell the way Excel reads it in the book's locale, without a book, and gives the cell's new input: a number, a date or a time with the format Excel sets for it (`15.03.2026` in ru-RU gets the short date, `5%` the format `0%`, `$1,234.50` the locale's money), a boolean, an error, text, or a formula in the stored form. The cell typed into keeps its style, meta, note and link, and its format unless Excel changes it: a number typed into a percent cell is percent, a time typed into a date cell sets `h:mm`, the number formats and the book's own formats stay.

```js
import { parseEntry } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

const ru = (text, cell) => parseEntry(text, { locale: ruRU, cell, now: '2026-10-10T12:00:00Z' })
console.log(ru('1 234,5').input, ru('15.03.2026').input, ru('5%').input) // → { value: 1234.5, numberFormat: 4 } { value: 46096, numberFormat: 14 } { value: 0.05, numberFormat: 9 }
console.log(ru('ИСТИНА').input, ru('#Н/Д').input, ru("'007").input, ru('15.3').input) // → { value: true } { formula: '#N/A' } { value: '007' } { value: 46096, numberFormat: 16 }
console.log(ru('=СУММ(A1;1,5)').input, ru('=FOO(1)').input, ru('-C1').input) // → { formula: 'SUM(A1,1.5)' } { formula: 'FOO(1)' } { formula: '-C1' }
const stop = ru('=SUM(A1,1.5)')
console.log(stop.ok, stop.code, stop.message, stop.at) // → false syntax Здесь не может стоять «.» 9
console.log(ru('18:30', { numberFormat: 14, style: 2, meta: { id: 7 }, note: { text: 'checked' } }).input)
// → {
//     value: 0.7708333333333334,
//     numberFormat: 20,
//     style: 2,
//     meta: { id: 7 },
//     note: { text: 'checked' }
//   }
console.log(parseEntry('3/15/2026').input, parseEntry('$1,234.50').input.value, parseEntry('5', { cell: { numberFormat: '0%' } }).input) // → { value: 46096, numberFormat: 14 } 1234.5 { value: 0.05, numberFormat: '0%' }
```

`opts.locale` is the book's locale, a loaded locale's tag or a locale object ([Localization](locale.md)), `opts.cell` the entry context (its effective `numberFormat`, `style`, `meta`, `note` and `link`), `opts.now` the moment whose year completes a date typed without one (`15.3`), and `opts.dateSystem` the book's date system. A `CellInput` or a `CellView` from `getCell` or `getCells` holds only the cell's own number format. To account for a row's or column's format, use the resolved cell from `getViewport`, or supply the effective `numberFormat` yourself:

```js
import { createSession, parseEntry } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', columns: { A: { numberFormat: '0%' } }, cells: { A1: { value: 0.1 } } }] })
const shown = session.getViewport('S', 'A1').cells.A1
console.log(parseEntry('5', { cell: shown }).input) // → { value: 0.05, numberFormat: '0%' }
session.close()
```

`parseEntry` stores the supplied format in its returned input. For editing a session, `setCellText` resolves the inherited format itself and keeps it on its row or column instead of storing it on the cell. Only a formula that does not parse stops the entry, with the check code, its position in the text and a message in the book's locale, as Excel's window does; an unknown function is taken and computes `#NAME?`. Empty text clears the value and keeps the rest of the cell (`null` when nothing is left). The rules, measured case by case in Excel for Mac and Excel for Windows, typed through the object model and on the keyboard, are in [docs/entry.md](entry.md). A number typed into a column too narrow for it widens the column: `entryWidth` gives the width ([Showing values](formats.md)).

A session takes typed text into a cell itself: `setCellText(sheet, cell, text, opts?)` reads it as `parseEntry` does, in the book's locale and date system, by the session's clock and the format the cell shows (its own, else its row's or column's, which the cell does not take as its own), keeps the rest of the cell, and widens the column in the same step of the undo as Excel does, unless `widen: false`; `previewCellText` says the same and changes nothing.

```js
import { createSession, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const session = createSession({
  locale: 'ru-RU',
  styles: [{ font: { bold: true } }],
  sheets: [{
    name: 'S',
    cells: { A1: { value: 1, style: 0 } },
    columns: { B: { numberFormat: '0%' } },
    validations: [{ range: 'C1:C9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning' }],
  }],
}, { clock: { now: '2026-10-10T12:00:00Z', seed: 1 } })
const typed = session.setCellText('S', 'A1', '1234567890')
console.log(typed.width, session.getBook().sheets[0].cells.A1) // → 11.1640625 { style: 0, value: 1234567890 }
console.log(session.setCellText('S', 'B1', '7').changes, session.getBook().sheets[0].cells.B1) // → [ { sheet: 'S', cell: 'B1', value: 0.07, text: '7%' } ] { value: 0.07 }
const refused = session.setCellText('S', 'A2', '=СУММ(A1;1.5)')
console.log(session.previewCellText('S', 'C1', '-2').validation.severity, refused.code, refused.at) // → warning syntax 10
```

A value the cell's rule of data validation does not take is written, with `validation`, the rule, its alert's style (`severity`) and message. A formula that does not read, a change of part of an old-kind array and a formula in a table's header cell are refused (`{ ok: false, code, message, at? }`, the message in the book's locale); text typed into a header cell renames its column, as `setCells` does.

Excel shows the rule's alert before it writes such a value: a `stop` alert never takes it (Retry or Cancel), a `warning` writes it on Yes, an `information` on OK. With `refuseInvalid: true` the call writes nothing and answers `{ ok: false, code: 'validation', sheet, cell, rule, severity, message }`, so the host shows the alert first. Where the person keeps the value, the host makes the same call again without the option, and the value is written with its alert. Without the option the value is written, and a host that refuses it afterwards calls `undo()`, which leaves a step to redo.

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{ name: 'S', cells: {}, validations: [{ range: 'A1:A9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning', error: 'Positive only' }] }],
})
const asked = session.setCellText('S', 'A1', '-2', { refuseInvalid: true })
console.log(asked.code, asked.cell, asked.severity, asked.message, session.revision, session.history) // → validation A1 warning Positive only 0 { undo: 0, redo: 0 }
// the person answers Yes: the same call without the option writes the value
const kept = session.setCellText('S', 'A1', '-2')
console.log(kept.validation.severity, session.getBook().sheets[0].cells.A1, session.revision) // → warning { value: -2 } 1
const many = session.setCellTexts([{ sheet: 'S', cell: 'A2', text: '5' }, { sheet: 'S', cell: 'A3', text: '1.5' }], { refuseInvalid: true })
console.log(many.code, many.cell, session.revision) // → validation A3 1
```

`setCellTexts(entries, opts?)` takes several texts typed at once: each `{ sheet, cell, text }` is read as `setCellText` reads it in its cell, in order, and all are written with one recomputation, one revision and one step of the undo. A cell given twice keeps the later text, read by the format the earlier one gave it. A column widens once, to the largest width a text typed alone into one of its cells would give it (`widths`, unless `widen: false`), and the cells whose values their rules do not take come back with their alerts (`validations`). A text `setCellText` refuses refuses the whole call, with its `sheet` and `cell`, and nothing is written; with `refuseInvalid`, so does the cell whose value breaks the strictest rule once all the texts are written (stop before warning before information, the first cell of that style). `previewCellTexts` answers each cell's input and value, the widths and the alerts, and changes nothing.

```js
import { createSession, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const session = createSession({
  locale: 'ru-RU',
  sheets: [{ name: 'S', cells: { A1: { value: 'Item' } }, validations: [{ range: 'C1:C9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning' }] }],
}, { clock: { now: '2026-10-10T12:00:00Z', seed: 1 } })
const typed = session.setCellTexts([
  { sheet: 'S', cell: 'B1', text: '15.03.2026' },
  { sheet: 'S', cell: 'B2', text: '1234567890' },
  { sheet: 'S', cell: 'C1', text: '-2' },
])
console.log(typed.widths, typed.validations.map((v) => `${v.cell} ${v.severity}`)) // → [ { sheet: 'S', column: 'B', width: 11.1640625 } ] [ 'C1 warning' ]
const refused = session.setCellTexts([{ sheet: 'S', cell: 'A2', text: '7' }, { sheet: 'S', cell: 'A3', text: '=СУММ(A1;1.5)' }])
console.log(refused.code, refused.cell, refused.at, session.revision) // → syntax A3 10 1
console.log(session.previewCellTexts([{ sheet: 'S', cell: 'A2', text: '5%' }, { sheet: 'S', cell: 'A3', text: '=A2*2' }]).cells.map((c) => [c.cell, c.input, c.value]))
// → [
//     [ 'A2', { value: 0.05, numberFormat: 9 }, 0.05 ],
//     [ 'A3', { formula: 'A2*2' }, 0.1 ]
//   ]
```

A cell edited in place, after F2, a double click or in the formula bar, opens with its own text, and Excel reads what comes back otherwise than text typed over the cell: its percent editor adds `%` only to a number begun by a sign. With `inPlace: true`, `setCellText`, `setCellTexts`, `fillCellText` and their previews read the text so: in a percent cell `6%` is 0.06 and `6` is 6 (600%), where typed over the cell `6%` is the text `6%%` and `6` is 6%.

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 0.05, numberFormat: '0%' } } }] })
console.log(session.previewCellText('S', 'A1', '6%').input, session.previewCellText('S', 'A1', '6%', { inPlace: true }).input)
// → { value: '6%%', numberFormat: '0%' } { value: 0.06, numberFormat: '0%' }
```

A paste and Control+Return read text otherwise than typing, and have their own calls. `pasteText(sheet, cell, text, opts?)` pastes text from the clipboard as Excel pastes it: `parseClipboard(text)` makes it a rectangle, rows at line breaks and fields at tabs, every row as wide as the longest, and the rectangle fills the cells from `cell`, an empty field, a short row's missing ones and the row after a last line break clearing theirs. A field begun by a quote runs to its closing quote, line breaks and tabs inside it, and wraps its cell. Each field is read as typed into its cell, but an apostrophe stays part of the text, a formula Excel does not take is pasted as text, a percent cell has no percent editor (`5` is 500%) and `€5` takes the euro format; no column widens and no rule is checked. The answer has the `range` filled; part of an array, a formula into a header cell and a rectangle past the sheet's edge are refused with the `sheet` and `cell`.

`fillCellText(sheet, range, text, opts?)` enters a text into every cell of a range, as Control+Return: it reads the text once in the active cell, the range's top-left one or `active`, and every cell takes what the active cell holds after it, a formula moved as a copy from the active cell moves it (`copyFormula`) and the active cell's number format where it has one. The columns widen as typing widens them (`widths`, unless `widen: false`), and only the active cell's rule is checked (`validation`, or with `refuseInvalid` the refusal before anything is written). Both are one step of the undo, and `previewPasteText` and `previewFillCellText` change nothing. The rules, as measured in Excel, are in [docs/entry.md](entry.md#a-paste-and-controlreturn).

```js
import { createSession, parseClipboard, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const session = createSession({
  locale: 'ru-RU',
  sheets: [{
    name: 'S',
    cells: { A1: { value: 2 }, A2: { value: 3 } },
    columns: { C: { numberFormat: '0%' } },
    validations: [{ range: 'D1:D9', type: 'whole', operator: 'greaterThan', formula1: '0', errorStyle: 'warning' }],
  }],
}, { clock: { now: '2026-10-10T12:00:00Z', seed: 1 } })
const text = "'007\t=СУММ(A1;1.5)\t5\t-2\r\n€5"
console.log(parseClipboard(text)) // → [ [ "'007", '=СУММ(A1;1.5)', '5', '-2' ], [ '€5', '', '', '' ] ]
const pasted = session.pasteText('S', 'A4', text)
console.log(pasted.range, Object.entries(session.getCells('S', pasted.range)).map(([a, c]) => `${a} ${c.text}`), 'validation' in pasted)
// → A4:D5 [ "A4 '007", 'B4 =СУММ(A1;1.5)', 'C4 500%', 'D4 -2', 'A5 € 5' ] false
const filled = session.fillCellText('S', 'D1:D2', '=A1-5')
console.log(filled.changes.map((c) => `${c.cell} ${c.formula} ${c.value}`), filled.validation.severity) // → [ 'D2 A2-5 -2', 'D1 A1-5 -3' ] warning
const refused = session.fillCellText('S', 'E1:E3', '=СУММ(A1)+', { active: 'E2' })
console.log(refused.code, refused.cell, refused.at, session.revision) // → syntax E2 10 2
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

Several texts are checked with all of them written, and the cell whose value breaks the strictest rule (stop, then warning, then information; the first cell of that style) refuses them all, so that a warning answered with Yes does not write a value a later cell's stop does not take; an entry into a selection checks its active cell alone. A rule that shows no alert (`showError: false`) refuses nothing, and a paste checks no rule.

Without a session, a host reads the cell (its own copy, or `cells`), calls `parseEntry` with the format the cell shows and sends the input with `setCells`, which keeps the fields the input leaves out.

A number typed into a column too narrow for it widens the column in Excel when nobody set its width by hand. `setCellText` sets it; without it, `entryWidth(value, { format, locale, width, bestFit })` gives that width for the value the entry computes (the cell's change from `setCells`) and the column as it was, and the host sets it with `setLayout([{ sheet, columns: 'B', set: { width, bestFit: true } }])`. Only typing widens a column, the object model does not; the rules are in [Fitting columns](formats.md#fitting-columns).

## The rules

### Formulas

- Text that begins with `=` is a formula, typed in the locale's formula text (`=СУММ(A1;1,5)` in ru-RU, `=SUM(A1,1.5)` in en-US), translated by `toInvariant` and stored in its canonical form (`SUM(A1,1.5)`).
- A `+` or a `-` that begins no number begins a formula as well: `-C1` is the formula `-C1`, `-qq` the formula `-qq` (`#NAME?`), `+C1` the formula `+C1`. Excel drops a `+` before a number there: `+1+2` is the formula `1+2`, `++5` the formula `+5`. `-5`, `+5`, `- 5` and `-5%` are numbers, and `+`, `-` and `=` alone are text.
- An unknown function or name does not stop the entry: `=FOO(1)` is stored and computes `#NAME?`. A formula that does not parse is refused with its position: `=1+`, `==1`, `-1,5` in en-US (where `,` is no decimal separator), `=SUM(A1,1.5)` in ru-RU. A formula left open at its end is closed, or refused with Excel's correction ([A formula left open](#a-formula-left-open)).
- A formula takes the format of what it computes: `TODAY()` and `DATE()` the short date (14), `NOW()` the date and time (22), `TIME()` the time with AM or PM (18), `PMT()`, `PV()`, `FV()`, `IPMT()`, `PPMT()` and `NPV()` the locale's money with two decimals. The format passes through a sign, parentheses, `SUM`, `MAX`, `MIN`, `AVERAGE` and `ROUND`, a sum (`1+TODAY()`) and a difference less a plain number (`TODAY()-1`); `YEAR(TODAY())`, `EDATE(…)`, `IF(…, TODAY(), …)`, `TODAY()*1` and `NOW()-TODAY()` have none.

### Text

- An apostrophe makes text of what follows it, and is not part of it: `'007` is the text `007`, `'=1+2` the text `=1+2`, `'` alone empty text.
- In a cell formatted as text (`@`) everything typed is text, formulas too.
- Text Excel reads as nothing else stays as typed, spaces too: ` abc `, `5 abc`, `#SPILL!`.

### Numbers

Numbers are read as `VALUE` reads them in the locale ([Localization](locale.md)), with 15 significant digits (`1234567890123456789` is 1234567890123450000), and their form sets the format of a cell in General:

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
- A built-in date, time, percent, money, scientific or fraction format gives way to a typed date, time, percent, money, scientific number or fraction of another kind: `18:30` typed into a cell of format 14 sets 20, `15.03.2026` typed into `0.00E+00` sets 14 (into `0%` too when the cell is edited in place, where no `%` is added: [A percent cell](#a-percent-cell)), `$5` typed into `0%` sets money. A plain or grouped number, a boolean, text and an error keep it. An accounting format (41 to 44) gives way to any of these forms.
- The built-in number formats (1 to 4, `0.00`, `#,##0.00`, and 37 to 40) and the book's own formats (`d mmmm yyyy`, `#,##0.00 "₽"`) are kept whatever is typed.
- A code counts as a built-in format where Excel takes a file's code so: where it is the code of that format in en-US and in the locale alike (`h:mm` is 20, `0%` is 9; `m/d/yyyy` is 14 in en-US and the book's own format in ru-RU). No code counts as a money or accounting format.
- A formula's format follows the same rules (`=TODAY()` typed into `0%` sets 14).

### A percent cell

As a number is typed into a cell whose format shows percent, Excel's editor adds `%` after it: `5` is 5% (0.05), `0.5` in en-US 0.5%, `1 1/2` 1.5%, `1/2` 0.5%, `-5` -5%. It adds it to text begun by a digit or a sign and a digit, so `5%` becomes the text `5%%`, `-5%` the formula `-5%%` (a sign before what is no number begins a formula) and `15.03.2026` the text `15.03.2026%`; `.5`, `(5)`, ` 5`, `abc` and formulas get none.

A cell edited in place, after F2, a double click or in the formula bar, opens with its own text (`5%`), and there the editor adds `%` only to a number begun by a sign: `6%` is 0.06, `6` is 6 (shown 600%), `0.5` in en-US is 0.5 (50%), `-5` is -5% and `+5` is 5%. A host that knows the text was edited in place passes `{ inPlace: true }` to `setCellText`, `setCellTexts`, `fillCellText` or their previews; without it the text reads as typed over the cell.

### A line break

Text with a line break (`\n`, as Excel's editor takes Control+Option+Return on the Mac and Alt+Enter on Windows) is text, and the cell wraps it: Excel gives the cell `alignment.wrap`, typed or set through the object model alike. `setCellText`, `setCellTexts` and `fillCellText` give the cell's style, its own or else its row's or column's, the field (an equal style of the book's table found, else one added; the answer's `styles` is the new table), and fit the row of a cell that wraps as AutoFit fits it, unless its height was set by hand or it is hidden: `a`, a line break and `b` give a row of the default height 32 points, marked `bestFit` ([Fitting rows](formats.md#fitting-rows)). A formula typed over lines wraps nothing; Excel keeps its line breaks in its text, the engine its canonical text without them.

### Empty text

Empty text clears the value or formula and keeps the format, style, meta, note and link: `{ numberFormat: '0.00', meta: 1 }`, or `null` when nothing is left.

## A formula left open

Excel closes what a typed formula leaves open when Return enters it, or asks first. `parseEntry` and a session's `setCellText` answer as it does, for text typed over a cell and edited in place:

- A formula begun by a function, a sign or a space before it or not, where that function's bracket is the only one left open and no `)` was typed, is closed without a word: `=SUM(1,2` is `=SUM(1,2)`, `=-SUM(C1` is `=-SUM(C1)`, `=FOO(1` is `=FOO(1)` and computes `#NAME?`. A formula that still does not read once closed is refused as before (`=ROUND(1.5`, too few arguments).
- Any other formula left open is refused, and Excel's window offers to correct it to the formula with its brackets, quotes and array braces closed and the spaces after `=` dropped. The refusal carries that text as `correction`, for a grid to ask the same question: `=(1+2` offers `=(1+2)`, `=1+SUM(2` offers `=1+SUM(2)`, `=SUM(MAX(1` offers `=SUM(MAX(1))`, `="ab` offers `="ab"`, `= (1+2` offers `=(1+2)`.
- Excel offers no correction when the formula closed would not read, when it calls a name that is no function in the locale (`=SUM((1` in ru-RU, where `SUM` is no Russian name), and for an array constant left open inside a function's brackets or after an operator (`=SUM({1,2`, `=1+{1,2`); a formula begun by an array constant it does correct (`={1,2` offers `={1,2}`, `={1,2}+{3` offers `={1,2}+{3}`).
- Text pasted from the clipboard is not closed: a formula that does not read is pasted as text.

```js
import { parseEntry, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
console.log(parseEntry('=СУММ(1;2', { locale: 'ru-RU' }).input) // → { formula: 'SUM(1,2)' }
const asked = parseEntry('=(1+2')
console.log(asked.ok, asked.code, asked.correction) // → false syntax =(1+2)
console.log(parseEntry('=SUM({1,2').correction) // → undefined
```

### Checked against Excel

Measured on the Mac's keyboard in Excel 16.113 and 16.114 for Mac, en-US and ru-RU, the Russian names typed in the Russian layout: three sets with Return after each text, 165 cases, each set predicted before Excel ran it. The engine gives Excel's answer on all of them but five known differences, none of them about brackets. Not measured: Excel for Windows, which offers the same windows on a screen the harness does not reach.

## A paste and Control+Return

`session.pasteText(sheet, cell, text)` pastes text from the clipboard as Excel pastes it, in both Excels and both locales alike:

- The text is a rectangle from the cell (`parseClipboard`): a row at each line break (`\r\n`, `\r` or `\n`), a field at each tab, every row as wide as the longest. An empty field, a short row's missing fields and the row after a line break that ends the text clear their cells, as empty text typed does (the cell keeps its format, style, note and `meta`).
- A field begun by a quote runs to its closing quote, tabs and line breaks inside it, a doubled quote one quote: `"a` line break `b"` is one cell holding `a` and `b` on two lines, which wraps, though its row keeps its height. What follows the closing quote is joined to the field (`"a"b` is `ab`), one never closed runs to the end of the text, and a quote that does not begin a field is a character (`a"b`, ` "a`). Measured in Excel for Windows (`Worksheet.Paste`, the sets `rows/keys` and `rows/quotes`).
- Each field is read as typed into its cell, with the cell's format, but for four things. An apostrophe is part of the text: `'007` holds `'007`. A formula Excel does not take is pasted as text (`=1+`, `=SUM(A1,1.5)` in ru-RU), and the other cells are written. A percent cell has no percent editor: `5` is 5, shown 500%, `5%` is 0.05, and `0.5` in ru-RU is text. `€5` is 5 in the format `[$€-2]\ #,##0;[Red]\-[$€-2]\ #,##0`.
- No column widens: a long number shows in scientific form, a date as `###`. No rule of data validation is checked, and no alert comes back.
- Part of an old-kind array formula and a formula into a table's header cell are refused, with the cell, as typing refuses them; a paste that covers an array's whole area replaces it. A rectangle past the sheet's last row or column is refused (`off_sheet`). Text pasted into a header cell renames its column.

The answer has the range filled (`B1:D4`); `previewPasteText` gives each of its cells' input and value and changes nothing.

`session.fillCellText(sheet, range, text, { active })` enters a text into every cell of a range, as Control+Return enters it from the selection's active cell, its top-left cell unless `active` names another:

- The text is read once, as typed into the active cell, by the active cell's format: with a percent cell active `5` is 5% (0.05) in every cell, in General too.
- Every cell takes what the active cell holds after it: the value, or the formula with its relative references moved as a copy from the active cell moves them (`=A1*2` entered into B1:B3 is `=A2*2` in B2 and `=A3*2` in B3, as `copyFormula` writes it), and the active cell's number format where it has one. With the active cell in General a cell keeps its own: `5` with a cell in General active is 5 in a percent cell, shown 500%.
- The columns widen as typing widens them (`widths`, unless `widen: false`); a text with a line break wraps every cell and fits their rows, as typing does.
- Only the active cell's rule of data validation is checked, once: its value is written and its alert comes back (`validation`), for the host to take the call back as Excel takes a stop alert's value back from every cell, or with `refuseInvalid` the call is refused with the active cell and nothing is written; a rule over another cell of the range gives no alert.
- A formula Excel does not take is refused with the active cell, as Excel's window refuses it; part of an array and a formula into a header cell are refused with their cell. A range holds 1,048,576 cells at most.

Both are one recomputation, one revision and one step of the undo.

## Copying plain text

`session.copyText(sheet, range, opts?)` prepares a selection's text in the library, from its computed values and effective number formats, including conditional formats. It works in synchronous, inline and pooled sessions, even with `display: false`, and changes no cells, revision or undo history. The host writes the returned string to its clipboard.

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', columns: { A: { width: 3 } }, cells: { A1: { value: 46096, numberFormat: 'yyyy-mm-dd' }, A2: { value: 1.23456789 } } }] })
console.log(JSON.stringify(session.copyText('S', 'A1:A2', { trailingNewline: false })))
// → "2026-03-15\r\n1.23456789"
```

Copy text uses the format without the column's clipping: a narrow date showing `##` copies `2026-03-15`; a narrow General cell showing `1` copies `1.23456789`. The book's locale and date system control the formatting; `trailingNewline` controls only the terminal CRLF (true by default, false for Mac). A General number still uses General's rounding, not the formula bar's full precision.

Rows and columns hidden by hand, by zero height or width, or by a filter are skipped. A merged area's value is in its first cell and its other cells are empty. A spill copies its computed values. Explicit rectangles keep selected blank cells, including trailing blank rows and columns. Whole rows and columns stop at the sheet's used extent. An empty sheet copies one empty field; a selection with no visible rows or columns gives an empty string.

`maxCells` bounds the visible selection (20,000 by default), and `maxChars` bounds its encoded text, including tabs, CRLF and quoted controls (16 Mi characters by default). A larger copy throws `CalcError` with code `limit`, before the cells or oversized output are allocated; it is never silently truncated. An asynchronous call also takes `signal` and `timeoutMs` and prepares the text in one read in the session's thread.

`formatClipboard(rows, opts?)` from `xlcalc/text` writes a rectangle of cell texts for the clipboard. Rows use CRLF and fields use tabs. A field containing a tab, CR or LF is wrapped in quotes, and quotes inside it are doubled. A quote without a control character stays as text. Windows ends the text with CRLF; pass `{ trailingNewline: false }` for Mac.

```js
import { formatClipboard } from 'xlcalc/text'

console.log(JSON.stringify(formatClipboard([['a\nb', 'c"d']])))
// → "\"a\nb\"\tc\"d\r\n"
```

Use `formatClipboard` when the host already has the complete rectangle of formatted texts, without column clipping. It only encodes those strings: it cannot recover a value from `##`. A host with a library session uses `copyText` to select and format the cells. The measurement's Mac copy used dot decimals in both interfaces, Windows used its regional decimal separator; the book's locale models those regional settings, not just the interface language.

`parseClipboard` reads a quoted field as a text paste does ([A paste and Control+Return](#a-paste-and-controlreturn)), but plain text is not a reversible encoding of every grid: the two cells `"first` and `last"` copy as `"first\r\nlast"`, which a text paste reads as one cell `first`, line break, `last`. Excel for Windows in ru-RU, pasting that copy as plain text, kept the two cells and the quotes of `"abc"`: a paste of a copy is not a paste of typed text. Typed text with quotes, the set `quotes` again in ru-RU (`quotes-ru`), Excel for Windows pastes as in en-US, a field begun by a quote read to its closing quote, all 12 cases alike, so `parseClipboard` reads quotes alike in every locale. Carry the rectangle separately when an exact roundtrip matters.

### Checked against Excel

Measured by copying in Excel for Mac and for Windows, en-US and ru-RU: ten cases in each locale on each, predicted before they ran.

## Point mode

While a formula is typed over a cell (Excel's Enter mode: the editing began with a key, not with F2), an arrow either points, enters the text, or does nothing. `refAt(text, caret?, { locale })` from `xlcalc/text` says which, as Excel decides it:

- `insert`: a reference to the cell beside the edited one goes in at the caret. After the formula's first character (`=`, and `+` or `-`, which begin a formula too), an operator, `(`, the locale's list separator (`,` in en-US, `;` in ru-RU), the range's `:`, a sheet's `!`, a point (`=1.`, `=A1.`) and a space (`=A1 D3`, the intersection operator). Before the formula's first character Excel begins a formula anew: the reference goes in after an `=` of its own (`equals: true`), so `=A1` becomes `=D3=A1` and `+1` becomes `=D3+1`.
- `commit`: the text is entered and the arrow moves the selection, as after a value. After anything else, when the formula reads in the locale it is typed in: `=5`, `=A1`, `=SUM(A1)`, `=TRUE`, `="a"`, `=5%`, `=SUM` (entered as a reference to the function), `=1.5` in en-US and `=1,5` in ru-RU. The whole formula decides, wherever the caret is: `=A1+(` with the caret after `A1` does nothing. A formula left open is entered when Return would close it without a word ([A formula left open](#a-formula-left-open): `=SUM(A1` and `=SUM(A1,` are entered), not when Return would bring Excel's window (`=(1+2`, `=SUM(1)+MAX(2` and `=ABS(-(1)` do nothing). Any text that does not begin a formula too, the caret anywhere.
- `none`: nothing happens. When the formula does not read (`=A1)`, `=1e`, `={1`, `=$`, `=1,5` in en-US, `=1.5` and `=#N/A` in ru-RU, which writes `#Н/Д`), after a character the locale's formulas do not have there (`;` in en-US, `,` after a reference in ru-RU, where it is the decimal separator), and inside a string, a quoted sheet's name or a table's brackets.

Once a reference is pointed the host moves it, as Excel does, by the keys alone; the rule above is for an arrow with nothing pointed since the text was typed:

- An arrow moves the reference from its first cell; Shift and an arrow move its far corner (the first Shift and arrow after `=` makes a range from the edited cell: `=C3:D3`); a plain arrow after a range moves from the range's first cell (`=D3:D4` and Down is `=D4`).
- Command and an arrow on the Mac (Control on Windows) jump as on the sheet, to the edge of the data (`=D3` and Command+Down over D3:D6 is `=D6`), with Shift extend to it (`=D3:D6`); Home points at the row's first column, Page Down a screen below.
- Typing after a reference ends it: the next arrow, where `refAt` says `insert`, points from the edited cell again (`=D3+` and Right is `=D3+D3`).
- F2 turns Enter mode to Edit mode, where arrows and Home move the caret, and F2 again back to pointing, where the rule above holds at the caret.

Return enters the formula as typed, closing what it leaves open as [A formula left open](#a-formula-left-open) says (`=SUM(D3` is `=SUM(D3)`).

### Checked against Excel

Measured on the Mac's keyboard in Excel 16 for Mac, en-US and ru-RU: six sets, each predicted before Excel ran it, 138 cases in each locale. Three known differences stay: a reference to a sheet the book does not have, and in en-US `=ROUND(1.5` and `=ROUND(A1*2`, where Excel does nothing as a function has too few arguments, which `refAt` does not read. Not measured: the mouse, and Excel for Windows.

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

## Checked against Excel

Measured in Excel 16.113 for Mac and Excel 16.0.20430 for Windows, en-US and ru-RU: about 2,000 typed texts through the object model and 236 on the Mac's keyboard, each set predicted before Excel ran it. The engine gives Excel for Windows's answer on 1,812 of the 1,983 cases of the object model and the Mac keyboard's on 232 of 236; the others are the known differences above. Pastes, Control+Return, the text a cell opens with and editing in place were measured the same way; the buttons of a validation rule's alert were not.

## Known differences

- English function names and TRUE in a formula typed in ru-RU: Excel ru-RU takes them as unknown names (`=SUM(A1;1,5)` and `=TODAY()` compute `#ИМЯ?` and set no format, `=TRUE` is a name); `parseEntry` reads them as `toInvariant` does, as the function and the boolean.
- A formula's format from a reference: Excel gives `=C2`, `=C2+1`, `=SUM(C2,1)`, `=MAX(C2,C1)` and `=C1+C2` the date format of C2; `parseEntry` has no book, and a reference gives no format. So does Control+Return, and it widens the column for the date (`=A1+A2` over a date in A2, `fill-held` in both locales).
- Excel for Mac's own short date in en-US, `m/d/yy`, is narrower than Windows's `m/d/yyyy`, which the engine follows: a date entered into a selection widens the column here and not in Excel for Mac (`fill-held`).
- A formula typed or pasted over lines (`=1+` Control+Option+Return `2`): Excel keeps the line break in its text, the engine stores the canonical text without it; the value is the same (`rows/keys`, `rows/held`, `rows/quotes`).
- `#GETTING_DATA` (`#ОЖИДАНИЕ_ДАННЫХ` in ru-RU) typed is an error in Excel; the engine has no such error and keeps the text.
- An unknown name typed in another case than before in the book: Excel writes it in its first case (`=Foo*2` after `=foo` is `=foo*2`); `parseEntry` has no book and keeps the case typed.
- A cell's word used as LET's name and again after the LET in one formula: Excel reads the later one as a defined name (`=LET(IW1,5,IW1)+IW1` holds `_xlnm.IW1` and is `#NAME?`); `parseEntry` reads it as the cell.
- Typed on the keyboard, `+5%+1` takes the percent format of its first number, and Excel drops a `+` before a number further into a formula begun by `+` (`+-+5` is `+-5`); `parseEntry` gives `+5%+1` no format and drops only the first `+`.
- Not measured: typing on the keyboard on Windows (Excel over SSH has no screen), so the percent cell's editor and a leading `+` or `-` follow the Mac's keyboard; the ruble sign typed on the keyboard (the layout used has none); the order of entries (Excel's AutoComplete of text from the column above, AutoCorrect).
- Not measured for a paste: a euro sign in a form other than `€5` (decimals, groups, a minus, after the number), which a paste reads as typing does, a paste over a whole array area, past the sheet's edge, over merged cells or into a selection larger than the clipboard's rectangle. For Control+Return: an active cell in a format other than General or percent, and a cell formatted as text in the selection, where every cell taking the active cell's format is the reading of the measured rule; a selection over a whole array area; Control+Return on Windows (no screen over SSH).
- Excel's formula bar keeps the apostrophe a person typed before text that needed none (`'abc`, `'TRUE` in ru-RU); a book holds no such mark, and `entryText` puts the apostrophe only before text that would read as something else.
- Excel for Mac's formula bar writes a ru-RU time with a leading zero (`01:30:00`, `00:00:00`), as its own regional data have it; `entryText` writes the locale's long time without it (`1:30:00`), as Excel for Windows's regional data have it and the engine follows. Excel for Windows's formula bar was not read (it has no screen over SSH).
