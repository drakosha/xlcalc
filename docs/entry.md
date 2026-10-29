# Cell entry

`parseEntry(text, opts)` parses what a user typed into a cell the way Excel parses it in the workbook's locale, and returns the cell's new input: a number, a date or a time with the number format Excel sets for it, a boolean, an error, text, or a formula in its stored form. An interface that edits a sheet calls it on the text of its cell editor and passes the input to `setCells`, or calls a session's `setCellText`, which does both. It needs no workbook, runs on the same thread whatever calculates the workbook, and returns immediately.

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

`xlcalc/entry` provides `parseEntry`, `needsApostrophe` and `entryText` without the engine (about 100 KB minified, with no function implementations), taking the locale as an object, for a grid that writes what the user types and opens cells for editing without loading the calculation engine.

## In short

`parseEntry` parses text typed into a cell the way Excel parses it in the workbook's locale, without a workbook, and returns the cell's new input: a number, a date or a time with the number format Excel sets for it (`15.03.2026` in ru-RU gets the short date format, `5%` the format `0%`, `$1,234.50` the locale's currency format), a boolean, an error, text, or a formula in its stored form. The cell keeps its style, meta, note and hyperlink, and its number format unless Excel changes it: a number typed into a percentage cell becomes a percentage, a time typed into a date cell sets `h:mm`, and the number formats and the workbook's custom formats stay.

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

`opts.locale` is the workbook's locale, given as a loaded locale's tag or a locale object ([Localization](locale.md)); `opts.cell` is the entry context (its effective `numberFormat`, `style`, `meta`, `note` and `link`); `opts.now` is the moment whose year completes a date typed without one (`15.3`); and `opts.dateSystem` is the workbook's date system. A `CellInput` or a `CellView` from `getCell` or `getCells` holds only the cell's own number format. To take a row's or column's format into account, use the resolved cell from `getViewport`, or supply the effective `numberFormat` yourself:

```js
import { createSession, parseEntry } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', columns: { A: { numberFormat: '0%' } }, cells: { A1: { value: 0.1 } } }] })
const shown = session.getViewport('S', 'A1').cells.A1
console.log(parseEntry('5', { cell: shown }).input) // → { value: 0.05, numberFormat: '0%' }
session.close()
```

`parseEntry` stores the supplied number format in the input it returns. When editing a session, `setCellText` resolves the inherited format itself and leaves it on the row or column instead of storing it on the cell. Only a formula that does not parse stops the entry, with the error code, its position in the text and a message in the workbook's locale, as Excel's alert does; an unknown function is accepted and evaluates to `#NAME?`. Empty text clears the value and keeps the rest of the cell (`null` when nothing is left). The rules, verified case by case in Excel for Mac and Excel for Windows, typed through the object model and on the keyboard, are in [docs/entry.md](entry.md). A number typed into a column that is too narrow for it widens the column: `entryWidth` returns the width ([Number formats and displayed text](formats.md)).

A session accepts typed text into a cell itself: `setCellText(sheet, cell, text, opts?)` parses it as `parseEntry` does, in the workbook's locale and date system, with the session's clock and the number format the cell displays (its own, otherwise its row's or column's, which the cell does not take as its own), keeps the rest of the cell, and widens the column in the same undo step as Excel does, unless `widen: false` is passed; `previewCellText` returns the same and changes nothing.

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

A value the cell's data validation rule does not allow is still written, with `validation`: the rule, its alert's style (`severity`) and message. A formula that does not parse, a change to part of a legacy array formula and a formula in a table's header cell are rejected (`{ ok: false, code, message, at? }`, with the message in the workbook's locale); text typed into a header cell renames its column, as with `setCells`.

Excel shows the rule's alert before it writes such a value: a `stop` alert never accepts it (Retry or Cancel), a `warning` writes it on Yes, and an `information` alert on OK. With `refuseInvalid: true` the call writes nothing and returns `{ ok: false, code: 'validation', sheet, cell, rule, severity, message }`, so the host can show the alert first. If the user keeps the value, the host repeats the call without the option, and the value is written with its alert. Without the option the value is written, and a host that rejects it afterwards calls `undo()`, which leaves a step to redo.

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

`setCellTexts(entries, opts?)` accepts several texts typed at once: each `{ sheet, cell, text }` is parsed as `setCellText` would parse it in its cell, in order, and all are written with one recalculation, one revision and one undo step. A cell given twice keeps the later text, parsed with the number format the earlier one gave it. A column widens once, to the largest width that any one of the texts typed alone into one of its cells would give it (`widths`, unless `widen: false`), and the cells whose values their rules do not allow are returned with their alerts (`validations`). A text that `setCellText` would reject rejects the whole call, with its `sheet` and `cell`, and nothing is written; with `refuseInvalid`, so does the cell whose value breaks the strictest rule once all the texts are written (stop before warning before information, and the first cell with that style). `previewCellTexts` returns each cell's input and value, the widths and the alerts, and changes nothing.

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

A cell edited in place, after F2, a double click or in the formula bar, opens with its own text, and Excel parses what comes back differently from text typed over the cell: its percentage entry adds `%` only to a number that starts with a sign. With `inPlace: true`, `setCellText`, `setCellTexts`, `fillCellText` and their previews parse the text that way: in a percentage cell, `6%` becomes 0.06 and `6` becomes 6 (600%), whereas typed over the cell, `6%` becomes the text `6%%` and `6` becomes 6%.

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 0.05, numberFormat: '0%' } } }] })
console.log(session.previewCellText('S', 'A1', '6%').input, session.previewCellText('S', 'A1', '6%', { inPlace: true }).input)
// → { value: '6%%', numberFormat: '0%' } { value: 0.06, numberFormat: '0%' }
```

Pasting and Control+Return parse text differently from typing, and have their own calls. `pasteText(sheet, cell, text, opts?)` pastes text from the clipboard as Excel pastes it: `parseClipboard(text)` turns it into a rectangle, with rows split at line breaks and fields at tabs, every row as wide as the longest, and the rectangle fills the cells starting at `cell`; an empty field, the missing fields of a short row and the row after a final line break clear their cells. A field that starts with a quote runs to its closing quote, can contain line breaks and tabs, and wraps its cell. Each field is parsed as if typed into its cell, except that a leading apostrophe stays part of the text, a formula Excel does not accept is pasted as text, a percentage cell has no percentage entry (`5` becomes 500%), and `€5` gets the euro format; no column widens and no rule is checked. The result has the `range` filled; part of an array, a formula in a header cell and a rectangle extending past the edge of the sheet are rejected with the `sheet` and `cell`.

`fillCellText(sheet, range, text, opts?)` enters text into every cell of a range, like Control+Return: it parses the text once in the active cell (the range's top-left cell, or `active`), and every cell gets what the active cell contains afterwards, a formula adjusted as copying it from the active cell would adjust it (`copyFormula`), plus the active cell's number format if it has one. The columns widen as typing widens them (`widths`, unless `widen: false`), and only the active cell's rule is checked (`validation`, or with `refuseInvalid` a rejection before anything is written). Both are one undo step, and `previewPasteText` and `previewFillCellText` change nothing. The rules, as verified in Excel, are in [docs/entry.md](entry.md#a-paste-and-controlreturn).

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

- `input` is a `CellInput`: `value` or `formula`, `numberFormat` if the cell has one after the entry, and the `style`, `meta`, `note` and `link` of the cell typed into, which Excel keeps whatever is typed. It is `null` when nothing is left of the cell: empty text typed into a cell with none of these.
- `ok: false` means a formula Excel does not accept as typed (its alert says "There's a problem with this formula"): the error code `syntax`, `empty`, `too_long` or `too_deep`, or `external_ref` for a workbook that `externalBooks` has no link to, a message and the position in `text`. Nothing else stops an entry.

The options:

- `locale`: the workbook's locale, given as the tag of a loaded locale (`LOCALES`) or a locale object such as `ruRU` from `xlcalc/locales/ru-RU`, which does not need to be loaded here; en-US if omitted.
- `cell`: the cell typed into, as the workbook stores it: its `numberFormat`, `style`, `meta`, `note` and `link` (a `CellInput`, or a `CellView` from `Session.getCells`; other fields are not read). Pass the number format the cell displays with: its row's or column's when it has none of its own.
- `now`: the moment whose year completes a date typed without one, in ISO form like `Clock.now`; the current time if omitted.
- `dateSystem`: the workbook's date system, 1900 or 1904, for the serial numbers of dates.
- `inPlace`: true when the text was edited in the cell's own text (F2, a double click, the formula bar) rather than typed over the cell, as `setCellText` takes it ([A percent cell](#a-percent-cell)); the text `entryText` writes is such text. False if omitted.

## In a session

`session.setCellText(sheet, cell, text)` does this in a session: it parses the text in the workbook's locale and date system, with the session's clock and the number format the cell displays (its own, otherwise its row's or column's, which the cell does not take as its own), keeps the rest of the cell, widens the column as Excel does and returns the data validation rule the value breaks ([Editing a grid](edits.md#typing-text)). `previewCellText` returns the same without writing anything, and `setCellTexts` accepts several texts at once, each parsed as if typed into its own cell, in one undo step. `pasteText` pastes text from the clipboard and `fillCellText` enters text into a selection with Control+Return, each parsing the text as Excel does in that case ([A paste and Control+Return](#a-paste-and-controlreturn)).

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

A value the cell's data validation rule does not allow is still written, with its alert (`validation`). Excel asks first: its alert holds the entry, a `stop` rule never accepts the value (Retry or Cancel), a `warning` accepts it on Yes and an `information` alert on OK. With `{ refuseInvalid: true }`, `setCellText`, `setCellTexts`, `fillCellText` and their previews reject such a value before anything is written: the result is `{ ok: false, code: 'validation', sheet, cell, rule, severity, message }`, the workbook, the revision and the history stay as they were, and the host shows the alert. If the user keeps the value, the host repeats the call without the option, which writes it.

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: {}, validations: [{ range: 'A1:A9', type: 'whole', operator: 'greaterThan', formula1: '0' }] }] })
const r = session.setCellText('S', 'A1', '-2', { refuseInvalid: true })
console.log(r.code, r.cell, r.severity, session.revision) // → validation A1 stop 0
```

Several texts are checked after all of them are written, and the cell whose value breaks the strictest rule (stop, then warning, then information; the first cell with that style) rejects them all, so that a warning answered with Yes does not write a value that a later cell's stop rule does not allow; entering text into a selection checks only its active cell. A rule that shows no alert (`showError: false`) rejects nothing, and a paste checks no rule.

Without a session, a host reads the cell (from its own copy, or from `cells`), calls `parseEntry` with the number format the cell displays, and sends the input with `setCells`, which keeps the fields the input leaves out.

A number typed into a column that is too narrow for it widens the column in Excel, unless its width was set manually. `setCellText` sets the width; without it, `entryWidth(value, { numberFormat, locale, width, bestFit })` returns that width for the value the entry produces (the cell's change from `setCells`) and the column's current state, and the host sets it with `setLayout([{ sheet, columns: 'B', set: { width, bestFit: true } }])`. Only typing widens a column, not the object model; the rules are in [Fitting columns](formats.md#fitting-columns).

## The rules

### Formulas

- Text that starts with `=` is a formula, typed in the locale's formula syntax (`=СУММ(A1;1,5)` in ru-RU, `=SUM(A1,1.5)` in en-US), translated by `toInvariant` and stored in its canonical form (`SUM(A1,1.5)`).
- A `+` or a `-` that does not start a number also starts a formula: `-C1` is the formula `-C1`, `-qq` the formula `-qq` (`#NAME?`), `+C1` the formula `+C1`. Excel drops a `+` before a number in that position: `+1+2` is the formula `1+2`, and `++5` the formula `+5`. `-5`, `+5`, `- 5` and `-5%` are numbers, and `+`, `-` and `=` alone are text.
- An unknown function or name does not stop the entry: `=FOO(1)` is stored and evaluates to `#NAME?`. A formula that does not parse is rejected with its position: `=1+`, `==1`, `-1,5` in en-US (where `,` is not a decimal separator), `=SUM(A1,1.5)` in ru-RU. A formula left open at its end is closed, or rejected with Excel's suggested correction ([A formula left open](#a-formula-left-open)).
- A formula gets the number format of what it calculates: `TODAY()` and `DATE()` the short date (14), `NOW()` the date and time (22), `TIME()` the time with AM or PM (18), and `PMT()`, `PV()`, `FV()`, `IPMT()`, `PPMT()` and `NPV()` the locale's currency format with two decimals.
- A reference gets the number format of the cell it refers to, its own or its row's or column's, of any kind: a percentage, a date, money, `0.00`, Text, a custom format. With a percentage in C3 and 2 in C1, `=C3`, `=C1+C3`, `=C3-1` and `=-C3` show a percentage. A range gets its first cell's format: `=SUM(C3:C9)` the format of C3, `=SUM(C:C)` that of C1, and `=SUM(C1:C3)` none when C1 is General, even if the cells after it have one. A reference to another sheet counts too (`=T!C1`); a defined name and a cell of another workbook get none, even when the name refers to a percentage.
- The format carries through a sign, parentheses, `SUM`, `MAX`, `MIN`, `AVERAGE`, `MEDIAN`, `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `TRUNC`, `INT` and `MOD` (their first argument with a format: `=SUM(C1,C3)` and `=ROUND(C1,C3)` get C3's), a sum (its first term with a format: `1+TODAY()`, `=C5+C3` gets C5's money, `=C3+C5` C3's percentage) and a difference (its first term with a format, else its second: `TODAY()-1`, `=C1-C5` gets C5's money). A difference of two dates is a number of days and gets none (`NOW()-TODAY()`, `=C2-C15`); a date minus a time keeps the date's. Other functions and operators get none: `YEAR(TODAY())`, `EDATE(…)`, `IF(…, C3, …)`, `ABS(C3)`, `MROUND(C3,1)`, `CEILING`, `FLOOR`, `SUMPRODUCT`, `INDEX`, `TODAY()*1`, `C3*2`, `C3/2`, `C3^2`, `C3%` and `C3&""`.
- `parseEntry` has no workbook and asks for the formats it needs: `opts.numberFormatAt(sheet, cell)` returns a cell's number format as the workbook stores it, its own or its row's or column's (`sheet` is null for the sheet typed into). It is called at most once for each reference, with the first cell of a range, so `=SUM(C:C)` asks for C1 alone. Without it, a reference gets no format. A session's `setCellText`, `pasteText`, `fillCellText` and AutoSum read the formats from the workbook themselves.

```js
import { parseEntry } from 'xlcalc'

const formats = { C1: 9, C2: 14 }
const numberFormatAt = (sheet, cell) => (sheet === null ? formats[cell] : undefined)
console.log(parseEntry('=C1+5', { numberFormatAt }).input, parseEntry('=C2-C2', { numberFormatAt }).input) // → { formula: 'C1+5', numberFormat: 9 } { formula: 'C2-C2' }
```

### Text

- An apostrophe makes what follows it text, and is not part of the text: `'007` is the text `007`, `'=1+2` the text `=1+2`, and `'` alone empty text.
- In a cell formatted as Text (`@`), everything typed is text, formulas included.
- Text that Excel does not recognize as anything else stays as typed, including spaces: ` abc `, `5 abc`, `#SPILL!`.

### Numbers

Numbers are parsed as `VALUE` parses them in the locale ([Localization](locale.md)), with 15 significant digits (`1234567890123456789` becomes 1234567890123450000), and their form sets the number format of a cell formatted as General:

| typed | ru-RU | en-US | format |
|---|---|---|---|
| a plain number | `1,5`, `-1,5`, `(5)` | `1.5`, `-1.5`, `(5)` | none |
| digit groups | `1 234`, `1 234,5`, `1 2345` | `1,234`, `1,234.5`, `1,234.` | 3 (`#,##0`), or 4 (`#,##0.00`) with a decimal separator |
| percentage | `5%`, `1,5%`, `%5`, `(5%)` | `5%`, `1.5%` | 9 (`0%`), or 10 (`0.00%`) with a decimal separator or a fraction (`5 1/2%`) |
| currency | `5 ₽`, `1 234,5 ₽`, `-5 ₽` | `$5`, `$1,234.50`, `-$5`, `($5)` | the locale's currency format with red negatives, with two decimals if there is a decimal separator: `#,##0\ "₽";[Red]\-#,##0\ "₽"`, `"$"#,##0_);[Red]\("$"#,##0\)` |
| scientific | `1,5E+3`, `1e5` | `1.5E+3` | 11 (`0.00E+00`) |
| a fraction | `1 1/2`, `0 1/2`, `1 1/12` | the same | 12 (`# ?/?`), or 13 (`# ??/??`) for a two-digit denominator |

A percentage keeps the digits typed: `1,1%` in ru-RU (`1.1%` in en-US) is stored as 0.011, the number its decimal text represents once the decimal point moves two places, as Excel stores it, not 1.1 divided by 100 (0.011000000000000001). A fraction with a percent sign is divided: `1 2/3%` is (1 + 2/3)/100. Currency with a percent sign is text in both locales (`$1.1%`, `1,1 ₽%`).

The currency format is built from the locale's currency (`RegionalData.currency`): the symbol before or after the amount, the space between them and the sign of a negative amount. A euro sign before or after a number is dropped in every locale, with no number format set (`€5`, `5 €`, `(€5)`). A number in scientific form that Excel cannot represent stays text, in the scientific format: `1E+308` and above, an exponent above 308 or below -324, or a mantissa that rounds to 0 (`1E-330`); `1E-309` becomes 0.

### Dates and times

Dates and times are parsed as a date argument is parsed in the locale, with no leading space (` 15.03.2026` is text), and set the date and time formats, which Excel displays according to the locale:

| typed | ru-RU | en-US | format |
|---|---|---|---|
| a full date | `15.03.2026`, `15.03.26`, `15/03/2026`, `2026-03-15`, `1-2-3` | `3/15/2026`, `3/15/26`, `2026-03-15`, `1-2-3` | 14 (`dd.mm.yyyy`, `m/d/yyyy`) |
| a day and a month | `15.3`, `15 мар`, `1.5` (1 May), `1/2` (1 February) | `3/15`, `15-Mar`, `1/2` (January 2) | 16 (`dd.mmm`, `d-mmm`), in the year of `now` |
| a month and a year | `мар 2026`, `1.13`, `12/31` (December 1931) | `Mar 2026`, `5/2026`, `2/29` in a year that is not leap (February 2029) | 17 (`mmm.yy`, `mmm-yy`) |
| a day, a month name, a year | `15 мар 2026` | `15-Mar-2026`, `March 15, 2026` | 15 (`dd.mmm.yy`, `d-mmm-yy`) |
| a time | `18:30`, `18:30:15`, `6:30 PM`, `6 PM` | the same | 20 (`h:mm`), 21 (`h:mm:ss`), 18 and 19 with AM or PM |
| hours beyond a day, a fraction of a second | `25:00`, `18:30:15,5` | `25:00`, `18:30:15.5` | 46 (`[h]:mm:ss`), 47 (`mm:ss.0`) |
| a date and a time | `15.03.2026 18:30`, `18:30 15.03.2026` | `3/15/2026 6:30 PM` | 22 (`dd.mm.yyyy h:mm`, `m/d/yyyy h:mm`) |

A one-digit year is in the 2000s (`1-2-3` is 2003), in `VALUE` too.

### Booleans and errors

The locale's words for TRUE and FALSE, in any case, are booleans: `ИСТИНА`, `истина`, `Ложь` in ru-RU, where `TRUE` is text; `TRUE`, `true` in en-US. Case is compared letter by letter, as in Excel for Windows: `falſe`, with the long s, is text (the entry set `name-case-values-held`; Excel for Mac parses it as FALSE). The locale's error texts, in any case, are errors: `#Н/Д`, `#н/д`, `#ДЕЛ/0!` in ru-RU, where `#N/A` is text. An error is stored as a formula of its literal (`{ formula: '#N/A' }`), the same way the xlsx reader reads an error a file stores as a value.

### A cell with a number format

A number format set on the cell before the entry stays, unless what is typed is of a different kind than the format displays:

- A cell formatted as General gets the format of the typed form.
- A built-in date, time, percentage, currency, scientific or fraction format is replaced by a typed date, time, percentage, currency amount, scientific number or fraction of a different kind: `18:30` typed into a cell with format 14 sets 20, `15.03.2026` typed into `0.00E+00` sets 14 (into `0%` too when the cell is edited in place, where no `%` is added: [A percent cell](#a-percent-cell)), and `$5` typed into `0%` sets the currency format. A plain or grouped number, a boolean, text and an error keep it. An accounting format (41 to 44) is replaced by any of these forms.
- The built-in number formats (1 to 4, `0.00`, `#,##0.00`, and 37 to 40) and the workbook's custom formats (`d mmmm yyyy`, `#,##0.00 "₽"`) are kept whatever is typed.
- A code counts as a built-in format where Excel treats a file's code that way: where it is the code of that format both in en-US and in the locale (`h:mm` is 20, `0%` is 9; `m/d/yyyy` is 14 in en-US and a custom format in ru-RU). No code counts as a currency or accounting format.
- A formula's number format follows the same rules, the format a reference gives it as well: `=TODAY()` typed into `0%` sets 14, a reference to money typed into a date cell sets the money format, and a reference to a percentage typed into `0.00` or into a custom format keeps it. A cell that took a format from an earlier formula is such a cell: `=C1` typed over a formula that made it `0%` keeps `0%`.

### A percent cell

When a number is typed into a cell whose format displays a percentage, Excel's editor appends `%` to it: `5` becomes 5% (0.05), `0.5` in en-US 0.5%, `1 1/2` 1.5%, `1/2` 0.5%, and `-5` -5%. It appends it to text that starts with a digit, or a sign and a digit, so `5%` becomes the text `5%%`, `-5%` the formula `-5%%` (a sign before something that is not a number starts a formula) and `15.03.2026` the text `15.03.2026%`; `.5`, `(5)`, ` 5`, `abc` and formulas get none.

A cell edited in place, after F2, a double click or in the formula bar, opens with its own text (`5%`), and there the editor appends `%` only to a number that starts with a sign: `6%` becomes 0.06, `6` becomes 6 (displayed as 600%), `0.5` in en-US becomes 0.5 (50%), `-5` becomes -5% and `+5` becomes 5%. A host that knows the text was edited in place passes `{ inPlace: true }` to `parseEntry`, `setCellText`, `setCellTexts`, `fillCellText` or their previews; without it, the text is parsed as typed over the cell.

### A line break

Text with a line break (`\n`, which Excel's editor enters with Control+Option+Return on the Mac and Alt+Enter on Windows) is text, and the cell wraps it: Excel sets `alignment.wrap` on the cell, whether typed or set through the object model. `setCellText`, `setCellTexts` and `fillCellText` set this field on the cell's style, its own or otherwise its row's or column's (an equal style is found in the workbook's table, otherwise one is added; the result's `styles` is the new table), and fit the row of a wrapping cell as AutoFit fits it, unless its height was set manually or it is hidden: `a`, a line break and `b` give a row of the default height 32 points, marked `bestFit` ([Fitting rows](formats.md#fitting-rows)). A formula typed over several lines wraps nothing; Excel keeps its line breaks in its text, and the engine stores the canonical text without them.

### Empty text

Empty text clears the value or formula and keeps the number format, style, meta, note and hyperlink: `{ numberFormat: '0.00', meta: 1 }`, or `null` when nothing is left.

## A formula left open

Excel closes what a typed formula leaves open when Return enters it, or asks first. `parseEntry` and a session's `setCellText` behave the same, for text typed over a cell and edited in place:

- A formula that starts with a function, with or without a sign or a space before it, where that function's parenthesis is the only one left open and no `)` was typed, is closed silently: `=SUM(1,2` becomes `=SUM(1,2)`, `=-SUM(C1` becomes `=-SUM(C1)`, and `=FOO(1` becomes `=FOO(1)` and evaluates to `#NAME?`. A formula that still does not parse once closed is rejected as before (`=ROUND(1.5`, too few arguments).
- Any other formula left open is rejected, and Excel's alert offers to correct it to the formula with its parentheses, quotes and array braces closed and the spaces after `=` removed. The rejection carries that text as `correction`, so a grid can ask the same question: `=(1+2` suggests `=(1+2)`, `=1+SUM(2` suggests `=1+SUM(2)`, `=SUM(MAX(1` suggests `=SUM(MAX(1))`, `="ab` suggests `="ab"`, and `= (1+2` suggests `=(1+2)`.
- Excel offers no correction when the closed formula would not parse, when it calls a name that is not a function in the locale (`=SUM((1` in ru-RU, where `SUM` is not a Russian name), and for an array constant left open inside a function's parentheses or after an operator (`=SUM({1,2`, `=1+{1,2`); it does correct a formula that starts with an array constant (`={1,2` suggests `={1,2}`, and `={1,2}+{3` suggests `={1,2}+{3}`).
- Text pasted from the clipboard is not closed: a formula that does not parse is pasted as text.

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

Tested on the Mac keyboard in Excel 16.113 and 16.114 for Mac, en-US and ru-RU, with the Russian names typed in the Russian keyboard layout: three sets with Return after each text, 165 cases, each set predicted before Excel ran it. The engine matches Excel on all of them except five known differences, none of them about parentheses. Not tested: Excel for Windows, which shows the same alerts on a screen the test harness cannot reach.

## A paste and Control+Return

`session.pasteText(sheet, cell, text)` pastes text from the clipboard as Excel pastes it, identically in both versions of Excel and both locales:

- The text becomes a rectangle starting at the cell (`parseClipboard`): a new row at each line break (`\r\n`, `\r` or `\n`), a new field at each tab, and every row as wide as the longest. An empty field, the missing fields of a short row and the row after a line break that ends the text clear their cells, as typing empty text does (the cell keeps its number format, style, note and `meta`).
- A field that starts with a quote runs to its closing quote, can contain tabs and line breaks, and a doubled quote inside it is one quote: `"a`, a line break, `b"` is one cell containing `a` and `b` on two lines, which wraps, although its row keeps its height. Text after the closing quote is appended to the field (`"a"b` becomes `ab`), a field whose quote is never closed runs to the end of the text, and a quote that does not start a field is an ordinary character (`a"b`, ` "a`). Tested in Excel for Windows (`Worksheet.Paste`, the sets `rows/keys` and `rows/quotes`).
- Each field is parsed as if typed into its cell, with the cell's number format, except in four ways. A leading apostrophe is part of the text: `'007` stays `'007`. A formula Excel does not accept is pasted as text (`=1+`, `=SUM(A1,1.5)` in ru-RU), and the other cells are written. A percentage cell has no percentage entry: `5` becomes 5, displayed as 500%, `5%` becomes 0.05, and `0.5` in ru-RU is text. `€5` becomes 5 in the format `[$€-2]\ #,##0;[Red]\-[$€-2]\ #,##0`.
- No column widens: a long number is displayed in scientific form, a date as `###`. No data validation rule is checked, and no alert is returned.
- Part of a legacy array formula and a formula in a table's header cell are rejected, with the cell, as typing rejects them; a paste that covers an array's entire range replaces it. A rectangle extending past the sheet's last row or column is rejected (`off_sheet`). Text pasted into a header cell renames its column.

The result has the range filled (`B1:D4`); `previewPasteText` returns the input and value of each of its cells and changes nothing.

`session.fillCellText(sheet, range, text, { active })` enters text into every cell of a range, as Control+Return enters it from the selection's active cell, which is the top-left cell unless `active` names another:

- The text is parsed once, as if typed into the active cell, with the active cell's number format: with a percentage cell active, `5` becomes 5% (0.05) in every cell, including General ones.
- Every cell gets what the active cell contains afterwards: the value, or the formula with its relative references adjusted as copying from the active cell would adjust them (`=A1*2` entered into B1:B3 becomes `=A2*2` in B2 and `=A3*2` in B3, as `copyFormula` writes it), plus the active cell's number format if it has one, including the format its formula takes from its references: `=A1+A3` entered into B1:B3 over a percentage in A1 shows a percentage in every cell, though `=A3+A5` in B3 alone would take A5's money. With a General cell active, each cell keeps its own format: `5` entered with a General cell active is 5 in a percentage cell, displayed as 500%.
- The columns widen as typing widens them (`widths`, unless `widen: false`); text with a line break wraps every cell and fits their rows, as typing does.
- Only the active cell's data validation rule is checked, once: its value is written and its alert is returned (`validation`), so the host can revert the call, as Excel removes a value rejected by a stop alert from every cell; or with `refuseInvalid` the call is rejected with the active cell and nothing is written. A rule over another cell of the range produces no alert.
- A formula Excel does not accept is rejected with the active cell, as Excel's alert rejects it; part of an array and a formula in a header cell are rejected with their cell. A range can hold at most 1,048,576 cells.

Each is one recalculation, one revision and one undo step.

## Copying plain text

`session.copyText(sheet, range, opts?)` builds a selection's text in the library, from its calculated values and effective number formats, including conditional formats. It works in synchronous, inline and pooled sessions, even with `display: false`, and changes no cells, revision or undo history. The host writes the returned string to its clipboard.

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', columns: { A: { width: 3 } }, cells: { A1: { value: 46096, numberFormat: 'yyyy-mm-dd' }, A2: { value: 1.23456789 } } }] })
console.log(JSON.stringify(session.copyText('S', 'A1:A2', { trailingNewline: false })))
// → "2026-03-15\r\n1.23456789"
```

Copied text uses the number format without clipping to the column width: a narrow date displayed as `##` copies as `2026-03-15`, and a narrow General cell displaying `1` copies as `1.23456789`. The workbook's locale and date system control the formatting; `trailingNewline` controls only the final CRLF (true by default, false for Mac). A General number still uses General's rounding, not the formula bar's full precision.

Rows and columns hidden manually, by a zero height or width, or by a filter are skipped. A merged area's value is in its first cell, and its other cells are empty. A spill copies its calculated values. Explicit rectangles keep selected blank cells, including trailing blank rows and columns. Entire rows and columns stop at the sheet's used range. An empty sheet copies one empty field; a selection with no visible rows or columns returns an empty string.

`maxCells` bounds the visible selection (20,000 by default), and `maxChars` bounds its encoded text, including tabs, CRLF and quoted control characters (16 Mi characters by default). A larger copy throws `CalcError` with code `limit`, before the cells or oversized output are allocated; it is never silently truncated. An asynchronous call also accepts `signal` and `timeoutMs` and builds the text in a single read on the session's thread.

`formatClipboard(rows, opts?)` from `xlcalc/text` writes a rectangle of cell texts for the clipboard. Rows use CRLF and fields use tabs. A field containing a tab, CR or LF is wrapped in quotes, and quotes inside it are doubled. A quote without a control character stays as is. Windows ends the text with CRLF; pass `{ trailingNewline: false }` for Mac.

```js
import { formatClipboard } from 'xlcalc/text'

console.log(JSON.stringify(formatClipboard([['a\nb', 'c"d']])))
// → "\"a\nb\"\tc\"d\r\n"
```

Use `formatClipboard` when the host already has the complete rectangle of formatted texts, without column clipping. It only encodes those strings: it cannot recover a value from `##`. A host with a library session uses `copyText` to select and format the cells. In the tests, the Mac copy used decimal points in both interface languages, while Windows used its regional decimal separator; the workbook's locale models those regional settings, not just the interface language.

`parseClipboard` parses a quoted field as a text paste does ([A paste and Control+Return](#a-paste-and-controlreturn)), but plain text cannot encode every grid reversibly: the two cells `"first` and `last"` copy as `"first\r\nlast"`, which a text paste reads as a single cell `first`, line break, `last`. Excel for Windows in ru-RU, pasting that copy as plain text, kept the two cells and the quotes of `"abc"`: pasting a copy is not the same as pasting typed text. Typed text with quotes (the `quotes` set again, in ru-RU: `quotes-ru`) Excel for Windows pastes as in en-US, reading a field that starts with a quote up to its closing quote, in all 12 cases, so `parseClipboard` handles quotes the same way in every locale. Carry the rectangle separately when an exact round trip matters.

### Checked against Excel

Tested by copying in Excel for Mac and for Windows, en-US and ru-RU: ten cases in each locale on each, predicted before they ran.

## Point mode

While a formula is being typed over a cell (Excel's Enter mode: editing started with a keystroke, not with F2), an arrow key either inserts a reference, enters the text, or does nothing. `refAt(text, caret?, { locale })` from `xlcalc/text` tells which, as Excel decides it:

- `insert`: a reference to the adjacent cell is inserted at the caret. This happens after the formula's first character (`=`, or `+` or `-`, which also start a formula), an operator, `(`, the locale's list separator (`,` in en-US, `;` in ru-RU), the `:` of a range, a sheet's `!`, a decimal point (`=1.`, `=A1.`) and a space (`=A1 D3`, the intersection operator). Before the formula's first character, Excel starts a new formula: the reference is inserted after an `=` of its own (`equals: true`), so `=A1` becomes `=D3=A1` and `+1` becomes `=D3+1`.
- `commit`: the text is entered and the arrow moves the selection, as after a value. This happens after anything else, when the formula parses in the locale it is typed in: `=5`, `=A1`, `=SUM(A1)`, `=TRUE`, `="a"`, `=5%`, `=SUM` (entered as a reference to the function), `=1.5` in en-US and `=1,5` in ru-RU. The whole formula decides, wherever the caret is: `=A1+(` with the caret after `A1` does nothing. A formula left open is entered when Return would close it silently ([A formula left open](#a-formula-left-open): `=SUM(A1` and `=SUM(A1,` are entered), but not when Return would show Excel's alert (`=(1+2`, `=SUM(1)+MAX(2` and `=ABS(-(1)` do nothing). Any text that does not start a formula is entered too, wherever the caret is.
- `none`: nothing happens. This is the case when the formula does not parse (`=A1)`, `=1e`, `={1`, `=$`, `=1,5` in en-US, `=1.5` and `=#N/A` in ru-RU, where it is written `#Н/Д`), after a character the locale's formulas do not allow in that position (`;` in en-US, `,` after a reference in ru-RU, where it is the decimal separator), and inside a string, a quoted sheet name or a table's brackets.

Once a reference is inserted, the host moves it by keys alone, as Excel does; the rule above applies to an arrow key pressed when no reference has been inserted since the text was typed:

- An arrow moves the reference from its first cell; Shift with an arrow moves its far corner (the first Shift+arrow after `=` makes a range from the edited cell: `=C3:D3`); a plain arrow after a range moves from the range's first cell (`=D3:D4` and Down gives `=D4`).
- Command and an arrow on the Mac (Ctrl on Windows) jump as on the sheet, to the edge of the data (`=D3` and Command+Down over D3:D6 gives `=D6`), and with Shift they extend to it (`=D3:D6`); Home points to the row's first column, and Page Down a screen below.
- Typing after a reference ends it: the next arrow, where `refAt` returns `insert`, points from the edited cell again (`=D3+` and Right gives `=D3+D3`).
- F2 switches Enter mode to Edit mode, where arrows and Home move the caret, and F2 again switches back to pointing, where the rule above applies at the caret.

Return enters the formula as typed, closing what it leaves open as described in [A formula left open](#a-formula-left-open) (`=SUM(D3` becomes `=SUM(D3)`).

### Checked against Excel

Tested on the Mac keyboard in Excel 16 for Mac, en-US and ru-RU: six sets, each predicted before Excel ran it, 138 cases in each locale. Three known differences remain: a reference to a sheet the workbook does not have, and in en-US `=ROUND(1.5` and `=ROUND(A1*2`, where Excel does nothing because the function has too few arguments, which `refAt` does not check. Not tested: the mouse, and Excel for Windows.

## The text a cell opens with

`entryText(value, opts)` works in the opposite direction: it returns the text a cell opens with in edit mode, as Excel's formula bar shows the value the cell contains, so that the editor shows what Excel shows and the value parses back from it as edited in place (`parseEntry(text, { cell, inPlace: true })`): a percentage cell's `12.5%` is 0.125 again, while a negative one, `-5%`, becomes the formula `=-5%%`, as Excel's editor reads it. `needsApostrophe(text, opts)` tells whether text typed into the cell would be parsed as something other than itself.

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

The options are the cell's `numberFormat` (the one it displays with: its row's or column's if it has none of its own), the workbook's `locale` and `dateSystem`, and `now`, the moment whose year completes text parsed as a date without one.

A grid without the engine takes the text from its viewport: `getViewport` and `viewportOf` with `editText: true` return each cell's `editText`. That is its formula with `=` in the workbook's locale (`=СУММ(A1:A3)`, or `=[@Qty]*2` inside its table), otherwise `entryText` of its value in the number format it displays, with the session's clock or the result's as `now`, and empty text for a cell without a value or a formula of its own or of its legacy array formula's range.

- A number is written in full: its 15 significant digits with the locale's decimal separator and no group separators, in positional notation up to 21 characters (`100000000000000000000`), and in scientific notation beyond that (`1E+21`, `1E-20`). Currency, fraction, scientific and number formats make no difference (`$1,234.50` opens as `1234.5`, and `1 1/2` as `1.5`).
- A number with a percentage format is written as a percentage: `12%`, `12.5%`, `-5%`, and `1E+309%` for 1E+307, whose hundredfold is past the largest number.
- A number with a date or time format is written in the locale's short date (the built-in format 14: `1/31/2025`, `31.01.2025`) and long time (`6:00:00 PM`, `18:00:00`), with two spaces between them, as the formula bar writes them: the date only when its time is midnight, the time only under a format without a day when it is less than one day (`6:30:00 PM`), and the date and the time otherwise (`1/1/1900  6:00:00 PM` for 1.75 under `h:mm`). Day 0 is `1/0/1900`; a negative number or one beyond 9999-12-31, also once rounded to the second, is written as a number.
- A boolean and an error are written in the locale's words: `TRUE`, `ЛОЖЬ`, `#DIV/0!`, `#Н/Д`.
- Every cell in the range of a legacy array formula opens with the range's formula, as F2 opens it (the formula bar shows it in braces, `{=SUM(C1:C3)}`). A cell a dynamic array spilled into opens empty; when it is selected, Excel's formula bar shows its anchor's formula.
- Text is written as it is, with an apostrophe before it when typing it would produce something else (`needsApostrophe`, based on `parseEntry`): a number, a date, a time, a boolean, an error or a formula in the locale (`'007`, `'1.5` in ru-RU where it would be a date, `'=1+2`), text a percentage cell's editor would turn into a number (`'5`), text that starts with an apostrophe (`''x`), and empty text. Under the Text format nothing is parsed as anything else (`007`).

## Checked against Excel

Tested in Excel 16.113 for Mac and Excel 16.0.20430 for Windows, en-US and ru-RU: about 2,000 texts typed through the object model and 236 on the Mac keyboard, each set predicted before Excel ran it. The format a formula takes from its references was tested in Excel 16.114 for Mac and 16.0.20430 for Windows, which answered alike: 230 formulas through the object model in en-US and 223 in ru-RU, 13 on the Mac keyboard and 9 entered with Control+Return and AutoSum, each set predicted before Excel ran it and the later ones held out after the corrections ([docs/development.md](development.md#cell-entry-a-formulas-format-from-its-references)). The engine matches Excel for Windows on 1,812 of the 1,983 object-model cases and the Mac keyboard on 232 of 236; the others are the known differences below. Pasting, Control+Return, the text a cell opens with and editing in place were tested the same way; the buttons of a validation rule's alert were not.

## Known differences

- English function names and TRUE in a formula typed in ru-RU: Excel ru-RU treats them as unknown names (`=SUM(A1;1,5)` and `=TODAY()` evaluate to `#ИМЯ?` and set no number format, and `=TRUE` is a name); `parseEntry` parses them as `toInvariant` does, as the function and the boolean.
- Excel for Mac's own short date in en-US, `m/d/yy`, is narrower than Windows's `m/d/yyyy`, which the engine follows: a date entered into a selection widens the column here but not in Excel for Mac (`fill-held`).
- A formula typed or pasted over several lines (`=1+`, Control+Option+Return, `2`): Excel keeps the line break in its text, while the engine stores the canonical text without it; the value is the same (`rows/keys`, `rows/held`, `rows/quotes`).
- `#GETTING_DATA` (`#ОЖИДАНИЕ_ДАННЫХ` in ru-RU), when typed, is an error in Excel; the engine has no such error and keeps it as text.
- An unknown name typed in a different case than earlier in the workbook: Excel writes it in its first case (`=Foo*2` after `=foo` becomes `=foo*2`); `parseEntry` has no workbook and keeps the case as typed.
- A cell address used as a LET name and again after the LET in the same formula: Excel parses the later one as a defined name (`=LET(IW1,5,IW1)+IW1` stores `_xlnm.IW1` and evaluates to `#NAME?`); `parseEntry` parses it as the cell.
- Typed on the keyboard, `+5%+1` gets the percentage format of its first number, and Excel drops a `+` before a number later in a formula that starts with `+` (`+-+5` becomes `+-5`); `parseEntry` gives `+5%+1` no number format and drops only the first `+`.
- Not tested: typing on the keyboard on Windows (Excel over SSH has no screen), so the percentage entry and a leading `+` or `-` follow the Mac keyboard; the ruble sign typed on the keyboard (the layout used has none); the order of entries (Excel's AutoComplete from the text above in the column, AutoCorrect).
- Not tested for pasting: a euro sign in a form other than `€5` (decimals, digit groups, a minus sign, after the number), which a paste parses as typing does; a paste over an entire array range, past the edge of the sheet, over merged cells, or into a selection larger than the clipboard's rectangle. For Control+Return: an active cell with a number format other than General or a percentage, and a cell formatted as Text in the selection, where every cell taking the active cell's number format is our reading of the tested rule; a selection over an entire array range; and Control+Return on Windows (no screen over SSH).
- Excel's formula bar keeps an apostrophe the user typed before text that did not need one (`'abc`, `'TRUE` in ru-RU); a workbook stores no such marker, and `entryText` adds the apostrophe only before text that would otherwise be parsed as something else.
- Excel for Mac's formula bar writes a ru-RU time with a leading zero (`01:30:00`, `00:00:00`), following its own regional data; `entryText` writes the locale's long time without it (`1:30:00`), following Excel for Windows's regional data, as the engine does. Excel for Windows's formula bar was not read (it has no screen over SSH).
