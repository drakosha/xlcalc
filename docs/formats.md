# Number formats and displayed text

Besides its value, each cell of a result carries what Excel shows in it: `text`, the same as VBA's `Range.Text`, and `color`, the colour of the format section that was used. The `TEXT` function formats with the same engine.

The rules below were measured in Excel 16.113 for Mac, en-US and ru-RU.

## In short

A result gives each cell, beside its `value`, the `text` Excel shows in it (VBA's `Range.Text`) and the `color` of its format's section. The format is the cell's own `numberFormat`, else its row's, else its column's: an Excel format code as an xlsx file stores it (`#,##0.00`, `d mmmm yyyy`, `[Blue]0;[Red]-0`), shown in the book's locale. The width is the column's (`columns: { A: { width: 16 } }`, as xlsx stores it; Excel's default without one), or a merged area's columns: General drops decimals or turns scientific to fit, a number that still does not fit shows `#` across the column, text is never cut (it runs over the next cells). The value stays the number whatever the text shows, and the text is not part of the values hash.

```js
import { compute, displayValue, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const book = {
  sheets: [{
    name: 'S',
    cells: {
      A1: { value: 1234567.891, numberFormat: '#,##0.00' },
      A2: { value: -42, numberFormat: '[Blue]0;[Red]-0' },
      A3: { value: { date: '2025-01-31' }, numberFormat: 'd mmmm yyyy' },
      A4: { value: 1 / 3 },
      B1: { value: 123456789 },
    },
    columns: { A: { width: 16 }, B: { width: 6 } },
  }],
}
const cells = compute(book).sheets[0].cells
console.log(cells.A1, cells.A2.text, cells.A2.color, cells.A3.text, cells.A4.text, cells.B1.text)
// → { value: 1234567.891, text: '1,234,567.89', numberFormat: '#,##0.00' } -42 #FF0000 31 January 2025 0.333333333 1E+08
const ru = compute({ ...book, locale: 'ru-RU' }).sheets[0].cells
console.log(ru.A1.text, ru.A3.text, ru.A4.text) // → 1 234 567,89 31 Январь 2025 0,333333333
console.log(compute(book, { display: false }).sheets[0].cells.A1) // → { value: 1234567.891, numberFormat: '#,##0.00' }
console.log(displayValue(0.4567, { numberFormat: '0.0%', locale: 'ru-RU' }), displayValue(1234.5, { width: 3 })) // → { text: '45,7%' } { text: '##' }
console.log(displayValue(46096, { numberFormat: 'dd.mm.yyyy' }), displayValue(46096, { numberFormat: 'dd.mm.yyyy', width: 8.83203125 })) // → { text: '15.03.2026' } { text: '#########' }
```

`compute(book, { display: false })` and `createSession(book, { display: false })` give values alone and spend nothing on formatting. A session's changes carry the text of the cells they list, and with the display on they also list the cells whose text or colour changed without a new value: a new `numberFormat` through `setCells` or `patchCells`, a width, a row's or a column's format or a merged area ([Layout](grid.md#layout)), a column moved by `editStructure`. `displayValue(value, { numberFormat, locale, width })` shows a value without a book, for a cell being edited, say, in a loaded locale's tag or a locale object; the number itself is always the cell's `value`. Without `width` the width is not limited: General keeps its digits and no number turns into `#`, as no column narrows it. With the column's width it shows what the cell shows (Excel's default column is `width: 8.83203125`, as xlsx stores it).

A format can also be the number of one of Excel's built-in formats, from 0 to 49, as an xlsx file stores a cell's format without a code. Excel shows some of them by the locale: `numberFormat: 14` is the short date, `1/31/2025` in en-US and `31.01.2025` in ru-RU, as Excel for Windows shows it. `builtinFormatCode(id, locale?)` gives the code a number shows by (`locale` a loaded locale's tag or a locale object), for a format menu or an editor, and undefined for a number that is no built-in format, so `builtinFormatCode(id) !== undefined` tells whether it is one. The table is in [docs/formats.md](formats.md).

```js
import { builtinFormatCode } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

console.log(builtinFormatCode(14), builtinFormatCode(14, ruRU), builtinFormatCode(7, ruRU)) // → m/d/yyyy dd.mm.yyyy #,##0.00 "₽";-#,##0.00 "₽"
console.log(builtinFormatCode(3), builtinFormatCode(50), builtinFormatCode(164) !== undefined) // → #,##0 undefined false
```

The Number Format box on Excel's Home tab names the category of the active cell's format, and its list applies a format of each category. `numberFormatCategory(format, locale?)` gives that category for a code or a built-in format's number, as a stable key and the name the box shows (`Number`, `Числовой`); `numberFormatChoices(locale?)` gives the list's items in order, each with its name, the format it applies and that format's code in the locale. Both come without the engine, in `xlcalc/text` too. The box names the category of Excel's Format Cells window whose formats hold the code as Excel keeps it, and Custom when none does: `0.000` is Number, `0.00;-0.00` and `#,##0%` are Custom, and so is `$#,##0.00`, as Excel keeps a dollar written bare as a literal; the Currency item writes `"$"#,##0.00`. In ru-RU the names are Russian and the rules Excel for Windows's: `0.00` is Числовой as in en-US, money is the ruble's (`#,##0.00\ "₽"` Денежный, while the dollar's `"$"#,##0.00` is `(все форматы)`), and the short date is `dd/mm/yyyy`, which Excel shows as `ДД.ММ.ГГГГ`. Excel for Mac names nearly every code `(все форматы)` there, as it compares the code with its formats written in Russian syntax. The rules and how they were measured are in [docs/formats.md](formats.md#the-ribbons-number-format-box).

```js
import { numberFormatCategory, numberFormatChoices } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

console.log(numberFormatCategory('0.000'), numberFormatCategory('$#,##0.00').name, numberFormatCategory('"$"#,##0.00').name, numberFormatCategory(14).name)
// → { category: 'number', name: 'Number' } Custom Currency Date
console.log(numberFormatCategory('0.000', ruRU).name, numberFormatCategory('#,##0.00\\ "₽"', ruRU).name, numberFormatCategory('"$"#,##0.00', ruRU).name)
// → Числовой Денежный (все форматы)
console.log(numberFormatChoices().map((c) => c.name).join(', '))
// → General, Number, Currency, Accounting, Short Date, Long Date, Time, Percentage, Fraction, Scientific, Text
const { name, format } = numberFormatChoices(ruRU)[2]
console.log(name, format) // → Денежный #,##0.00\ "₽"
```

`TEXT` uses the same formats, read in the book's locale (`TEXT(x,"# ##0,00")` in ru-RU), without the width: it drops fills, has no colour and answers `#VALUE!` where a cell shows `#`. The rules, as measured in Excel 16.113 for Mac in en-US and ru-RU on 11,649 cells and 771 `TEXT` cases, are in [docs/formats.md](formats.md), with the differences.

Excel widens columns by itself in two places, and the engine gives the same widths. AutoFit (a double click on a column's border) gives a column the width of its widest cell: `fitColumns(book, result, sheet, columns)` computes the widths for a computed book, and the layout edit `{ sheet, columns, fit: true }` applies them in a session ([Layout](grid.md#layout)). And a number typed into a column whose width nobody set widens it until the number shows whole: `entryWidth(value, opts)` gives that width for the value a typed text computes, for a host to set with `bestFit: true`. A width Excel fitted (`ColumnInput.bestFit`) widens again as numbers are typed; one set by hand stays.

```js
import { compute, entryWidth, fitColumns, parseEntry, registerLocale } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'

registerLocale(ruRU)
const book = {
  locale: 'ru-RU',
  sheets: [{ name: 'S', cells: { A1: { value: 'Итого' }, B1: { value: 1234567.5, numberFormat: '#,##0.00' }, C1: { value: 46096, numberFormat: 14 } } }],
}
const result = compute(book)
console.log(result.sheets[0].cells.C1.text, fitColumns(book, result, 'S', 'A:C'))
// → ######### { A: 5.6640625, B: 11.33203125, C: 10.1640625 }
const typed = parseEntry('1234567890', { locale: 'ru-RU' }).input
console.log(entryWidth(typed.value, { locale: 'ru-RU' }), entryWidth(3.14159265, { locale: 'ru-RU' }), entryWidth(typed.value, { locale: 'ru-RU', width: 9 }))
// → 11.1640625 undefined undefined
```

AutoFit counts every cell of the column with something to show, hidden rows too, but not cells merged across columns; a column with nothing to show keeps its width, and a hidden column is shown. Typing widens a column for a number, a date or a formula's number that AutoFit would give more room, but not for text, which runs over the next cells, nor for a number in General that shows with fewer decimals (`3,141593`). Excel widens nothing for text set through its object model, only for text typed. The widths are Excel for Mac's, measured on 994 columns in en-US and ru-RU ([docs/formats.md](formats.md#fitting-columns)); Excel for Windows lays text out in pixels and fits it to other widths.

Rows fit the same way. `fitRows(book, result, sheet, rows)` gives the heights AutoFit gives rows `5` or `2:9` of a computed book, by spans of rows of one height, and `{ sheet, rows, fit: true }` applies them in a session: a row stands its cells on one line at its bottom, one line of each cell's font or the lines of text that wraps (`alignment.wrap`) at its column's width, and gets the most any cell needs above that line and the most any needs below, a line of Calibri 11 15 points alone and 16 wrapped; a cell merged down counts for nothing. A text with a line break typed (`setCellText`, `setCellTexts`, `fillCellText`) wraps its cell and fits its row in the same step, unless the row's height was set by hand; one pasted wraps its cell alone (a quoted field of the clipboard keeps its line breaks). A height Excel fitted is marked `bestFit` and written to xlsx as Excel writes it. The heights are Excel for Mac's, measured on 723 rows in en-US and ru-RU; Excel for Windows fits rows to its pixels ([docs/formats.md](formats.md#fitting-rows)).

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 'Total', style: 0 } } }], styles: [{ font: { size: 16 } }] })
session.setLayout([{ sheet: 'S', rows: '1', fit: true }])
session.setCellText('S', 'B2', 'two\nlines')
console.log(session.getLayout().sheets[0].rows, session.getBook().styles)
// → {
//     '1': { height: 21, bestFit: true },
//     '2': { height: 32, bestFit: true }
//   } [ { font: { size: 16 } }, { alignment: { wrap: true } } ]
```

A sheet's `conditionalFormats` are Excel's conditional formatting rules: values compared with a formula, formulas that hold, text, blanks, dates, top and average rules, duplicates, colour scales, data bars and icon sets. `Session.getCells` and `getViewport` give each cell of a window the `appearance` its rules give it by the current values, as Excel shows it: the rules' styles merged by priority, a colour scale's colour, a bar, an icon, and the text by a rule's number format.

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{
    name: 'S',
    cells: { A1: { value: 3 }, A2: { value: 12 }, A3: { value: 7 } },
    conditionalFormats: [
      { type: 'cellIs', ranges: ['A1:A3'], operator: 'greaterThan', formulas: ['10'], style: { font: { bold: true } } },
      { type: 'colorScale', ranges: ['A1:A3'], points: [{ value: { type: 'min' }, color: '#FFFFFF' }, { value: { type: 'max' }, color: '#63BE7B' }] },
      { type: 'expression', ranges: ['B1:B3'], formula: '$A1<5', style: { fill: { color: '#FFC7CE' } } },
    ],
  }],
})
const w = session.getCells('S', 'A1:B3')
console.log(w.A2.appearance, w.A3.appearance, w.B1.appearance)
// → { style: { font: { bold: true }, fill: { color: '#63BE7B' } } } { style: { fill: { color: '#BAE3C5' } } } { style: { fill: { color: '#FFC7CE' } } }
```

The rules follow Excel's, measured in Excel for Mac and Windows: a formula is written for the top-left cell of the rule's cells and moves to each, statistics leave out text and errors, a colour moves channel by channel. `readXlsx` and `writeXlsx` keep the rules, Excel 2010's among them, and structure operations move them as Excel does. Edits list no appearances: a host reads its window again where the rules' cells in `repaint` meet it. `setLayout` replaces a sheet's rules ([Layout](grid.md#layout)). [docs/conditional-formats.md](conditional-formats.md) has the rules and how they were checked.

## API

- `compute(book, { display })` and `createSession(book, { display })`. `display` is `true` by default, and then every cell in `sheets[].cells` has `text` (spilled cells too) and `color` when its section has one. Session edits return these fields for the cells they list, and also list cells whose text or colour changed while their value did not: a new `numberFormat` through `setCells`, a width, a row or column format or a merged area through `setLayout`, a moved cell through `editStructure`.
- `display: false` returns values only and spends nothing on formatting.
- `displayValue(value, { format, locale, width, dateSystem })` formats one value without a book, for example a cell being edited. It takes a value as results give it: a number, text, a boolean, `null` or `{ error }`. Without `width` the width is not limited: General keeps the digits it shows in a wide column, no number turns into `#`, and a fill (`*x`) has no column to fill and drops out, as in `TEXT`. With a width it shows what a cell of that column shows; Excel's default column is `width: 8.83203125`.
- `builtinFormatCode(id, locale)` gives the code a built-in format shows by in a locale, from the tables below (`builtinFormatCode(14, 'ru-RU')` is `dd.mm.yyyy`), and undefined for a number that is none: `builtinFormatCode(id) !== undefined` tells whether a number is a built-in format.
- `numberFormatCategory(format, locale)` gives the category of a format as the ribbon's Number Format box shows it, and `numberFormatChoices(locale)` the items of that box's list ([The ribbon's Number Format box](#the-ribbons-number-format-box)).
- `fitColumns(book, result, sheet, columns)`, the layout edit `{ sheet, columns, fit: true }` and `entryWidth(value, opts)` give the widths Excel gives columns by AutoFit and as numbers are typed ([Fitting columns](#fitting-columns)); `fitRows(book, result, sheet, rows)` and `{ sheet, rows, fit: true }` the heights of rows ([Fitting rows](#fitting-rows)).

The number always stays in `value`. The text depends on the value, the format, the locale and the column width, so it is not part of the values hash, and a hash does not change when a column gets wider.

A cell's format is its own `numberFormat`, else its row's, else its column's. A format is a code, or the number of a built-in format. Its width is its column's `width` (in xlsx units), or the width of all columns of the merged area it starts.

## Width

Excel lays out text in points of Calibri 11, the default font of xlsx files. A digit is 6 points wide. A column of xlsx width `w` is `trunc(6w + 21·6/256)` points; a column without a width is 53 points. Five points go to padding, and text fits when the widths of its characters add up to no more than the rest.

Character widths were measured by fills (`0*x` repeats `x` across the column, and the count gives the width of `x`):

| points | characters |
|---|---|
| 2 | space, `'` |
| 3 | `. , - ( ) : { } f i j l I` |
| 4 | `/ ! r s t z J г т` |
| 5 | `+ # ^ ~ < > = ? _ E F L S T Y Z`, `a c e g k v x y`, `а в е ё з к с у ч х ь э я Г Е Ё З Т` |
| 6 | digits, `$ ₽ € £ ¥` and most letters |
| 7 | `D G H N O Q U м ф ы` and similar |
| 8 | `% & w ж ш щ ю Ф Ы` |
| 9 | `m M Ж М` |
| 10 | `W Ш Щ Ю` |

Latin letters, digits and signs are in `src/numfmt/glyphs.ts`, and so are the Cyrillic letters and `₽`. They count in books of every locale, whatever locales are loaded. A character that no table has counts as a digit.

A merged area is as wide as its columns together.

## Values that do not fit

- A formatted number or date that does not fit shows `#` across the cell, one per 5 points of room, at most 255. A cell with less than 5 points of room shows nothing.
- Shown text longer than 255 characters, such as a long fill, shows `#` as well: `0* ` in a column of width 86 shows 253 characters, at width 87 it shows `#`.
- A boolean or an error shows whole when it fits, else `#` in the same way. `TRUE` needs 23 points, `ИСТИНА` 38.
- Text is never cut. It runs over the next cells, and the shown text is the text itself however narrow the column.

## General

A number in the General format shows what `TEXT(x,"General")` gives (at most 11 characters, not counting the sign) when that fits. A leading minus counts as 3 points and every other character as a digit. When it does not fit:

1. A number from 10^-4 to below 10^11 drops decimals, rounding, as long as at least one digit shows. `3.14159265` becomes `3.142`, then `3.1`, then `3`; `0.000123456789` becomes `0.0001235`, then `0.0001`.
2. Then comes scientific notation with a shorter mantissa: `1.23E+08`, `1.2E+08`, `1E+08`.
3. A number that rounds to 0 in a room of one to four digits shows `0` (`1E-10` in a one-digit column). A positive number from 10^-4 that rounds to one digit shows that digit at any width: `0.5` shows `1` even with no room.
4. The shortest scientific form shows when it fits and the number rounds to 6 digits or more, or is below 10^-4. 999999.5 shows `1E+05` in 28 points, but 99999.4 shows `#`.
5. Otherwise the cell shows `#`, at least one.

## Sections

A format has up to four sections separated by `;`. Without conditions, the first is for positive numbers (and zero, when there are fewer than three sections), the second for negative numbers, the third for zero, and the fourth, or a section with `@`, for text. A negative number in the second section writes no minus of its own.

With conditions (`[>100]`, `[<=-5]`, `[=1]`, `[<>0]`), the value takes the first of the first three sections whose condition holds or that has no condition. A value that no section takes shows General when the format has one section; otherwise it shows `#` in a cell and gives `#VALUE!` in `TEXT`.

A section writes the minus of a negative number unless it can only take negative numbers: `[Red][<0]0;0` shows −5 as `5`, while `[<=-5]"low";[>=5]"high";0` shows −1 as `-1`. A section of literals writes the minus before them: `[=1]"one";"other"` shows −2 as `-other`. An empty section shows nothing (`0;;`). Booleans and errors show as they are; under a text section, a boolean shows as that section's text (`0;-0;0;"is "@` shows FALSE as `is FALSE`).

## Numbers

- `0` is a digit or zero, `#` a digit, `?` a digit or a space. A group separator between digit placeholders groups thousands; after them it divides by 1000 (`0.0,,`). `%` multiplies by 100. Both round to 15 significant digits first. A second decimal point is a literal.
- Scientific notation shows as many integer digits as there are integer placeholders, and the exponent is a multiple of their count: `##0.0E+0` shows 0.5 as `500.0E-3`. `E+` always writes the exponent's sign, `E-` only a minus, and the exponent is padded to its placeholders (`1.00E+300`). A lower-case exponent (`0.00e+00`) is `#VALUE!` in `TEXT`, and Excel will not load an xlsx file that has one.
- Fractions put the integer part before the numerator when the format has one, and take the closest fraction for the denominator's placeholders or a fixed denominator (`?/8`). A numerator above 2,147,483,646 shows `#` (`#VALUE!` in `TEXT`): `?/?` shows 1E+15 as `#`. With an integer part, a negative number keeps its minus and a whole 0 shows `0`.
- Positional numbers keep 15 significant digits and pad with zeros: `TEXT(1E+250,"0")` is a 1 followed by 250 zeros.
- Literals are `"…"`, `\x`, and characters that need no quotes (`$ - + / ( ) :` and space). `_x` is one space, whatever `x` is (Excel leaves the width of `x`, but `Range.Text` has a space). `*x` repeats `x` across the room the rest of the text leaves; only the last fill counts. `[$₽-419]` writes the symbol `₽` (any symbol works) with the locale 419.

## Dates and times

- Excel's 1900 calendar: serial 0 is `1/0/1900`, 60 is `2/29/1900`. A negative serial, or one from 2,958,466 (past 9999-12-31), shows `#` and is `#VALUE!` in `TEXT`.
- In a book of the 1904 date system, serial 0 is `1/1/1904` and the last date is serial 2,957,003. A negative date or time shows as its size with a minus: −1 with `m/d/yyyy` shows `-1/2/1904`, −0.5 with `h:mm` shows `-12:00`, and `TEXT` gives the same.
- Time is rounded to what the format shows: whole seconds, or the decimals of `ss.0` to `ss.000`. 45688.999999 is `2025-02-01 00:00:00` with `yyyy-mm-dd hh:mm:ss` and `23:59:59.9` with `h:mm:ss.0`. Whole seconds are taken as `SECOND` takes them: half a second added and the sum split in floating point, so an exact half second of a date of today stays below (45688+1/256, 00:05:37.5, shows `00:05:37`).
- `m` means minutes after hours or before seconds, and the month elsewhere. `mmmmm` is the month's first letter. `ddd`, `dddd`, `mmm` and `mmmm` are the locale's names. `A/P` writes the letter in the format's case (`a/p` shows `6:00 p`).
- `[h]`, `[mm]` and `[ss]` count elapsed hours, minutes or seconds from 0. A negative value shows with a minus when the section has elapsed time alone (`[h]` shows −1 as `-24`), and `#` otherwise.
- In a cell, a bare `/` in a date format is the locale's date separator: `dd/mm/yyyy` shows `31.01.2025` in ru-RU. In `TEXT` it stays `/`. Quoted and escaped separators stay as written.
- `[$-409]` and `[$-419]` select English and Russian names. `[$-F800]` is the system's long date (en-US `dddd, mmmm d, yyyy`; ru-RU `пятница, 31 января 2025 г.`, with months in the genitive) and `[$-F400]` the long time. Other tags are read and ignored.

## Built-in formats

An xlsx file stores some formats by number alone, without a code. A format here can be such a number, from 0 to 49: `{ value: 45688.75, numberFormat: 14 }`. Excel shows some of them by the locale, so the same number shows differently in en-US and ru-RU.

| number | en-US | ru-RU |
|---|---|---|
| 0, 23 to 26 | `General` | `General` |
| 1 to 4 | `0`, `0.00`, `#,##0`, `#,##0.00` | the same |
| 5 to 8 | `$#,##0_);($#,##0)` and the like | `#,##0 "₽";-#,##0 "₽"` and the like |
| 9 to 13 | `0%`, `0.00%`, `0.00E+00`, `# ?/?`, `# ??/??` | the same |
| 14, 27 to 31, 36 | `m/d/yyyy` | `dd.mm.yyyy` |
| 15, 16, 17 | `d-mmm-yy`, `d-mmm`, `mmm-yy` | `dd.mmm.yy`, `dd.mmm`, `mmm.yy` |
| 18, 19 | `h:mm AM/PM`, `h:mm:ss AM/PM` | the same |
| 20, 21, 32 to 35 | `h:mm`, `h:mm:ss` | the same |
| 22 | `m/d/yyyy h:mm` | `dd.mm.yyyy h:mm` |
| 37 to 40 | `#,##0_);(#,##0)` and the like | `#,##0;-#,##0` and the like |
| 41 to 44 | the accounting formats with `$` | the accounting formats with `₽` after the amount |
| 45 to 49 | `mm:ss`, `[h]:mm:ss`, `mm:ss.0`, `##0.0E+0`, `@` | the same |

In en-US, 14 shows 2025-01-31 as `1/31/2025` and 38 shows −1234.5 as `(1,235)` in red. In ru-RU they show `31.01.2025` and `-1 235`, and 20 shows 6 in the morning as `6:00`. A result keeps the number in `numberFormat`. `src/numfmt/builtin.ts` has the en-US codes, ru-RU's locale file (`src/locales/ru-RU.ts`) the ru-RU ones, and `builtinFormatCode` gives them; a value shows by the code as by the number in both locales (`test/builtin-format-code.test.ts`).

The codes are those Excel for Windows shows. Excel for Mac shows the short date of en-US as `m/d/yy` (`1/31/25`) and the times of ru-RU with a leading zero (`06:00`); the engine follows Windows there.

A format written as a code always shows as written. `numberFormat: 'm/d/yy'` shows `1/31/25` in ru-RU too.

## Locales

A cell's format is read the way xlsx stores it, in en-US syntax, whatever the book's locale. It is shown with the locale's separators, names, booleans, error texts and date separator: in ru-RU `#,##0.00` shows `1 234 567,89`, `mmmm` shows `Январь`, TRUE shows `ИСТИНА`, `#N/A` shows `#Н/Д`.

`TEXT` reads the locale's own codes (`# ##0,00`, `ДД.ММ.ГГГГ`, `Основной`) and colour names (`[Красный]`, `[Синий]`, `[Цвет10]`); in ru-RU, `[Red]` and `[Color10]` are `#VALUE!`. Month names in ru-RU are in the nominative with a capital (`31 Январь 2025`), and short ones have three letters (`фев`, `сен`, `ноя`), as Excel for Windows writes them; the long date takes the genitive (`31 января 2025 г.`). Excel for Mac writes them in lower case and its short forms are `февр`, `март`, `июнь`, `июль`, `сент`, `нояб`.

## Colours

Format colours are the eight names `[Black]`, `[White]`, `[Red]`, `[Green]`, `[Blue]`, `[Yellow]`, `[Magenta]`, `[Cyan]` in any case, and `[ColorN]`, an index into Excel's default palette of 56 colours (xlsx's indexed colours 8 to 63). `[Red]` gives `#FF0000`, `[Color10]` gives `#008000`. Black is the automatic colour and is left out of results. The names and the palette come from Microsoft's "Review guidelines for customizing a number format"; the 13 colours measured in Excel agree with it.

## How `TEXT` differs from a cell

| | `TEXT` | a cell |
|---|---|---|
| format codes | the book's locale (`# ##0,00` in ru-RU) | as xlsx stores them (`#,##0.00`) |
| General | at most 11 characters | narrowed to the column |
| a number wider than the column | no width | `#` across the cell |
| a value its section cannot show | `#VALUE!` | `#` across the cell |
| `*x` | dropped | `x` across the room |
| `_x` | a space | a space |
| colour | none | the section's |
| `/` in a date | `/` | the locale's date separator |

## Fitting columns

Excel widens columns by itself in two places, and a host that shows a book as Excel does needs the same widths: AutoFit, a double click on a column's border, and a number typed into a column too narrow for it. The rules below were measured in Excel 16.113 for Mac.

```js
import { compute, entryWidth, fitColumns } from 'xlcalc'

const book = { sheets: [{ name: 'S', cells: { A1: { value: 'abc' }, B1: { value: 46096, numberFormat: 'd mmmm yyyy' }, C1: { value: 'Итого' } } }] }
console.log(fitColumns(book, compute(book), 'S', 'A:C')) // → { A: 3.83203125, B: 12.6640625, C: 5.6640625 }
console.log(entryWidth(1234567890), entryWidth(46096, { numberFormat: 'd mmmm yyyy', width: 20 })) // → 11.1640625 undefined
```

### Column geometry for a grid

`columnWidthToPoints(width?)` and `columnWidthFromPoints(points)` from `xlcalc/text` let a grid use the geometry that the library uses for shown text, `####` and overflow. They use the same helper as the formatter and bring in no engine.

```js
import { columnWidthToPoints, columnWidthFromPoints } from 'xlcalc/text'

console.log(columnWidthToPoints(8.83203125), columnWidthToPoints(), columnWidthFromPoints(53), columnWidthFromPoints(53.75))
// → 53 53 8.83203125 8.83203125
```

The input is the stored xlsx width, not the object model's `ColumnWidth` characters. Missing or null width means 53 points. The inverse rounds points down to a whole point and stores the width in 256ths of a digit: 53.75 points become 53 points and width 8.83203125. The helpers accept the full measured file range, through width 255.83203125 or 1,535 points; `ColumnInput.width` still accepts at most 255, so cap a wider returned width at 255 before setting a book's column.

Use points consistently. At 96 CSS pixels per inch, a host converts points to CSS pixels by multiplying by 4/3 and pixels to points by 3/4 before calling the inverse. This is a unit conversion, not a measurement of screen pixels, zoom or display scaling. The measured object-model widths use [Range.Width in points](https://learn.microsoft.com/en-us/office/vba/api/excel.range.width); [ColumnWidth](https://learn.microsoft.com/en-us/office/vba/api/excel.range.columnwidth) uses the Normal style's digit.

`columnWidthToChars(width?)` and `columnWidthFromChars(chars)` give the width Excel for Mac's Column Width box shows and takes, in characters of the default font to two decimals: 6 points a character past a margin of 5, 11 points a character under one (`columnWidthToChars()` is 8, `columnWidthFromChars(8)` is 8.83203125). They agree with every column the Mac's fit sets read (`ColumnWidth` beside the stored width); Excel for Windows counts 7 pixels a character and shows the default as 8.43. `DEFAULT_ROW_HEIGHT` (15), `MAX_ROW_HEIGHT` (409.5) and `MAX_OUTLINE` (7 levels) are the sizes of rows and outlines the library uses.

Windows has another geometry. Its default is `StandardWidth` 8.43, 48 points (64 pixels at 96 dpi), saved as width 9.140625 when made explicit. An explicit file width 8.83203125 is 46.5 points there (62 pixels), while it is 53 points in the Mac model. For the measured file widths Windows gives `trunc((256w + trunc(128/7)) / 256 * 7) * 3/4` points. That formula does not convert Mac-model defaults into Windows defaults. The library keeps its Mac geometry; switching it to Windows would change which values fit.

### AutoFit

`fitColumns(book, result, sheet, columns)` gives the widths for a computed book, and the layout edit `{ sheet, columns: 'B:D', fit: true }` sets them in a session (`setLayout` returns the cells that show otherwise, `getLayout()` the widths). A column gets the width of its widest cell, each value by its format in the book's locale, laid out as Excel lays it out for AutoFit, which is not quite how it shows it:

- Text of ASCII characters alone is measured by the widths of the table above. Of `w` points of text, the column gets `w + 7` points, two more than the padding needs; from about 90 points of text Excel counts a thirtieth more of it, and the column gets `floor(w · 31/30) + 5`. `abc` (16 points) gets 23, a text of 300 points 315.
- Text with any other character (Cyrillic, `₽`, `é`) is measured, all its characters, by the font's advance widths, which `src/numfmt/advances.ts` takes from Carlito, a font metrically compatible with Calibri (in 1/2048 of 11 points): of `r` points, the column gets the larger of `round(r) + 7` and `floor(r · 31/30) + 5`. `Итого` is 26.7 points and gets 34.
- A padding `_)` is as wide as `)` and a `?` without a digit as a digit, though the shown text has spaces there; a fill takes no room. `$#,##0.00_)` gives 1234.5 58 points, accounting's `"-"??` 28.
- A number in General counts each character as a digit and a leading minus as 3 points, two points past the padding, or none in scientific form: 1234567890 gets 67 points, 123456789012 (`1.23457E+11`) 71.
- Booleans and errors are text: `TRUE` gets 30 points, `ИСТИНА` 46.

The xlsx width of a column of `p` points is `p/6` in 256ths: 67 points are 11.1640625. A column fitted to the default 53 points gets no width of its own (null), as Excel saves it; the widest is 255 (Excel's widest, 255.83, reads as 255).

Every cell of the column counts, spilled ones and those of hidden rows too; a cell merged with the columns beside it does not, one merged down does. A column with nothing to show (blanks, empty text, `;;;`, a negative date) keeps its width. AutoFit marks the width as Excel's own (`bestFit: true`) and shows a hidden column.

### Typing a number

`entryWidth(value, { format, locale, dateSystem, width, bestFit })` gives the width Excel gives a column as `value` is typed into one of its cells, or undefined when it keeps its width; a host sets it with `bestFit: true`.

- A number, a date or a formula's number widens a column of the default width or one Excel fitted (`bestFit`) to the width AutoFit gives it, when that is more: `15.03.2026` in ru-RU needs 54 points of the default 48 and widens the column to 61. A number that shows already widens it too when AutoFit gives it more: `15-Mar-26` fills 48 points exactly and gets 55.
- A number in General widens the column when it would show in scientific form or whole (`1234567890` to 67 points, `12345678` to 55), but not when it shows with fewer decimals: `3.14159265358979` shows `3.141593` and the column stays.
- Text, booleans and errors widen nothing (text runs over the next cells), nor does a width set by hand, and a column never narrows. A width assigned with `setLayout([{ sheet: 'S', columns: 'A', set: { width: 3 } }])` is one set by hand: it clears the column's `bestFit`, so a number typed later leaves the column as it is (an explicit `bestFit: true` keeps it fitted). A fitted width equal to the default stays a width (8.83203125 with `bestFit`).
- Only typing widens a column: text set through the object model (`Value`, `Formula2Local`) widens nothing, in Excel for Mac and Excel for Windows. So a host calls `entryWidth` for text a user typed, with the value it computed (from `setCells`) and the cell's format after the entry; `setCellText` does this itself, in the same step of the undo ([Editing a grid](edits.md#typing-text)).

## Fitting rows

Excel sets a row's height by itself as well: AutoFit (a double click on the border below a row's number) and text that wraps, typed into the row. Excel for Mac keeps every row whose height nobody set at that height, on opening a file too; Excel for Windows the same, at its own heights. The rules below were measured in Excel 16.113 for Mac.

```js
import { compute, fitRows } from 'xlcalc'

const book = {
  sheets: [{ name: 'S', cells: { A1: { value: 'Total', style: 0 }, A2: { value: 'two\nlines', style: 1 } } }],
  styles: [{ font: { size: 16 } }, { alignment: { wrap: true } }],
}
console.log(fitRows(book, compute(book), 'S', '1:3')) // → { '1': 21, '2': 32, '3': null }
```

`fitRows(book, result, sheet, rows)` gives the heights for a computed book, by spans of rows of one height, and the layout edit `{ sheet, rows: '2:5', fit: true }` sets them in a session, each as Excel's own (`bestFit: true`), the default height as none, a hidden row shown. Excel stands the cells of a row on one line at its bottom, and the row gets the most any cell needs above that line plus the most any cell needs below it:

- A cell holding a number, a boolean, an error, text that does not wrap (its line breaks too) or nothing is one line of its font. A line of Calibri of `s` points is `round(1.27·s + 0.96)` points high, `s` of them above the line and the rest below: 15 for 11 (11 and 4), 21 for 16 (16 and 5), 31 for 24, 92 for 72. A size of a half point counts as the next whole one.
- Text that wraps (`alignment.wrap`) takes its lines, each line a point higher than a line alone, for the first 20 lines: Calibri 11 is 16 points a line, 320 for 20 lines and 335 for 21. Its last line stands on the row's line, so it needs below what one line of its font does. A line breaks at each line break (one at the end of the text makes an empty line) and where the words no longer fit the column's width less 7 points, two more than the padding of shown text: the default column of 53 points holds 46 points of text, `b `, a space and 38 points more. Spaces at the end of a line hang past it; a word wider than the room breaks between its characters. A character is the font's advance rounded to a point at its size, as Excel shows Calibri 11 ([Width](#width)): `b` is 6 points at 11, 8 at 16, 13 at 24.
- Every cell counts, blank ones by the font their row's style, else their column's style, else the book's default gives them. So a row is never lower than a line of the default font unless its own style is smaller (a row of Calibri 8 is 11 points), a column of a larger font raises every row, and small wrapped text stands a point higher than its lines, the default font needing more below.
- Because cells share the line, a row can be higher than its highest cell: two wrapped lines of Calibri 11 (32 points) beside a line of 24 points (31) make a row of 35.
- A cell merged down does not count. One merged across counts as a line of its font, or, when it wraps, only below the line. A hidden column counts with its width.
- The highest row is 409.5 points.

Text with a line break typed into a cell wraps it, on the keyboard (Control+Option+Return on the Mac) and through the object model alike, and its row fits it unless its height was set by hand: `a`, a line break and `b` in a row of the default height give it 32 points. A row of a cell that wraps fits again as text is typed there, shorter or longer. Pasted, a field with a line break inside quotes wraps its cell, and the row keeps its height. A formula typed over lines wraps nothing. `setCellText`, `setCellTexts` and `fillCellText` do this in the step of the entry: the cell's style gains `alignment.wrap` (an equal style of the table is found, else one is added) and the row its height with `bestFit`; a row set by hand keeps its height, and so does a hidden one. `pasteText` wraps the cells and fits no row.

Assigning a height with `setLayout([{ sheet: 'S', rows: '2', set: { height: 40 } }])` clears the row's previous `bestFit` mark: later typing keeps the 40 points, even when it needs more room. This applies when assigning the same height too. Removing a height (`height: null`) also removes the mark and returns to the default height, which can fit on entry again. These are exceptions to a line patch preserving fields it leaves out. An explicit `bestFit: true` keeps a height fitted; a patch without `height` preserves the mark. AutoFit (`fit: true`) enables automatic fitting again. Undo and redo restore the height and its mark together; a manually assigned height is written to xlsx as `customHeight`.

## The ribbon's Number Format box

The Number Format box on Excel's Home tab names the category of the active cell's format, and its list applies a format of each category. A host that draws Excel's ribbon needs both: `numberFormatCategory(format, locale)` gives the category as a key (`number`, `currency`, `custom`) and the name the box shows (`Number`, `Числовой`), and `numberFormatChoices(locale)` gives the list's eleven items in order, each with its key, its name, the format it applies and that format's code in the locale. Both take a code or a built-in format's number, and come in `xlcalc/text`, without the engine. The rules below were measured in Excel 16.113 for Mac in en-US and in Excel 16.0 for Windows in ru-RU.

```js
import { numberFormatCategory, numberFormatChoices } from 'xlcalc/text'

console.log(numberFormatCategory('0.000'), numberFormatCategory('0.00;-0.00').name)
// → { category: 'number', name: 'Number' } Custom
console.log(numberFormatChoices()[2])
// → {
//     key: 'currency',
//     name: 'Currency',
//     format: '"$"#,##0.00',
//     code: '"$"#,##0.00'
//   }
```

The box names the category of Excel's Format Cells window whose formats hold the code, and Custom when none does. The keys are `general`, `number`, `currency`, `accounting`, `date`, `time`, `percentage`, `fraction`, `scientific`, `text`, `special` and `custom`. In en-US:

- The code is compared as Excel keeps it, a literal written bare kept escaped. So `$#,##0.00` is Custom: Excel keeps it as `\$#,##0.00`, and Currency's dollar is quoted, `"$"#,##0.00`. A parenthesis or a space counts the same bare or escaped.
- Number is digits with up to 30 decimals, with groups or without (`0`, `0.000`, `#,##0.0`), and Excel's negatives: red without a minus (`0.00;[Red]0.00`), in parentheses (`0.00_);(0.00)`) and red in parentheses. A minus of the code's own (`0.00;-0.00`), a red one, another colour, literals and `#.##` make it Custom.
- Percentage is `0%` with up to 30 decimals, Scientific `0.0E+00` with 1 to 30 (`0E+00` is Custom), Fraction the nine of Format Cells: `# ?/?`, `# ??/??`, `# ???/???`, and halves, quarters, eighths, sixteenths, tenths and hundredths (`# ?/2` to `# ??/100`).
- Currency is the quoted dollar before grouped digits, or a currency in brackets (`[$€-407]`, `[$USD]`) before them, with a space or without, or after them with a space; negatives red with a minus, or in parentheses. Accounting is Excel's form `_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)` with the dollar, a bracketed currency or none, and with a bracketed currency also Excel for Mac's forms with minus signs (`_-* #,##0.00 [$€-407]_-;…`).
- A date is the short date (built-in 14: `m/d/yyyy` as the package shows it, or the Mac's own `m/d/yy`), the system's long date `[$-F800]dddd, mmmm dd, yyyy` exactly, or a date followed by `;@`, as Format Cells writes its other dates (`m/d/yyyy;@`, `[$-409]mmmm d, yyyy;@`, `dd.mm.yyyy;@`). A time is the system's long time `[$-F400]h:mm:ss AM/PM` exactly, or a time followed by `;@` (`h:mm;@`). A date or time without `;@` is Custom (`h:mm`, `d-mmm-yy`, `yyyy-mm-dd`), and so are built-in 15 to 22 and 45 to 48.
- Text is `@` alone, General `General` alone. Special is `00000`, `00000-0000`, `000-00-0000`, the phone `[<=9999999]###-####;(###) ###-####` and `000000`.
- A built-in format by number is its code as Excel keeps it: 5 to 8 and 41 to 44 have the dollar quoted, so 5 is Currency, while the code the package writes for it, `$#,##0_);($#,##0)`, is Custom, as Excel calls that code.
- The value does not count: a format shows its category on a number, text, a date or a blank.

In ru-RU the box follows Excel for Windows. It writes the code in Russian syntax (`0.00` as `0,00`, `#,##0` as `# ##0`, `[Red]` as `[Красный]`) and compares it with the formats of its Format Cells window written so. What that gives:

- Numbers, percents, exponents and text are en-US's: `0.00`, `#,##0.000`, `0;[Red]0`, `0.0%`, `0.000E+00` and `@` name the same category as in en-US. Negatives in parentheses (`0.00_);(0.00)`) and with a minus (`0.00;[Red]-0.00`) are `(все форматы)`.
- Money is the ruble's, quoted after the amount with a space: `#,##0.00 "₽"` and `#,##0.00\ "₽"` are Денежный, with negatives with a minus, red or not (`#,##0.00 "₽";[Red]-#,##0.00 "₽"`). Accounting is ru-RU's form with minus signs, with the ruble or without a symbol (`_-* #,##0.00 "₽"_-;-* #,##0.00 "₽"_-;_-* "-"?? "₽"_-;_-@_-`). The ruble written bare (`#,##0.00 ₽`), `"р."` and the dollar's forms (`"$"#,##0.00`, `_("$"* #,##0.00_);…`) are `(все форматы)`. Bracketed currencies, their accounting forms and the special formats are one table for every locale, and match here as in en-US, a bracketed currency with a minus too (`#,##0.00 [$₽-419];-#,##0.00 [$₽-419]`).
- A fraction has its space quoted, as Excel writes built-in 12: `#" "?/?`, `#" "?/4` and `#" "???/???` are Дробный, `# ?/?` is `(все форматы)`.
- The short date is `dd/mm/yyyy`, ru-RU's `ДД.ММ.ГГГГ` in en-US syntax. `m/d/yyyy` written as a code is `(все форматы)`, as Excel shows it `М.Д.ГГГГ`, and so is `dd.mm.yyyy`, its dots literal. The long date and time and the dates and times with `;@` are as in en-US.
- A built-in format by number has the category of ru-RU's own code for it: 5 to 8 are Денежный, 41 to 44 Финансовый, 37 to 40 `(все форматы)`, 14 and 27 to 31 and 36 Дата; the others as in en-US. A code equal to a built-in format's en-US code is not that format here: written as a code of the file's own, `# ?/?` and `m/d/yyyy` are `(все форматы)`.

Excel for Mac in ru-RU does otherwise: it compares the code, kept in en-US syntax, with its formats written in Russian syntax, so only a code that syntax leaves as it is matches (`0`, `0%`), and `0.00`, `#,##0`, the fractions and every date and time are `(все форматы)` there.

The list's items in en-US are General (0), Number (2), Currency (`"$"#,##0.00`), Accounting (44), Short Date (14), Long Date (`[$-F800]dddd\,\ mmmm\ dd\,\ yyyy`), Time (`[$-F400]h:mm:ss\ AM/PM`), Percentage (10), Fraction (12), Scientific (11) and Text (49); in ru-RU the same, by their Russian names, with Currency `#,##0.00\ "₽"`. An item applies the same format whatever the cell held: Number gives `0.000` the code `0.00`, Currency gives a euro amount the dollar. The codes are written as Excel writes them to a file. Another locale gives its names and its Currency item's code in `regional.formats` ([docs/locale.md](locale.md)), en-US's without it.

## How the engine does it

`src/numfmt/` holds the engine, assembled in `src/numfmt.ts`. `parse.ts` reads a format into sections (colour, condition, locale tag, tokens) and caches it per dialect and format, so a column with one format is parsed once. `select.ts` picks the section. `number.ts`, `date.ts` and `general.ts` format numbers, dates and General. `glyphs.ts` has the character widths (Cyrillic too), `palette.ts` the colours, `style.ts` what a locale writes. `formatWith` is `TEXT` and `displayWith` a cell.

`src/shown.ts` finds a cell's format and width. `src/layout/fit.ts` fits columns and widens them on entry, by `fitPoints` and `typedPoints` of `src/numfmt.ts`, which lay a value out with `Style.layout` (paddings and blanks marked) and measure it by `fitText` (`glyphs.ts`, `advances.ts`). Sessions compare what cells show before and after an edit that recomputes nothing (`src/session/looks.ts`) and after a structure operation (`src/session/moved.ts`).

Most cells are text or General numbers in a default column, and their text is the value's shortest decimal form, so they cost almost nothing. On the synthetic book (54,778 cells) a cold computation takes 310 ms with display and 279 ms without, in the closures mode.

## Checked against Excel

Measured in Excel 16.113 for Mac, en-US and ru-RU, and checked again in Excel for Windows: the shown text of 12,949 cells, column widths, AutoFit and row heights, the later sets predicted before Excel ran them.

## Known differences

- **Fonts.** Widths are those of Calibri 11 in Excel for Mac. A cell's own font (another font or size, bold), indent, wrapping, shrinking and rotation are ignored, and characters outside the measured table count as a digit. Excel for Windows lays out text in pixels, with a digit 7 pixels wide at 96 dpi, so near the edge of a column it can show `#` where the engine shows the number, or the other way round.
- **Fitting rows.** Heights are Excel for Mac's for Calibri; a cell in another font is fitted as Calibri of its size (Arial 10 wraps two lines in 29 points in Excel and 30 here, Arial 16 is 20 there and 21 here), and Excel for Windows fits to its pixels (above). Excel keeps a row whose height nobody set at its fitted height whenever the row changes and when a file opens; the engine fits rows on AutoFit and as text that wraps is typed, not when a file opens or a cell's font changes. A formula typed over lines keeps its line breaks in Excel's text and not in the engine's.
- **Locale tags.** Only `[$-409]`, `[$-419]`, `[$-F800]` and `[$-F400]` change names or formats. Other locale codes, calendars and numeral systems (`[$-x-sysdate]`, `[DBNum1]`, `B2`) are read and ignored.
- **Fitting columns.** Widths are Excel for Mac's; Excel for Windows fits to its pixels (above). A cell's own font is not taken into account (bold text that the engine fits to 115 points gets 117 in Excel), nor characters outside the tables (`中文字符` counts as four digits). Text outside ASCII is measured by the advance widths of Carlito, metrically Calibri's, which give Excel's width on short text but a point off on some long runs (19 of 70 runs of one Cyrillic letter 600 points wide, `Итого по разделу 1`, ru-RU's long date). Where Excel for Mac shows its own short date (`3/15/26`) or ru-RU month names (`март`), the engine fits the text Excel for Windows shows (`3/15/2026`, `Март`). The widest column is 255 here and 255.83 in Excel.
- **The ribbon's Number Format box.** The rules are Excel for Mac's in en-US and Excel for Windows's in ru-RU; Excel for Windows in en-US was not measured. Excel for Windows names a date or time with `;@` by the lists of its Format Cells window, which the package does not hold: in ru-RU `dd/mm/yy;@`, `d/m;@`, `[$-419]d mmm;@` and en-US's `m/d/yyyy;@` are Дата, but `dd/mm;@`, `d mmm;@` and `[$-419]d mmmm yyyy;@` none, and `dd/mm/yy h:mm;@` is Время. The functions name every such code by its letters (16 of the 50 held-out cells otherwise), as codes with `;@` come from such lists. Excel for Mac in ru-RU names nearly every code `(все форматы)` (above). The code `m/d/yyyy` is the short date in en-US here, as built-in 14 shows by it (Excel for Windows's code); Excel for Mac, whose short date is `m/d/yy`, calls it Custom. Excel for Mac names the category of a format its list applied by the list's state as well: in ru-RU an item applied showed its own category (`Числовой`) in some runs and `(все форматы)` in others. Its Short Date, Long Date and Time items were not the same in two sessions of one day: one offered built-in 14, the long date and the long time, which `numberFormatChoices` gives, the other `yyyy-mm-dd;@`, `d/m;@` and `h:mm;@`.
- **A second decimal separator** among the decimals of a format (`#,##0,00` in ru-RU's `TEXT`) is a literal between digits; Excel shows one digit fewer before it for some numbers.
