# Number formats and shown text

Besides its value, each cell of a result carries what Excel shows in it: `text`, the same as VBA's `Range.Text`, and `color`, the colour of the format section that was used. The `TEXT` function formats with the same engine.

The rules below were measured in Excel 16.113 for Mac, en-US and ru-RU.

## API

- `compute(book, { display })` and `openSession(book, { display })`. `display` is `true` by default, and then every cell in `sheets[].cells` has `text` (spilled cells too) and `color` when its section has one. Session edits return these fields for the cells they list, and also list cells whose text or colour changed while their value did not: a new `format` through `setCells`, a width, a row or column format or a merged area through `setLayout`, a moved cell through `restructure`.
- `display: false` returns values only and spends nothing on formatting.
- `displayValue(value, { format, locale, width, dateSystem })` formats one value without a book, for example a cell being edited. It takes a value as results give it: a number, text, a boolean, `null` or `{ error }`. Without `width` the width is not limited: General keeps the digits it shows in a wide column, no number turns into `#`, and a fill (`*x`) has no column to fill and drops out, as in `TEXT`. With a width it shows what a cell of that column shows; Excel's default column is `width: 8.83203125`.
- `builtinFormatCode(id, locale)` gives the code a built-in format shows by in a locale, from the tables below (`builtinFormatCode(14, 'ru-RU')` is `dd.mm.yyyy`), and undefined for a number that is none: `builtinFormatCode(id) !== undefined` tells whether a number is a built-in format.
- `fitColumns(book, result, sheet, columns)`, the layout edit `{ sheet, columns, fit: true }` and `entryWidth(value, opts)` give the widths Excel gives columns by AutoFit and as numbers are typed ([Fitting columns](#fitting-columns)).

The number always stays in `value`. The text depends on the value, the format, the locale and the column width, so it is not part of the values hash, and a hash does not change when a column gets wider.

A cell's format is its own `format`, else its row's, else its column's. A format is a code, or the number of a built-in format. Its width is its column's `width` (in xlsx units), or the width of all columns of the merged area it starts.

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
- Time is rounded to what the format shows: whole seconds, or the decimals of `ss.0` to `ss.000`. 45688.999999 is `2025-02-01 00:00:00` with `yyyy-mm-dd hh:mm:ss` and `23:59:59.9` with `h:mm:ss.0`.
- `m` means minutes after hours or before seconds, and the month elsewhere. `mmmmm` is the month's first letter. `ddd`, `dddd`, `mmm` and `mmmm` are the locale's names. `A/P` writes the letter in the format's case (`a/p` shows `6:00 p`).
- `[h]`, `[mm]` and `[ss]` count elapsed hours, minutes or seconds from 0. A negative value shows with a minus when the section has elapsed time alone (`[h]` shows −1 as `-24`), and `#` otherwise.
- In a cell, a bare `/` in a date format is the locale's date separator: `dd/mm/yyyy` shows `31.01.2025` in ru-RU. In `TEXT` it stays `/`. Quoted and escaped separators stay as written.
- `[$-409]` and `[$-419]` select English and Russian names. `[$-F800]` is the system's long date (en-US `dddd, mmmm d, yyyy`; ru-RU `пятница, 31 января 2025 г.`, with months in the genitive) and `[$-F400]` the long time. Other tags are read and ignored.

## Built-in formats

An xlsx file stores some formats by number alone, without a code. A format here can be such a number, from 0 to 49: `{ value: 45688.75, format: 14 }`. Excel shows some of them by the locale, so the same number shows differently in en-US and ru-RU.

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

In en-US, 14 shows 2025-01-31 as `1/31/2025` and 38 shows −1234.5 as `(1,235)` in red. In ru-RU they show `31.01.2025` and `-1 235`, and 20 shows 6 in the morning as `6:00`. A result keeps the number in `format`. `src/numfmt/builtin.ts` has the en-US codes, ru-RU's locale file (`src/locales/ru-RU.ts`) the ru-RU ones, and `builtinFormatCode` gives them; a value shows by the code as by the number in both locales (`test/builtin-format-code.test.ts`).

The codes are those Excel for Windows shows. Excel for Mac shows the short date of en-US as `m/d/yy` (`1/31/25`) and the times of ru-RU with a leading zero (`06:00`); the engine follows Windows there.

A format written as a code always shows as written. `format: 'm/d/yy'` shows `1/31/25` in ru-RU too.

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

const book = { sheets: [{ name: 'S', cells: { A1: { value: 'abc' }, B1: { value: 46096, format: 'd mmmm yyyy' }, C1: { value: 'Итого' } } }] }
console.log(fitColumns(book, compute(book), 'S', 'A:C')) // → { A: 3.83203125, B: 12.6640625, C: 5.6640625 }
console.log(entryWidth(1234567890), entryWidth(46096, { format: 'd mmmm yyyy', width: 20 })) // → 11.1640625 undefined
```

### AutoFit

`fitColumns(book, result, sheet, columns)` gives the widths for a computed book, and the layout edit `{ sheet, columns: 'B:D', fit: true }` sets them in a session (`setLayout` returns the cells that show otherwise, `layout()` the widths). A column gets the width of its widest cell, each value by its format in the book's locale, laid out as Excel lays it out for AutoFit, which is not quite how it shows it:

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
- Text, booleans and errors widen nothing (text runs over the next cells), nor does a width set by hand, and a column never narrows. A fitted width equal to the default stays a width (8.83203125 with `bestFit`).
- Only typing widens a column: text set through the object model (`Value`, `Formula2Local`) widens nothing, in Excel for Mac and Excel for Windows. So a host calls `entryWidth` for text a user typed, with the value it computed (from `setCells`) and the cell's format after the entry; `setCellText` does this itself, in the same step of the undo ([Editing a grid](edits.md#typing-text)).

## How the engine does it

`src/numfmt/` holds the engine, assembled in `src/numfmt.ts`. `parse.ts` reads a format into sections (colour, condition, locale tag, tokens) and caches it per dialect and format, so a column with one format is parsed once. `select.ts` picks the section. `number.ts`, `date.ts` and `general.ts` format numbers, dates and General. `glyphs.ts` has the character widths (Cyrillic too), `palette.ts` the colours, `style.ts` what a locale writes. `formatWith` is `TEXT` and `displayWith` a cell.

`src/shown.ts` finds a cell's format and width. `src/layout/fit.ts` fits columns and widens them on entry, by `fitPoints` and `typedPoints` of `src/numfmt.ts`, which lay a value out with `Style.layout` (paddings and blanks marked) and measure it by `fitText` (`glyphs.ts`, `advances.ts`). Sessions compare what cells show before and after an edit that recomputes nothing (`src/session/looks.ts`) and after a structure operation (`src/session/moved.ts`).

Most cells are text or General numbers in a default column, and their text is the value's shortest decimal form, so they cost almost nothing. On the synthetic book (54,778 cells) a cold computation takes 310 ms with display and 279 ms without, in the closures mode.

## How it was checked

`npx tsx bench/coverage/run.ts display <set> [ru]` writes a book of values or formulas with their formats, one column per width, and merged areas where a case needs one. Excel reads back each cell's `string value` (its `Range.Text`) and `color of font object of display format`. Answers are in `test/excel/display/` and `test/excel/display/ru-RU/`, and `test/display-excel.test.ts` compares the engine with them.

| set | cells | what |
|---|---|---|
| `widths` | 1,523 | character widths, column room, fractional and default widths |
| `general` | 5,166 | 82 numbers in General, columns from half a character to 14 characters |
| `codes` | 1,674 per locale | 99 number formats on 12 values, 38 date formats on 15 values |
| `fit` | 678 per locale | booleans, errors and text in narrow columns, numbers near the edge, fills, merged areas, colours, date separators, calendar edges |
| `names` | 70 per locale | month names, long dates and times, large fractions, values no section takes |
| `held` | 58 per locale | predicted by the engine before Excel computed them; 57 of 58 matched in each locale, and the miss became the rule for values no section takes |
| `builtin` | 300 per locale | the built-in formats 0 to 49 by number on six values, with the codes Excel reports for them |
| `builtin-held` | 350 per locale | built-in formats on other values and widths, predicted by the engine before Excel computed them, all matched |

`TEXT` was measured as formulas in the case sets `text-format` and `text-format-more` (en-US) and `text-format-ru` and `text-format-ru-more` (ru-RU), 771 cases in all.

The cases write a format as a custom code, so a code equal to a built-in format does not show by the locale. The sets of built-in formats write the number alone, as Excel stores it.

The 1904 date system was checked with the workbook set `date1904` (`bench/coverage/workbook/`): the same cells and `TEXT` formulas in a book of each system.

Excel for Windows showed the same sets (`bench/windows`, `test/display-win.test.ts`). Its regional data, where they part from the Mac's, are the engine's: the built-in codes, month names in ru-RU, the times of built-in formats. The cells that show otherwise meet a column's edge, where Windows, laying text out in pixels, shows what the engine shows at another width, and the ru-RU long date `[$-F800]`, which the engine keeps as the Mac's. `test/display-excel.test.ts` covers the Mac's answers in the places where the engine follows Windows (`MAC_SHOWN`).

AppleScript returns no colour for an exact 0 in a format of fewer than three sections (`[Red]0`) and for text under a coloured text section (`[Red]@`), though Excel draws the colour on screen there. The engine gives the section's colour in these places, and the comparison accepts any colour.

Fitting columns was measured by `npx tsx bench/coverage/run.ts fit <set> [ru]`: a book of one column per case, which Excel fits (`autofit` of the entire column) or types the case's texts into in turn, through `value`, `formula2 local` or the keyboard; then it reads each column's width and what its cells show and saves the book, whose `<col>` elements give the width Excel stores. Answers are in `test/excel/fit/`, and `test/fit-excel.test.ts` compares `fitColumns`, `parseEntry` and `entryWidth` with them, width and shown text.

| set | columns, en-US and ru-RU | what | predicted before the run |
|---|---|---|---|
| `autofit` | 161 and 161 | text of every width from 3 to 60 points and wider, numbers in General and by formats, booleans, errors, several cells, empty columns, merged areas, a hidden row and column, fonts, formulas and a spill | 11 and 11 |
| `typed` | 87 and 82 | numbers, dates, money, percents and formulas typed through the object model and on the keyboard, into columns of the default width, set by hand and fitted, and in turn | 21 and 19 |
| `glyphs` | 332, en-US | each character of the table alone and in a run of about 600 points, Russian words | 253 |
| `held` | 86 and 85 | held out from the others: text near 90 points and outside ASCII, General near its own text, paddings and blanks, typing that fits exactly or drops decimals | 76 and 74 |

The first sets showed the rules: the predictions of `autofit` and `typed` came from a guess (text width plus the padding, no widening of a General number); `glyphs` showed that Excel measures text outside ASCII by the font's advances. Of the misses in `held`, five were typed through the object model, which the prediction widened though `typed` had shown it does not; the others became rules: a General number that fits exactly widens its column, a whole number's blank fraction keeps the room of its digits and slash, characters of Latin-1 have their own widths, a fitted width equal to the default stays. The engine now gives Excel's width and text on all 994 columns but the 67 in `test/fit-known.ts`, for the reasons below.

Excel 16.0.20430 for Windows fitted the same books and typed the same texts through `Value2` and `Formula2Local` (`bench/windows/run.ts fit <set>`, `test/excel/win/fit/`, `test/fit-win.test.ts`; 912 columns, the keyboard's cases being the Mac's alone). It does what the Mac does: it fits a column with something to show and marks it `bestFit`, shows a hidden column, leaves a column with nothing to show, and widens nothing for text set through the object model. Its widths are its own: it lays text out in pixels, a digit 7 at 96 dpi, and saves the width in sevenths of a digit, so of the columns it fitted only 16 of 161 (`autofit`, either locale), 34 of 332 (`glyphs`) and 15 of 71 (`held`) have the Mac's width. 1234567890 gets 11 on Windows and 11.1640625 on the Mac, a long date 21.7109375 and 19.6640625, 300 points of `b` 59.85546875 and 52.5.

## Known differences

- **Fonts.** Widths are those of Calibri 11 in Excel for Mac. A cell's own font (another font or size, bold), indent, wrapping, shrinking and rotation are ignored, and characters outside the measured table count as a digit. Excel for Windows lays out text in pixels, with a digit 7 pixels wide at 96 dpi, so near the edge of a column it can show `#` where the engine shows the number, or the other way round.
- **Locale tags.** Only `[$-409]`, `[$-419]`, `[$-F800]` and `[$-F400]` change names or formats. Other locale codes, calendars and numeral systems (`[$-x-sysdate]`, `[DBNum1]`, `B2`) are read and ignored.
- **Fitting columns.** Widths are Excel for Mac's; Excel for Windows fits to its pixels (above). A cell's own font is not taken into account (bold text that the engine fits to 115 points gets 117 in Excel), nor characters outside the tables (`中文字符` counts as four digits). Text outside ASCII is measured by the advance widths of Carlito, metrically Calibri's, which give Excel's width on short text but a point off on some long runs (19 of 70 runs of one Cyrillic letter 600 points wide, `Итого по разделу 1`, ru-RU's long date). Where Excel for Mac shows its own short date (`3/15/26`) or ru-RU month names (`март`), the engine fits the text Excel for Windows shows (`3/15/2026`, `Март`). The widest column is 255 here and 255.83 in Excel.
- **A second decimal separator** among the decimals of a format (`#,##0,00` in ru-RU's `TEXT`) is a literal between digits; Excel shows one digit fewer before it for some numbers.
