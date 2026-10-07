# Number formats and shown text

Besides its value, each cell of a result carries what Excel shows in it: `text`, the same as VBA's `Range.Text`, and `color`, the colour of the format section that was used. The `TEXT` function formats with the same engine.

The rules below were measured in Excel 16.113 for Mac, en-US and ru-RU.

## API

- `compute(book, { display })` and `openSession(book, { display })`. `display` is `true` by default, and then every cell in `sheets[].cells` has `text` (spilled cells too) and `color` when its section has one. Session edits return these fields for the cells they list, and also list cells whose text or colour changed while their value did not: a new `format` through `setCells`, a width, a row or column format or a merged area through `setLayout`, a moved cell through `restructure`.
- `display: false` returns values only and spends nothing on formatting.
- `displayValue(value, { format, locale, width })` formats one value without a book, for example a cell being edited. It takes a value as results give it: a number, text, a boolean, `null` or `{ error }`.

The number always stays in `value`. The text depends on the value, the format, the locale and the column width, so it is not part of the values hash, and a hash does not change when a column gets wider.

A cell's format is its own `format`, else its row's, else its column's. Its width is its column's `width` (in xlsx units), or the width of all columns of the merged area it starts.

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
- Time is rounded to what the format shows: whole seconds, or the decimals of `ss.0` to `ss.000`. 45688.999999 is `2025-02-01 00:00:00` with `yyyy-mm-dd hh:mm:ss` and `23:59:59.9` with `h:mm:ss.0`.
- `m` means minutes after hours or before seconds, and the month elsewhere. `mmmmm` is the month's first letter. `ddd`, `dddd`, `mmm` and `mmmm` are the locale's names. `A/P` writes the letter in the format's case (`a/p` shows `6:00 p`).
- `[h]`, `[mm]` and `[ss]` count elapsed hours, minutes or seconds from 0. A negative value shows with a minus when the section has elapsed time alone (`[h]` shows −1 as `-24`), and `#` otherwise.
- In a cell, a bare `/` in a date format is the locale's date separator: `dd/mm/yyyy` shows `31.01.2025` in ru-RU. In `TEXT` it stays `/`. Quoted and escaped separators stay as written.
- `[$-409]` and `[$-419]` select English and Russian names. `[$-F800]` is the system's long date (en-US `dddd, mmmm d, yyyy`; ru-RU `пятница, 31 января 2025 г.`, with months in the genitive) and `[$-F400]` the long time. Other tags are read and ignored.

## Locales

A cell's format is read the way xlsx stores it, in en-US syntax, whatever the book's locale. It is shown with the locale's separators, names, booleans, error texts and date separator: in ru-RU `#,##0.00` shows `1 234 567,89`, `mmmm` shows `январь`, TRUE shows `ИСТИНА`, `#N/A` shows `#Н/Д`.

`TEXT` reads the locale's own codes (`# ##0,00`, `ДД.ММ.ГГГГ`, `Основной`) and colour names (`[Красный]`, `[Синий]`, `[Цвет10]`); in ru-RU, `[Red]` and `[Color10]` are `#VALUE!`. Month names in ru-RU are in the nominative (`31 январь 2025`), except in the long date.

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

## How the engine does it

`src/numfmt/` holds the engine, assembled in `src/numfmt.ts`. `parse.ts` reads a format into sections (colour, condition, locale tag, tokens) and caches it per dialect and format, so a column with one format is parsed once. `select.ts` picks the section. `number.ts`, `date.ts` and `general.ts` format numbers, dates and General. `glyphs.ts` has the character widths, `palette.ts` the colours, `style.ts` what a locale writes. `formatWith` is `TEXT` and `displayWith` a cell.

`src/shown.ts` finds a cell's format and width. Sessions compare what cells show before and after an edit that recomputes nothing (`src/session/looks.ts`) and after a structure operation (`src/session/moved.ts`).

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

`TEXT` was measured as formulas in the case sets `text-format` and `text-format-more` (en-US) and `text-format-ru` and `text-format-ru-more` (ru-RU), 771 cases in all.

The cases write every format as a custom code. Excel shows built-in formats 14 to 22 by the regional settings, so a format equal to one of them would show differently.

AppleScript returns no colour for an exact 0 in a format of fewer than three sections (`[Red]0`) and for text under a coloured text section (`[Red]@`), though Excel draws the colour on screen there. The engine gives the section's colour in these places, and the comparison accepts any colour.

## Known differences

- **Fonts.** Widths are those of Calibri 11 in Excel for Mac. A cell's own font (another font or size, bold), indent, wrapping, shrinking and rotation are ignored, and characters outside the measured table count as a digit. Excel for Windows lays out text in pixels, with a digit 7 pixels wide at 96 dpi, so near the edge of a column it can show `#` where the engine shows the number, or the other way round.
- **Built-in formats.** xlsx stores some formats by number, and Excel shows numbers 14 to 22 by the regional settings (`d-mmm-yy` as `31.янв.25` in ru-RU). Here a format is a code and shows as written (`31-янв-25`).
- **Locale tags.** Only `[$-409]`, `[$-419]`, `[$-F800]` and `[$-F400]` change names or formats. Other locale codes, calendars and numeral systems (`[$-x-sysdate]`, `[DBNum1]`, `B2`) are read and ignored.
- **A second decimal separator** among the decimals of a format (`#,##0,00` in ru-RU's `TEXT`) is a literal between digits; Excel shows one digit fewer before it for some numbers.
