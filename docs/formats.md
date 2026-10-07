# Number formats and what cells show

A cell of a result carries its value and, by default, what Excel shows in it: `CellResult.text` (VBA's `Range.Text`) and `color`, the colour of its format's section. `TEXT` formats with the same engine. This document has the rules as measured in Excel, how they were measured, how the engine is built and where it differs.

## The API

- `compute(book, { display })` and `openSession(book, { display })`: `display` is `true` by default. With it every cell of `sheets[].cells` (spilled cells too) has `text`, and `color` when its section has one; session changes (`setCells`, `setLayout`, `restructure`, `setNames`, `setClock`) carry them for the cells they list, and also list the cells whose text or colour changed while their value did not. With `display: false` results and changes are as before: values alone, nothing spent on formatting.
- `displayValue(value, { format, locale, width })` shows one value without a book: a cell being edited, a value a host computed itself. It takes the cell's value as the result gives it (a number, text, a boolean, `null`, `{ error }`).
- The value is always the number in `value`, whatever the text shows; the text is derived from the value, the format, the locale and the width, so it is not part of the values hash (`hash` and `inputHash` do not change).
- The format of a cell is its own `format`, else its row's, else its column's. The width is its column's `width`, in xlsx's units, or the sum of the columns of a merged area it is the top-left cell of.

Alternatives that were weighed:

- **Text instead of the number in `value` by default** (what a grid draws, with an option for numbers): one field fewer, but every consumer of values (hashes, comparisons, exports) would have to ask for numbers back, and a number would turn into text depending on a column's width. Rejected: Mikhail's rule is that `value` stays the number.
- **Display only through a separate call** (`displayValue` per visible cell, or `session.render(range)`): no cost on computing, but the default would not be "as in Excel", and the host would have to track which cells changed their look after an edit of a width or a format; the session knows it and returns them. Kept as `displayValue` for cells outside a book.
- **A `shown` object (`{ text, color }`) per cell instead of two flat fields**: one more object for each of tens of thousands of cells, and `CellChange` would nest too. Rejected for flat optional fields, typed together as `CellDisplay`.
- **Display options per call of `result()`** instead of per session: a session that tracks looks must know before edits whether to compare them, so the option belongs to the session.

## How it was measured

`npx tsx bench/coverage/run.ts display <set> [ru]` (`bench/coverage/display.ts`) writes a book whose cells hold the cases' values (or formulas) with their formats, one column per width, merged areas where a case asks for one, and reads back from Excel 16.113 for Mac `string value` of each cell (its `Range.Text`) and `color of font object of display format` (the colour Excel draws the text in). Excel en-US runs with the English interface and locale, ru-RU with the system's Russian ones, under the shared Excel lock. Answers are in `test/excel/display/` (en-US) and `test/excel/display/ru-RU/`; `test/display-excel.test.ts` computes the same books with the engine.

ExcelJS writes a format equal to one of Excel's built-in formats by its number, and Excel shows built-in formats 14 to 22 by the regional settings (`d-mmm-yy` as `31.янв.25` in ru-RU). The harness writes every case's format as a custom code, which is what a `format` string is here.

| set | cases | what |
|---|---|---|
| `widths` | 1,523 (en-US) | widths of characters by fills, the room of columns by `#`, digits against widths, fractional widths and the default column |
| `general` | 5,166 (en-US) | 82 numbers in the General format in columns from half a character to 14 characters, in steps of a sixth |
| `codes` | 1,674 each | 99 number codes on 12 values and 38 date codes on 15 values, in a wide column |
| `fit` | 678 each | booleans, errors and text in narrow columns, formatted numbers near the edge, fills and padding against the width, text past 255 characters, merged areas, the default column, colours, date separators, edges of the calendar and of fractions |
| `names` | 70 each | month names after a day, the long date and time, large improper fractions, numbers under `@` in narrow columns, values no conditional section takes |
| `held`, `held-ru` | 58 each | held out: predicted by the engine and committed before Excel ran (below) |

`TEXT` was measured as formulas in the case sets `text-format` (479) and `text-format-more` (22) in en-US and `text-format-ru` (254) and `text-format-ru-more` (16) in ru-RU (`test/excel/cases/`).

### Colours

AppleScript gives the colour of the display format's font, which is the section's colour: `[Red]` reads as `#FF0000`, `[Color10]` as `#008000`, and black (the automatic colour, `[Black]`, `[Color1]`) as `#000000`, which the answers leave out. It gives no colour for an exact 0 in a format of fewer than three sections (`[Red]0`, `[Red]0;[Blue]0`) and for text under a coloured text section (`[Red]@`, `0;0;0;[Red]@`), where Excel draws the colour on the screen: a screenshot of Excel 16.113 with those cells shows them red (and `x` under `[Red]0;[Blue]0;[Green]0;[Magenta]@` magenta). The engine gives the section's colour there, and `test/display-api.test.ts` pins it; the Excel comparison accepts any colour in those two places. The names are Microsoft's eight (`[Black]`, `[White]`, `[Red]`, `[Green]`, `[Blue]`, `[Yellow]`, `[Magenta]`, `[Cyan]`, any case) and `[ColorN]` an index into Excel's default palette of 56 colours ("Review guidelines for customizing a number format", Microsoft Support; the palette is xlsx's indexed colours 8 to 63); the measured ones (`[Color1]`, `[Color10]`, `[Color56]`, the names, and five held-out indexes) agree with it.

### Held-out cases

Before the run of `held` and `held-ru`, each of their 58 cases got the engine's answer as its `predicted` field, committed (`95cf949`). Excel confirmed 57 of 58 in each locale. The miss, in both: 0 in `[<0]0;[>0]0`, which no section takes, shows `#` in Excel; the engine showed General. Two more cases in each locale and in `TEXT` measured the rule (with conditions in two or three sections a value none of them takes is `#VALUE!` in `TEXT` and `#` in a cell; with one section it shows General), the engine follows it now, and `test/display-excel.test.ts` lists the miss.

## The rules

### Width

Excel lays out text in points of Calibri 11, the default font of an xlsx file. A digit is 6 points. A column of xlsx width `w` (the `width` of `ColumnInput`, in digits with the padding) is `trunc(6w + 21·6/256)` points wide, as xlsx's formula with a maximum digit width of 6; a column without a width is 53 (8 digits and the padding of 5). Of that, 5 points are padding: the text fits when the sum of its characters' widths is at most the rest. The widths of 155 characters were measured by fills (`0*x` repeats `x` across the column: the count gives `x`'s width): space and `'` 2; `. , - ( ) : { } f i j l I` 3; `/ ! r s t z J г т` 4; `+ # ^ ~ < > = ? _ E F L S T Y Z`, lower-case `a c e g k v x y`, Cyrillic `а в е ё з к с у ч х ь э я Г Е Ё З Т` 5; digits, `$ ₽ € £ ¥` and the most letters 6; `D G H N O Q U м ф ы` and the like 7; `% & w ж ш щ ю Ф Ы` 8; `m M Ж М` 9; `W Ш Щ Ю` 10. Every case of the `widths` set agrees with this model exactly.

A merged area is as wide as its columns together.

### What does not fit

- A formatted number or date that does not fit shows `#` across the cell, one per 5 points of room (`#` is 5 wide), at most 255; nothing when the room is under 5.
- A shown text longer than 255 characters (a fill, a long number) shows `#` too: `0* ` in a column of width 86 shows 253 characters, at 87 `#`.
- A boolean or an error shows whole when it fits, else `#` the same way (`TRUE` needs 23 points, `ИСТИНА` 38).
- Text is never cut: it runs over the next cells, and `Range.Text` is the text itself, however narrow the column.

### General

A number in the General format shows the text `TEXT(x,"General")` gives (at most 11 characters, the sign aside) when that fits, counting each character as a digit and a leading minus as 3 points. Else, as measured on 5,166 cells:

1. A number from 10^-4 to under 10^11 drops decimals, rounding, down to the most that fit, as long as some digit of the number shows (`3.14159265` is `3.142`, then `3.1`, then `3`; `0.000123456789` is `0.0001235`, then `0.0001`).
2. Then scientific notation with fewer digits of the mantissa (`1.23E+08`, `1.2E+08`, `1E+08`).
3. A number that rounds to 0 in a room of one to four digits shows `0` (`1E-10` in a column one digit wide); a positive number from 10^-4 whose rounding is one digit shows it in any width (`0.5` shows `1` even with no room).
4. The shortest scientific form when its real width fits and the number rounds to 6 digits or more, or is under 10^-4 (`1E+05` for 999999.5 in 28 points, but `#` for 99999.4).
5. Else `#` across the cell, one at least.

### Sections

Without conditions: the first section for positive numbers (and zero when there are fewer than three), the second for negative ones (which then write no minus of their own), the third for zero, the fourth, or else a section with `@`, for text. With conditions (`[>100]`, `[<=-5]`, `[=1]`, `[<>0]`) the first of the first three sections whose condition holds or that has none; a value none of them takes shows General when the code has one section, and nothing otherwise (`#VALUE!` in `TEXT`, `#` in a cell). A section writes the minus of a negative number unless every number it can take is negative (`[Red][<0]0;0` shows −5 as `5`, `[<=-5]"low";[>=5]"high";0` shows −1 as `-1`); a section of literals writes the minus before them (`[=1]"one";"other"` shows −2 as `-other`). An empty section shows nothing (`0;;`). Booleans and errors show as they are, booleans under a text section as its text (`0;-0;0;"is "@` shows FALSE as `is FALSE`).

### Numbers

- Digits as before: `0` a digit or zero, `#` a digit, `?` a digit or a space; a group separator between digits groups by thousands, after them scales by 1000 (`0.0,,`); `%` multiplies by 100 (both rounded to 15 significant digits first); a second point is a literal.
- Scientific notation: as many integer digits as integer placeholders, the exponent a multiple of their count (`##0.0E+0` shows 0.5 as `500.0E-3`), `E+` writes the exponent's sign always, `E-` only a minus, the exponent padded to its placeholders (`1.00E+300`). An exponent written in lower case is `#VALUE!` in `TEXT`, and Excel refuses to load an xlsx file with one.
- Fractions: the integer part before the numerator if the code has one; a denominator by its placeholders (the closest fraction) or fixed (`?/8`); a numerator over 2,147,483,646 shows `#` (`#VALUE!` in `TEXT`): `?/?` shows 1E+15 as `#` but 2,000,000,001/2 as itself. With an integer part a negative number keeps its minus and a whole 0 shows `0` (`-0    `).
- Positional numbers keep 15 significant digits and zeros after them (`TEXT(1E+250,"0")` is a 1 and 250 zeros).
- Literals: `"…"`, `\x`, and characters that need no quotes (`$ - + / ( ) : space`); `_x` is one space whatever `x` (Excel draws the width of `x`, `Range.Text` has a space, and fitting counts a space); `*x` repeats `x` across the room the rest leaves, the last fill only, earlier ones dropping out; `[$₽-419]` is the symbol `₽` (any symbol) and the locale `419`.

### Dates and times

- Excel's 1900 calendar: serial 0 is `1/0/1900`, 60 is `2/29/1900`; a negative serial or one from 2,958,466 (past 9999-12-31) shows `#` (`#VALUE!` in `TEXT`).
- The time is rounded to what the code shows: whole seconds, or the decimals of `ss.0` to `ss.000` (45688.999999 is `2025-02-01 00:00:00` by `yyyy-mm-dd hh:mm:ss`, `23:59:59.9` by `h:mm:ss.0`).
- `m` is minutes after hours or before seconds, else the month; `mmmmm` the month's first letter; `ddd`, `dddd`, `mmm`, `mmmm` the locale's names; `A/P` writes the letter in the code's case (`a/p` shows `6:00 p`).
- Elapsed time: `[h]`, `[mm]`, `[ss]` count hours, minutes or seconds from 0; a negative value shows with a minus when the section has elapsed time alone (`[h]` shows −1 as `-24`), else `#`.
- In a cell, a bare `/` in a date is the locale's date separator (`dd/mm/yyyy` shows `31.01.2025` in ru-RU); in `TEXT` it stays `/`. Other separators and quoted or escaped ones are as written.
- `[$-409]` and `[$-419]` take English and Russian names; `[$-F800]` is the system's long date (en-US `dddd, mmmm d, yyyy`, ru-RU `dddd, d mmmm yyyy "г."` with month names in the genitive, `пятница, 31 января 2025 г.`) and `[$-F400]` the long time (`h:mm:ss AM/PM`, `hh:mm:ss`). Other tags are read and ignored.

### Locales

A cell's format is read as an xlsx file stores it, in en-US syntax, whatever the book's locale, and shown with the locale's decimal and group separators, names, booleans, error texts and date separator: `#,##0.00` is `1 234 567,89` in ru-RU, `mmmm` is `январь`, `TRUE` is `ИСТИНА`, `#N/A` is `#Н/Д`. `TEXT` reads the locale's own codes (`# ##0,00`, `ДД.ММ.ГГГГ`, `Основной`) and its colour names: in ru-RU `[Красный]`, `[Синий]`, `[Цвет10]` and the other Russian names, while `[Red]` and `[Color10]` are `#VALUE!`. Month names in ru-RU are nominative (`31 январь 2025`) except in the long date.

### `TEXT` and a cell

| | `TEXT` | a cell |
|---|---|---|
| codes | the book's locale's (`# ##0,00` in ru-RU) | as xlsx stores them (`#,##0.00`) |
| General | 11 characters at most | narrowed to the width |
| a number wider than the column | (no width) | `#` across the cell |
| a value its section cannot show (a date out of range, no section takes it, more than 255 characters) | `#VALUE!` | `#` across the cell |
| `*x` | dropped | `x` across the room |
| `_x` | a space | a space |
| colour | none | the section's |
| `/` in a date | `/` | the locale's date separator |

## The engine

`src/numfmt.ts` assembles `src/numfmt/`: `parse.ts` reads a code into sections (colour, condition, locale tag, tokens) by a dialect's letters and separators and caches the result per dialect and code, so a column of one format is parsed once; `select.ts` chooses the section; `number.ts` (digits, fractions, exponents), `date.ts` (dates, times, elapsed time) and `general.ts` (General as text, as `TEXT` writes it and narrowed to a width) format; `glyphs.ts` has the widths; `palette.ts` the colours; `style.ts` what a locale writes. `formatWith` is `TEXT`, `displayWith` a cell. A locale gives a dialect for `TEXT` and a style for cells (`src/locale/rules/format.ts`); `Rules.display` is the book's (`src/locale/rules.ts`). `src/shown.ts` finds a cell's format and width (row, column, merged area) and is used by results (`src/result.ts`) and sessions; `src/session/looks.ts` compares what cells show before and after an edit that recomputes nothing, and `src/session/moved.ts` after a structure operation.

Speed, on the synthetic book (3,400 rows, 3,907 formulas, 54,778 cells; cold compute, median of 5 runs after a warm-up, Apple M5 Pro, Node 26): 308 ms with the display and 276 ms without (`node --expose-gc --import tsx test/fixtures/bench-child.ts 3400 closures [off]`). Most cells are text or numbers in the General format of a default column; their text is the value's shortest decimal text when that has at most 11 characters, so most cost nothing beyond a string.

## Known differences

- **Fonts.** Widths are those of Calibri 11 in Excel for Mac. A cell's own font (`style` with another font or size, bold) is not taken into account, nor indent, wrapping, shrinking or rotation; characters outside the measured table (other scripts, rare signs) count as a digit. Excel for Windows lays text out in pixels of a digit 7 wide at 96 dpi: near the edge of a column it can show `#` where the engine shows the number, or the other way round.
- **Built-in formats.** An xlsx file stores some formats by number, and Excel shows numbers 14 to 22 by the regional settings (`d-mmm-yy` as `31.янв.25` in ru-RU); a `format` here is a code, shown as written (`31-янв-25`).
- **Locale tags.** Only `[$-409]`, `[$-419]`, `[$-F800]` and `[$-F400]` change names or formats; other locale codes, calendars and numeral systems (`[$-x-sysdate]`, `[DBNum1]`, `B2`) are read and ignored.
- **A second decimal separator** among the decimals of a number format (`#,##0,00` in ru-RU's `TEXT`) is a literal between digits; Excel shows one digit fewer before it in some numbers (`RU_KNOWN`, `secondDecimal`).
