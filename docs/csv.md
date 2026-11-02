# Reading and writing CSV

The `xlcalc/csv` subpath reads a CSV file into a workbook of one sheet as Excel opens it, and writes a sheet of a workbook back as Excel saves it as CSV. Every field becomes what Excel makes of it in the workbook's locale: `1 234,5` in ru-RU is a number with the format Excel sets, `3/15/2026` in en-US a date, `5%` a percentage, `TRUE` a boolean. The subpath has no dependencies and none of the calculation: the workbook it returns goes straight to `compute`, a session or `writeXlsx`.

```js
import { compute } from 'xlcalc'
import { readCsv, writeCsv } from 'xlcalc/csv'

const text = 'Item,Qty,Price,Sold\r\nApple,3,$1.50,3/15/2026\r\n"Pears, green",10,$0.25,3/16/2026\r\n'
const read = readCsv(text, { name: 'sales.csv' })
const sheet = read.book.sheets[0]
console.log(sheet.name, read.delimiter, sheet.cells.B3, sheet.cells.C2.value, sheet.cells.D2) // → sales , { value: 10 } 1.5 { value: 46096, numberFormat: 14 }
const book = { sheets: [{ ...sheet, cells: { ...sheet.cells, E2: { formula: 'B2*C2' } } }] }
console.log(writeCsv({ book, result: compute(book) }, { bom: false }).text.split('\r\n')[1]) // → Apple,3,$1.50 ,3/15/2026,4.5
```

## In short

`readCsv(data, opts)` takes the file's bytes (a `Uint8Array`, a Node `Buffer`, an `ArrayBuffer`) or its text, and returns `book`, a workbook of one sheet; `delimiter` and `encoding`, what the file was read with; and `issues`, where the workbook holds a field otherwise than Excel would. `writeCsv({ book, result }, opts)` writes a sheet, the active one by default, with the values of its formulas from `result` (`compute`'s, a session's `getSnapshot()`, or `readXlsx`'s `cached`), and returns the file's bytes `data`, its `text` and its `issues`.

- `locale`: the workbook's locale, a loaded locale's tag or a locale object, en-US by default. Fields are read as text typed in it, and the workbook takes its tag.
- `name`: the file's name. The sheet is named after it, without the extension, as Excel names it, and a `.tsv` file is read at tabs.
- `delimiter`: the character between fields, or `'locale'` for the locale's list separator.
- `encoding`: the file's encoding, by its label (`utf-8`, `windows-1251`), when the host knows it.
- `formulas`: `true` reads formulas as Excel does; by default they stay text.
- `codesAsText`: `true` keeps `00123` and numbers of more than 15 digits as text.
- `now`: the moment whose year completes a date typed without one, as for `parseEntry`.
- `limits.cells`: the most fields that are not empty; a larger file throws an `Error`.

## Reading

### Encoding

A byte order mark says the encoding: UTF-8, UTF-16LE or UTF-16BE. A file without one is UTF-16 when every other byte of its start is zero, UTF-8 when its bytes are valid UTF-8, and otherwise in the Windows code page of the locale: Windows-1252 in en-US, Windows-1251 in ru-RU, as Excel for Windows reads such a file. `opts.encoding` overrides all of this; a mark of the given encoding is skipped. Malformed bytes become U+FFFD. `encoding` in the result names the encoding by its WHATWG name; it is null for text.

```js
import { readCsv } from 'xlcalc/csv'
import { ruRU } from 'xlcalc/locales/ru-RU'

const cp1251 = new Uint8Array([0xcf, 0xf0, 0xe8, 0xe2, 0xe5, 0xf2, 0x3b, 0x31, 0x2c, 0x35])
const read = readCsv(cp1251, { locale: ruRU })
console.log(read.encoding, read.delimiter, read.book.sheets[0].cells) // → windows-1251 ; { A1: { value: 'Привет' }, B1: { value: 1.5 } }
```

### Delimiter

A first line `sep=;` names the delimiter, as Excel takes it, and is no row of the sheet. Otherwise `opts.delimiter` gives it, a `.tsv` name gives tabs, and without either the reader guesses from the first ten lines: of `;`, `,` and a tab, the one that splits the most lines into the same number of fields, `;` first on a tie, so a ru-RU file of `1,5;2,5` splits at `;`. A file that no candidate splits is read at the locale's list separator. `delimiter: 'locale'` is what Excel does: it reads a `.csv` file at the locale's list separator (`,` in en-US, `;` in ru-RU), and a file marked as UTF-16 at tabs.

```js
import { readCsv } from 'xlcalc/csv'

console.log(readCsv('a;b;c\r\n1;2;3').delimiter, readCsv('sep=;\r\na;b').book.sheets[0].cells) // → ; { A1: { value: 'a' }, B1: { value: 'b' } }
console.log(readCsv('a;b;c', { delimiter: 'locale' }).book.sheets[0].cells) // → { A1: { value: 'a;b;c' } }
```

### Fields

Fields follow RFC 4180 as Excel reads it. A field that begins with a quote runs to its closing quote; a doubled quote inside is one quote, and delimiters and line breaks inside stay. Whatever follows the closing quote up to the delimiter is added (`"a"b` is `ab`), a quote elsewhere is an ordinary character, and a quote never closed runs to the end of the file. Rows end at CRLF, LF or CR, and a line break at the end of the file adds no row. Empty fields and empty lines leave their cells empty. A field with a line feed wraps its cell: the workbook gets the style `{ alignment: { wrap: true } }`.

### Values

Each field is read as Excel reads text pasted into a cell in the locale (`parseEntry`, [Cell entry](entry.md)): numbers with the locale's separators, dates and times, percentages, money and fractions get the number format Excel sets, `TRUE` is a boolean and `#N/A` an error. Spaces around a number are dropped, and a quoted number is still a number. A leading apostrophe stays part of the text. A field longer than 32,767 characters is cut there, rows past row 1,048,576 and fields past column XFD are left out, each with an issue.

Excel takes a field beginning with `=`, a `+` or `-` that begins no number, or `@` before a function call (`@SUM(A1:A3)`) as a formula. A file from elsewhere could bring in formulas this way (CSV injection), so by default the reader keeps such fields as text and lists them in an issue of code `formula`. With `formulas: true` it reads them as Excel does: in the locale's formula text (`=СУММ(A1;1,5)` in ru-RU), as formulas of the old kind, so a range takes its row's value (`=A1:A2` reads as `@A1:A2`, [Legacy formulas](legacy.md)). A formula that does not parse stays text, with an issue.

Excel reads `00123` as 123 and a number of more than 15 digits with its digits past the fifteenth as zeros. `codesAsText: true` keeps a field of digits that begins with zero or has more than 15 digits as text, for codes and article numbers.

```js
import { readCsv } from 'xlcalc/csv'

const text = '=1+2,00123,1234567890123456789,@SUM(1),\'007\r\n'
const plain = readCsv(text)
console.log(plain.book.sheets[0].cells, plain.issues[0].cells)
// → {
//     A1: { value: '=1+2' },
//     B1: { value: 123 },
//     C1: { value: 1234567890123450000 },
//     D1: { value: '@SUM(1)' },
//     E1: { value: "'007" }
//   } [ 'A1', 'D1' ]
console.log(readCsv(text, { formulas: true, codesAsText: true }).book.sheets[0].cells)
// → {
//     A1: { formula: '1+2' },
//     B1: { value: '00123' },
//     C1: { value: '1234567890123456789' },
//     D1: { formula: 'SUM(1)' },
//     E1: { value: "'007" }
//   }
```

## Writing

`writeCsv` writes what Excel saves as CSV: each cell's value as it is displayed by its number format in the locale, but not cut to its column (a date never becomes `####`, and General shows its full eleven characters). It writes the sheet's used range, from its first row and column holding a value or a format to its last, with every row as wide as the range. A field that holds the delimiter, a quote, a tab or a line break is quoted, with its quotes doubled. Rows are joined by CRLF, with none after the last. Formulas are written as their values, arrays with every cell they fill.

`delimiter` is the locale's list separator by default (`'locale'`), as Excel writes; `locale` is the workbook's unless given. `encoding` is `utf-8` by default, with its byte order mark, as Excel's CSV UTF-8; `utf-16le`, also with its mark; or `ansi`, the locale's Windows code page, as Excel's plain CSV, where a character the code page lacks is written as `?` with an issue. `bom: false` leaves the mark out. Writing a workbook of several sheets reports the others, which the file cannot hold.

```js
import { compute } from 'xlcalc'
import { writeCsv } from 'xlcalc/csv'
import 'xlcalc/locales/ru-RU'

const book = { locale: 'ru-RU', sheets: [{ name: 'Лист', cells: { B2: { value: 1234.5, numberFormat: 4 }, C2: { formula: 'B2>1000' }, B3: { value: 'а;б' } } }] }
const csv = writeCsv({ book, result: compute(book) })
console.log(JSON.stringify(csv.text), csv.data.length) // → "1 234,50;ИСТИНА\r\n\"а;б\";" 34
```

## Checked against Excel

Excel 16.114 for Mac opened every file of the sets in `bench/coverage/csv/` in en-US and ru-RU, saved it as xlsx and as CSV, and saved the case books as CSV (`test/csv-excel.test.ts`): fields, quotes and line breaks, empty fields and rows, numbers, dates, times, percentages and money, booleans and errors, formulas, leading zeros and long numbers, `sep=`, other delimiters, encodings, sheet names, the used range, short rows, arrays. The sets `held` and `held-ru` were predicted before Excel ran them. Excel for Windows was not run on CSV files.

## Known differences

- By default the reader guesses the delimiter, where Excel always takes the locale's list separator: a file of semicolons opens in en-US Excel as one column. `delimiter: 'locale'` reads as Excel does.
- By default formulas are text; `formulas: true` reads them as Excel does.
- A file without a byte order mark that is valid UTF-8 is read as UTF-8. Excel for Mac reads it in Mac Roman or Mac Cyrillic, and reads no UTF-16BE file and no UTF-16 file without a mark; a file in a Windows code page Excel for Mac reads in its Mac code page too. The reader follows Excel for Windows's code pages.
- Excel for Mac saves CSV in Mac Roman or Mac Cyrillic; `encoding: 'ansi'` writes the Windows code page, as Excel for Windows. Excel for Mac shows format 14 as `m/d/yy` and names months in ru-RU otherwise than Excel for Windows, whose text the writer writes ([Localization](locale.md)).
- Excel reads `#SPILL!`, `#CALC!` and `#GETTING_DATA` (as `#N/A`) as errors; the reader keeps them as text. Excel gives `5 €`, `€-5` and `-€5` euro formats; the reader reads the number without a format.
