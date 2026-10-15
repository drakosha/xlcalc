# Addresses and formula text

A grid, a dialog or a server that keeps versions of a book needs Excel's rules for addresses, names and values without computing anything. The package gives them by the engine's own code: `xlcalc/a1` holds the addresses and names in a few kilobytes and none of the engine, `xlcalc/text` the references of formula text and the names a reference writes, for a browser that colours a formula as it is typed without the engine, and `xlcalc` exports the same functions with the helpers that need a locale. These take a loaded locale's tag or a locale object, which needs no loading; `quoteSheetIfNeeded` takes the object alone, so `xlcalc/a1` carries no list of locales.

```js
import { cellAddress, linesKey, parseLines, parseRange, quoteSheetIfNeeded, rangeAddress, sameName } from 'xlcalc/a1'
import { entryText, errorText, needsApostrophe, sameValue, validateSheetName } from 'xlcalc'
import { ruRU } from 'xlcalc/locales/ru-RU'
import { refAt, scanReferences } from 'xlcalc/text'

console.log(parseRange('$d$9:B2'), parseRange('ZZZ1'), rangeAddress({ top: 1, left: 2, bottom: 1048576, right: 4 }))
// → { kind: 'cells', top: 2, left: 2, bottom: 9, right: 4 } null B:D
console.log(cellAddress(12, 2), parseLines('B:D', 'columns'), linesKey(9, 5, 'rows')) // → B12 { first: 2, last: 4 } 5:9
console.log(quoteSheetIfNeeded('Sales'), quoteSheetIfNeeded('Q1 data'), quoteSheetIfNeeded('R2D2'), quoteSheetIfNeeded('ИСТИНА', ruRU), sameName('straße', 'STRASSE'))
// → Sales 'Q1 data' 'R2D2' 'ИСТИНА' true
const bad = validateSheetName('Итоги/2025', { locale: ruRU })
console.log(bad.code, bad.at, bad.message) // → character 5 Имя листа не может содержать «/» (Excel запрещает : \ / ? * [ ])
console.log(errorText('DIV0', ruRU), entryText(0.125, { numberFormat: '0.0%' }), entryText(45688.75, { numberFormat: 14, locale: ruRU }), entryText('007'), needsApostrophe('1.5', { locale: ruRU }))
// → #ДЕЛ/0! 12.5% 31.01.2025  18:00:00 '007 true
console.log(scanReferences('=СУММ(B2:B9;\'Q1 data\'!C1;"D4")', { locale: ruRU }).map((r) => `${r.at}-${r.end} ${r.text} ${r.sheet ?? ''}`))
// → [ '6-11 B2:B9 ', "12-24 'Q1 data'!C1 Q1 data" ]
console.log(scanReferences('=A1+)B2*"C3').map((r) => r.text)) // → [ 'A1', 'B2' ]
console.log(refAt('=SUM(A1,').action, refAt('=A1').action, refAt('="a').action, refAt('=SUM(A1;', undefined, { locale: ruRU })) // → insert commit none { action: 'insert', at: 8 }
console.log(sameValue(0.1 + 0.2, 0.3), sameValue(1, true)) // → true false
```

Rows and columns are numbered from 1, as A1 numbers them. `parseCell` and `parseRange` take `$` and letters in either case and give null past `XFD1048576` or for anything else; a range given corner to corner the other way round is put in order, as Excel orders a typed reference, and its `kind` says whether it named cells, whole columns (`B:D`) or whole rows (`2:9`). `cellAddress`, `colLetters`, `colNumber` and `rangeAddress` write them back, `rangeAddress` as Excel writes a range: `B2` for one cell, `B:D` for a range over every row. `parseLines` and `linesKey` read and write the keys of `SheetInput.columns` and `rows` (`C`, `B:D`, `5`, `5:9`), null for a key a book does not take, and `patchLines` patches such a map as `setLayout` does ([Layout](grid.md#layout)); `rectsMeet` and `rectContains` compare rectangles; `MAX_ROWS` and `MAX_COLS` are the sheet's size.

`refAt(text, caret?, { locale })` says what an arrow does in formula text typed over a cell, as Excel's Point mode decides it: `insert`, a reference to the cell the arrow points at goes in at `at` (after an `=` of its own with `equals`, at the start of the text); `commit`, the text is entered and the selection moves, as after an operand; `none`, nothing happens (inside a string, after a character the locale's formulas do not have). Moving or extending a reference already pointed is the host's ([Point mode](entry.md#point-mode)).

`callAt(text, caret?, { locale })` says which call the caret stands in, for that tip: the innermost function whose brackets hold it, its `name` as typed and `fn` in English (for `functionSignature`), the argument's `index` from 0 and where the call's brackets `open` and `close` (null while open); null outside any call. Strings, quoted sheets' names, tables' brackets, array constants and a group's brackets hold no separator of the call, and half typed text gives its call (the package's own rule, read as `refAt` reads, not measured in Excel).

`functionSignature(name, locale?)`, in the root alone (its names would take `xlcalc/text` past 50 KB), gives a function's arguments for a tip while a formula is typed, as Excel's tip lists them: the name in the locale and its English name `fn`, `args` (each argument's name and whether a call can leave it out), `repeat` (how many of the last arguments go on past the list, 1 for SUM, 2 for SUMIFS) and the tip's `text`; null for a name outside the whitelist. The name is taken in English or in the locale, in any case. The names of the arguments are the library's own, in the style of Excel's (common words, Excel's abbreviations spelled out); which arguments are optional and how many a call takes are the measured ones of each function. For a window like Excel's Insert Function, the signature has the function's `category`, as that window lists the functions (`FUNCTION_CATEGORIES` gives them in its order, `compatibility` holding the old names such as `NORMDIST`), and the category's name in the locale (`categoryName`). The window's "Most Recently Used" and "All" are the host's to keep.

What a function does, one line in the library's own words, is `functionDescription(fn, locale?)` of `xlcalc/function-text`, by the English name. The texts of the tips are apart from the engine, so that a host that only computes takes none of them: the en-US descriptions are `xlcalc/function-text`, and the Russian words, the arguments' names, the categories' names and the descriptions, are `ruRUFunctionText` of `xlcalc/function-text/ru-RU`, which a host joins to the locale it gives (`{ ...ruRU, ...ruRUFunctionText }`). The ru-RU locale alone writes the tips in English words.

```js
import { callAt, functionSignature } from 'xlcalc'
import { functionDescription } from 'xlcalc/function-text'
import { ruRUFunctionText } from 'xlcalc/function-text/ru-RU'
import { ruRU } from 'xlcalc/locales/ru-RU'

const ru = { ...ruRU, ...ruRUFunctionText }
console.log(functionSignature('sumifs').text) // → SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
console.log(functionSignature('ВПР', ru).text) // → ВПР(искомое;таблица;номер_столбца;[приближённо])
const vlookup = functionSignature('VLOOKUP')
console.log(vlookup.categoryName, '|', functionDescription(vlookup.fn)) // → Lookup & Reference | A value from a column of a table, found in its first column.
const call = callAt('=СУММЕСЛИ(A:A;">1";', undefined, { locale: ru })
console.log(call.fn, call.index, functionSignature(call.fn, ru).args[call.index].name, '|', functionDescription(call.fn, ru)) // → SUMIF 2 диапазон_суммирования | Сумма ячеек, отвечающих условию.
```

`nameKey` and `sameName` are the rule by which the engine compares names of sheets, tables and defined names: in any case, by Unicode's upper case in every locale (`straße` is `STRASSE`), but the dotless `ı` and the long `ſ`, which Excel for Windows keeps apart from `I` and `S`. `quoteSheetIfNeeded` writes a sheet's name in a reference as Excel writes it, measured in Excel for Mac and for Windows: in quotes (an apostrophe inside doubled) when the name is anything but letters, digits, `_` and `.` after a letter or `_`, when it is a cell of the sheet (`A1`; not `XFE1`), when it begins with an R1C1 reference with a number (`R2D2`, `C3PO`; not `Region`, nor `C1.x`, where a `.` follows), or when it is `R`, `C`, `RC` or one of the locale's booleans (`TRUE`, `'ИСТИНА'` in ru-RU's formula text). Stored formulas keep every sheet in quotes, as their canonical text; the text shown to a person (`toLocal`, a window's `formulaLocal` and `editText`, the text of `explain`) quotes a sheet by this rule. `validateSheetName` checks a sheet's name as a book's sheets are checked (31 characters at most, none of `: \ / ? * [ ]`, no apostrophe at either end, not `History`): `{ ok: true }`, or a refusal with the code `empty`, `too_long`, `character`, `apostrophe` or `reserved`, a message in `opts.locale`'s language and, for a character, its place.

`errorText` writes an error as a cell shows it in a locale; any other value without a book is `displayValue(value, { locale })`. `entryText` is the text a cell opens with for editing, as Excel's formula bar shows its value: a number whole (15 digits, positional up to 21 characters), a percent cell's number as a percent, a date or a time by the locale's short date and long time, booleans and errors in the locale's words, and text with an apostrophe where typing it would read as something else, which `needsApostrophe` tells by `parseEntry`. `sameValue` compares two values as the values hash does (numbers to 15 significant digits, `-0` as 0, errors by code), and `canonValue` is a value's text in the hash. `scanReferences` finds the references in formula text as it is typed, stored or in a locale, with their places for colouring them: cells, whole columns and rows, sheets, references across sheets and spills, and structured references with their table, columns and specifiers. It reads the text with the engine's tokenizer: a token that does not read (a character a formula does not take, an error literal half typed, a sheet with nothing after it) is skipped and the text read on past it, and an open string, quoted sheet name or bracket takes the rest of the text, so a formula half typed or broken in the middle gives the references around the place (the package's own rule, not one measured in Excel). How the rules were measured is in [docs/references.md](references.md#sheet-names-in-references) and [docs/entry.md](entry.md#the-text-a-cell-opens-with).
