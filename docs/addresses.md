# Addresses and formula text

A grid, a dialog or a server that keeps versions of a workbook needs Excel's rules for addresses, names and values without calculating anything. The package provides them using the engine's own code: `xlcalc/a1` contains addresses and names in a few kilobytes without any of the engine; `xlcalc/text` contains the references in formula text and the names a reference uses, for a browser that colors a formula as the user types it without the engine; and `xlcalc` exports the same functions plus the helpers that need a locale. These accept a loaded locale's tag or a locale object, which does not need to be loaded; `quoteSheetIfNeeded` accepts only the object (a tag throws an `InputError`), so `xlcalc/a1` carries no list of locales.

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

Rows and columns are numbered from 1, as in A1 notation. `parseCell` and `parseRange` accept `$` and letters in either case, and return null beyond `XFD1048576` or for anything else; a range given with its corners in reverse order is normalized, as Excel normalizes a typed reference, and its `kind` says whether it refers to cells, entire columns (`B:D`) or entire rows (`2:9`). `cellAddress`, `colLetters`, `colNumber` and `rangeAddress` convert them back, `rangeAddress` as Excel writes a range: `B2` for one cell, `B:D` for a range covering every row. `parseLines` and `linesKey` read and write the keys of `SheetInput.columns` and `rows` (`C`, `B:D`, `5`, `5:9`), returning null for a key a workbook does not accept, and `patchLines` patches such a map as `setLayout` does ([Layout](grid.md#layout)); `rectsMeet` and `rectContains` compare rectangles; `MAX_ROWS` and `MAX_COLS` are the size of a sheet.

`refAt(text, caret?, { locale })` tells what an arrow key does in formula text typed over a cell, as Excel's Point mode decides it: `insert` means a reference to the cell the arrow points to is inserted at `at` (after an `=` of its own with `equals`, at the start of the text); `commit` means the text is entered and the selection moves, as after an operand; `none` means nothing happens (inside a string, or after a character the locale's formulas do not allow there). Moving or extending a reference already inserted is up to the host ([Point mode](entry.md#point-mode)).

`callAt(text, caret?, { locale })` tells which call the caret is in, for the argument tooltip: the innermost function whose parentheses contain it, its `name` as typed and `fn` in English (for `functionSignature`), the argument's `index` from 0, and where the call's parentheses `open` and `close` (null while still open); null outside any call. Strings, quoted sheet names, table brackets, array constants and grouping parentheses contain no argument separators of the call, and partially typed text still returns its call (this is the package's own rule, parsed as `refAt` parses, not verified in Excel).

`nameAt(text, caret?, { locale })` returns the name of a function being typed at the caret, for Formula AutoComplete's list: `{ from, to, typed }`, where it starts and ends and the part before the caret, in a position where an operand may start (after the formula's first character, an operator, the range's `:` included, `(`, the locale's list separator, or a space); null in a string, a quoted sheet name or table brackets, or after `!` or `$`. `FUNCTION_CATEGORIES` lists the categories of Insert Function in Excel's order.

`functionSignature(name, locale?)`, available only from the root entry (its names would push `xlcalc/text` past its 55 KB budget), returns a function's arguments for a tooltip while a formula is typed, as Excel's tooltip lists them: the name in the locale and its English name `fn`, `args` (each argument's name and whether a call can omit it), `repeat` (how many of the last arguments repeat beyond the list: 1 for SUM, 2 for SUMIFS) and the tooltip's `text`; null for a name outside the whitelist. The name is accepted in English or in the locale, in any case. The argument names are the library's own, in Excel's style (common words, with Excel's abbreviations spelled out); which arguments are optional and how many a call accepts were verified for each function. For a dialog like Excel's Insert Function, the signature includes the function's `category`, as that dialog groups the functions (`FUNCTION_CATEGORIES` lists them in its order, with `compatibility` holding the old names such as `NORMDIST`), and the category's name in the locale (`categoryName`). The dialog's "Most Recently Used" and "All" lists are up to the host.

A one-line description of what a function does, in the library's own words, comes from `functionDescription(fn, locale?)` in `xlcalc/function-text`, by English name. The tooltip texts are kept apart from the engine, so a host that only calculates does not load any of them: the en-US descriptions are in `xlcalc/function-text`, and the Russian texts (argument names, category names and descriptions) are `ruRUFunctionText` in `xlcalc/function-text/ru-RU`, which a host merges into the locale it passes (`{ ...ruRU, ...ruRUFunctionText }`). The ru-RU locale on its own writes the tooltips in English.

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

`nameKey` and `sameName` are the rule by which the engine compares names of sheets, tables and defined names: case-insensitively, by Unicode uppercase in every locale (`straße` is `STRASSE`), except for the dotless `ı` and the long `ſ`, which Excel for Windows keeps distinct from `I` and `S`. `quoteSheetIfNeeded` writes a sheet name in a reference as Excel writes it, as verified in Excel for Mac and for Windows: in quotes (with an apostrophe inside doubled) when the name contains anything other than letters, digits, `_` and `.` after a leading letter or `_`, when it is a cell of the sheet (`A1`, but not `XFE1`), when it starts with an R1C1 reference containing a number (`R2D2`, `C3PO`, but not `Region`, nor `C1.x`, where a `.` follows), or when it is `R`, `C`, `RC` or one of the locale's booleans (`TRUE`, or `'ИСТИНА'` in ru-RU formula text). Stored formulas keep every sheet in quotes, as in their canonical text; text displayed to users (`toLocal`, a viewport's `formulaLocal` and `editText`, the text of `explain`) quotes a sheet by this rule. `validateSheetName` checks a sheet name as a workbook's sheets are checked (at most 31 characters, none of `: \ / ? * [ ]`, no apostrophe at either end, not `History`): it returns `{ ok: true }`, or a rejection with the code `empty`, `too_long`, `character`, `apostrophe` or `reserved`, a message in the language of `opts.locale` and, for a character, its position.

`errorText` writes an error as a cell displays it in a locale; any other value, without a workbook, is formatted by `displayValue(value, { locale })`. `entryText` is the text a cell opens with in edit mode, as Excel's formula bar shows its value: a number in full (15 digits, positional notation up to 21 characters), a percentage cell's number as a percentage, a date or a time in the locale's short date and long time formats, booleans and errors in the locale's words, and text with an apostrophe where typing it would produce something else, which `needsApostrophe` determines using `parseEntry`. `sameValue` compares two values as the values hash does (numbers to 15 significant digits, `-0` as 0, errors by code), and `canonValue` is a value's text in the hash.

`scanReferences` finds the references in formula text as it is typed, stored or written in a locale, with their positions for coloring: cells, entire columns and rows, sheets, references across sheets, spill references, and structured references with their table, columns and specifiers. It reads the text with the engine's tokenizer: a token that does not parse (a character not allowed in a formula, a partially typed error literal, a sheet name with nothing after it) is skipped and reading continues after it, and an unclosed string, quoted sheet name or bracket takes the rest of the text, so a partially typed or broken formula still returns the references around the problem (this is the package's own rule, not verified in Excel). How the rules were verified is described in [docs/references.md](references.md#sheet-names-in-references) and [docs/entry.md](entry.md#the-text-a-cell-opens-with).
