# API reference

Every function, value and type the package exports, entry by entry, with the comments of its source. Guides with
examples are the [README](../README.md) and the other files of `docs/`. This file is made from the package's types
(`npx tsx test/api-reference.ts --write`): change the comments in the source, not here.

- [`xlcalc`](#xlcalc)
- [`xlcalc/a1`](#xlcalca1)
- [`xlcalc/text`](#xlcalctext)
- [`xlcalc/entry`](#xlcalcentry)
- [`xlcalc/xlsx`](#xlcalcxlsx)
- [`xlcalc/xlsx/keep`](#xlcalcxlsxkeep)
- [`xlcalc/worker`](#xlcalcworker)
- [`xlcalc/function-text`](#xlcalcfunction-text)
- [`xlcalc/function-text/ru-RU`](#xlcalcfunction-textru-ru)
- [`xlcalc/locales`](#xlcalclocales)
- [`xlcalc/locales/ru-RU`](#xlcalclocalesru-ru)
- [`xlcalc/regex`](#xlcalcregex)
- [`xlcalc/contract`](#xlcalccontract)
- [Types](#types)

## xlcalc

Types: see [Types](#types).

### <a id="xlcalc-builtinformatcode"></a>`builtinFormatCode`

```ts
function builtinFormatCode(id: number, locale?: string | Locale): string | undefined
```

The format code with which a built-in number format (0 to 49, [`NumberFormat`](#numberformat) as a number) is displayed in `locale` (a
loaded locale's tag or a locale object; en-US without one), in en-US syntax as a cell's `numberFormat` accepts it;
undefined for a number that is not a built-in format.

### <a id="xlcalc-calcerror"></a>`CalcError`

```ts
class CalcError extends Error
```

Why an asynchronous or limited call failed, by `code`; a `limit` error names the limit, its maximum, the size and the formula's cell.

### <a id="xlcalc-callat"></a>`callAt`

```ts
function callAt(text: string, caret?: number, opts?: ScanOptions): CallAt | null
```

The innermost call whose brackets enclose `caret` (the end of the text by default) in formula text `text`, and the
argument the caret is in; null outside any call, and for text that is not a formula (not beginning with `=`, `+` or
`-`). `opts.locale`: the locale the text is typed in (its list separator and names); without it, the text is treated
as en-US.

### <a id="xlcalc-canonvalue"></a>`canonValue`

```ts
function canonValue(value: CellValue): string
```

A value's canonical text, as its line in the values hash holds it: `n 1234.5` (15 significant digits, `-0` as 0),
`s "text"`, `b TRUE`, `e DIV0`, `_` for an empty cell.

### <a id="xlcalc-celladdress"></a>`cellAddress`

```ts
function cellAddress(row: number, col: number): string
```

The address of a cell (`B12`).

### <a id="xlcalc-colletters"></a>`colLetters`

```ts
function colLetters(n: number): string
```

The letters of column `n` (1 is `A`, 16,384 is `XFD`).

### <a id="xlcalc-colnumber"></a>`colNumber`

```ts
function colNumber(letters: string): number | null
```

The number of a column from its letters, in any case (`B` is 2), or null when they are not a column of the sheet (`XFE`).

### <a id="xlcalc-columnwidthfromchars"></a>`columnWidthFromChars`

```ts
function columnWidthFromChars(chars: number): number
```

Converts a width in characters, as typed into Excel for Mac's Column Width box (0 to 255), to the stored xlsx width Excel makes of it.

### <a id="xlcalc-columnwidthfrompoints"></a>`columnWidthFromPoints`

```ts
function columnWidthFromPoints(points: number): number
```

Whole points back to a stored width in 256ths; fractional points round down, as the model's geometry does.

### <a id="xlcalc-columnwidthtochars"></a>`columnWidthToChars`

```ts
function columnWidthToChars(width?: number | null): number
```

A stored xlsx width as Excel for Mac's Column Width box displays it, in characters of the workbook's default font, to
two decimal places (the default column, null or undefined, is 8); in the library's Mac point model, like
[`columnWidthToPoints`](#xlcalc-columnwidthtopoints).

### <a id="xlcalc-columnwidthtopoints"></a>`columnWidthToPoints`

```ts
function columnWidthToPoints(width?: number | null): number
```

A stored xlsx width in the library's Mac point model; null or undefined is the default 53 points.

### <a id="xlcalc-comparevalues"></a>`compareValues`

```ts
function compareValues(a: CellValue, b: CellValue, opts?: CompareValuesOptions): number
```

Excel's sort order for two cells' values, as the Sort command of a host's display filter sorts them: numbers, then
text in any case (in `locale`'s text order, en-US's without it), then FALSE and TRUE, then errors, all equal to each
other, and blanks (null) last, even with `descending`; negative when `a` comes first, 0 for equal values, which a
stable sort keeps in their original order, as Excel does. The SORT function instead orders errors by their
ERROR.TYPE number.

### <a id="xlcalc-compute"></a>`compute`

```ts
function compute(book: BookInput, opts?: ComputeOptions): ComputeResult
```

Calculates the entire workbook. Without `opts.clock`, the clock is the current time and a random seed; the
result records it in `clock`, and passing that clock again reproduces the result. Input of the wrong
shape throws [`InputError`](#xlcalc-inputerror); invalid formulas do not throw, they produce error values with reasons in `diagnostics`.

### <a id="xlcalc-config_version"></a>`CONFIG_VERSION`

```ts
const CONFIG_VERSION: "xlcalc/1;canon15"
```

Engine configuration version: part of the values hash. It changes with any change that can shift
values (function semantics, runtime rules, canonical form).

### <a id="xlcalc-copyformula"></a>`copyFormula`

```ts
function copyFormula(text: string, from: CellRef, to: CellRef, book?: CopyContext): string
```

The text of a formula copied from cell `from` to cell `to`, as `copyRange` pastes it: relative parts shift by the
offset, absolute ones stay, a reference pushed beyond the sheet's edge becomes `#REF!`, and only the changed
references are rewritten. With the workbook (its sheet names and tables), a formula copied out of a table names the
table in its references to it, as Excel pastes it; without it, the sheet names in the formula are taken as written.

### <a id="xlcalc-createinlinecalc"></a>`createInlineCalc`

```ts
function createInlineCalc(opts?: InlineCalcOptions): Calc
```

A calculator on the same thread ([`Calc`](#calc)): it calculates in slices of `sliceMs` that yield to the event loop, with the
same deadlines, cancellation and limits as a pool's calls; its sessions return the same results as a pool's.

### <a id="xlcalc-createsession"></a>`createSession`

```ts
function createSession(book: BookInput, opts?: SessionOptions): Session
```

Opens a session. Without options, it keeps the clock of the time it was opened (with a random seed)
until `setClock`; `{ live: true }` takes a fresh clock on every recalculation.

### <a id="xlcalc-datetoserial"></a>`dateToSerial`

```ts
function dateToSerial(text: string, opts?: DateSystemOptions): number
```

The serial number of a date and time in the form [`CellInput.date`](#cellinput) accepts (`2026-01-31` is 46053, `2026-01-31T18:00`
is 46053.75), in `opts.dateSystem` (1900 without one; 1904 counts from 1904-01-01).

### <a id="xlcalc-debuggeneratedcode"></a>`debugGeneratedCode`

```ts
function debugGeneratedCode(book: BookInput, opts?: ComputeOptions): string
```

The code that `eval` mode generates for the workbook, in evaluation order, for reading while debugging; not part of the contract.

### <a id="xlcalc-default_row_height"></a>`DEFAULT_ROW_HEIGHT`

```ts
const DEFAULT_ROW_HEIGHT: 15
```

The height of a row without a height of its own, in points: one line of Calibri 11, the workbook's default font.

### <a id="xlcalc-displayvalue"></a>`displayValue`

```ts
function displayValue(value: CellValue, opts?: DisplayOptions): CellDisplay
```

A value as a cell would display it, without a workbook (for example, a cell being edited): in `numberFormat` (General
without one), in `locale` and `dateSystem`, in a column of `width` (no limit without one: a value is not shortened to
fit a column it was not given), as [`CellResult.text`](#cellresult) and `color` display it.

### <a id="xlcalc-editstructure"></a>`editStructure`

```ts
function editStructure(book: BookInput, op: StructureOp): StructureResult
```

Restructures a workbook; the input workbook is not modified. A copy that pastes values (from a filtered sheet)
calculates the workbook at the current time.

### <a id="xlcalc-entrytext"></a>`entryText`

```ts
function entryText(value: CellValue, opts?: EntryTextOptions): string
```

The text a cell holding `value` opens with in edit mode, as Excel's formula bar displays it, in `opts.locale` and
according to the cell's `opts.numberFormat`: a number to 15 significant digits with the locale's decimal separator
(`1234.5`, `1234,5`), a percentage cell's number as a percentage (`12%`), a date or time cell's value in the locale's
short date and long time formats (`1/31/2025  6:00:00 PM`, `31.01.2025  18:00:00`), a boolean or an error in the
locale's words, and text with an apostrophe where typing it would otherwise be parsed as something else (`'007`).
Parsing the text as edited in place gives back the value ([`parseEntry`](#xlcalcentry-parseentry) with `inPlace: true`), but for a negative
percentage, which Excel's in-place editor reads as a formula (`-5%` is `=-5%%`).

### <a id="xlcalc-entrywidth"></a>`entryWidth`

```ts
function entryWidth(value: CellValue, opts?: EntryWidthOptions): number | undefined
```

The width Excel gives a column when `value`, displayed in `opts.numberFormat`, is typed into one of its cells
(`test/excel/fit/`): a number for which AutoFit would make more room widens a column whose width is the default or
was fitted by Excel (`bestFit`), as AutoFit fits it, even to the default width (Excel keeps a fitted width that
equals it); undefined when the column keeps its width (text overflows, General drops decimals, a manually set width
stays, a column never narrows).

### <a id="xlcalc-enus"></a>`enUS`

```ts
const enUS: Locale
```

en-US formula text, the invariant syntax; the package's other locales are separate modules (`xlcalc/locales/ru-RU`).

### <a id="xlcalc-errortext"></a>`errorText`

```ts
function errorText(code: ErrorCode, locale?: string | Locale): string
```

An error as a cell displays it in `locale` (a loaded locale's tag or a locale object, en-US without one): `#DIV/0!`, `#ДЕЛ/0!` in ru-RU.

### <a id="xlcalc-escapecolumn"></a>`escapeColumn`

```ts
function escapeColumn(name: string): string
```

A column's name in a structured reference as Excel writes it: `'` before each `[`, `]`, `#` and `'` in it, and before
a leading `@` (`x#y` becomes `x'#y`, `@at` becomes `'@at`), everything else as is (`Unit Price`).

### <a id="xlcalc-findmatcher"></a>`findMatcher`

```ts
function findMatcher(query: string, opts?: FindTextOptions): ((text: string) => boolean) | null
```

Excel's Find over a host's own texts, matching as `findCells` matches a cell's text: a test for a text, or null when nothing can match (an empty query, `~` alone).

### <a id="xlcalc-findreplacer"></a>`findReplacer`

```ts
function findReplacer(query: string, by: string, opts?: FindTextOptions): ((text: string) => string | null) | null
```

Excel's Replace over a host's own texts, as `replaceInCells` replaces in a cell's text: the text with each match replaced by `by`, or null where nothing matches.

### <a id="xlcalc-fitcolumns"></a>`fitColumns`

```ts
function fitColumns(book: BookInput, result: ComputeResult, sheet: string, columns: string): Record<string, number | null>
```

The widths AutoFit gives columns `C` or `B:D` of the sheet `sheet` of a calculated workbook, from the values of its
result (`compute(book)`), by column letter: null for the default width; a column with nothing to display is left out.

### <a id="xlcalc-fitrows"></a>`fitRows`

```ts
function fitRows(book: BookInput, result: ComputeResult, sheet: string, rows: string): Record<string, number | null>
```

The heights AutoFit gives rows `5` or `2:9` of the sheet `sheet` of a calculated workbook, from the values of its
result (`compute(book)`), by spans of rows with the same height (`2:4`): null for the default height.

### <a id="xlcalc-formatclipboard"></a>`formatClipboard`

```ts
function formatClipboard(rows: string[][], opts?: { trailingNewline?: boolean | undefined; }): string
```

Excel's plain-text copy: tabs, CRLF rows and quoted control characters; Windows ends with CRLF, Mac does not.

### <a id="xlcalc-function_categories"></a>`FUNCTION_CATEGORIES`

```ts
const FUNCTION_CATEGORIES: readonly FunctionCategory[]
```

The function categories in the order Excel's Insert Function dialog lists them (their members come from `function-info.ts`).

### <a id="xlcalc-function_names"></a>`FUNCTION_NAMES`

```ts
const FUNCTION_NAMES: readonly string[]
```

The whitelist: every function a formula can call, by its English name, in alphabetical order.

### <a id="xlcalc-functionsignature"></a>`functionSignature`

```ts
function functionSignature(name: string, locale?: string | Locale): FunctionSignature | null
```

The signature of a function, by its name in English or in the locale (in any case), for an argument tooltip while a
formula is being typed; null for a name that is not a function in the whitelist.

### <a id="xlcalc-inputerror"></a>`InputError`

```ts
class InputError extends TypeError
```

Input of the wrong shape. `path` points at the offending part, for example `book.sheets[1].cells["B0"]`.

### <a id="xlcalc-internstyle"></a>`internStyle`

```ts
function internStyle(styles: readonly CellStyle[], style: CellStyle): { index: number; styles: CellStyle[]; }
```

A style's index in a style table, as `patchCells` finds it: the index of the first equal style and the same array, or
a new array with a copy of the style appended and its index. Equal styles have the same fields with the same values,
in any order, a protection's defaults aside (`{ protection: { locked: true } }` is `{}`). A table is read once per array: the next call, given the returned table, reads only the added style. A
table changed in place (a style added, replaced or removed) is read again; a style changed in place goes unnoticed.

### <a id="xlcalc-lineskey"></a>`linesKey`

```ts
function linesKey(first: number, last: number, axis: "columns" | "rows"): string
```

The key for lines `first` to `last` (`C`, `B:D`; `5`, `5:9`), in ascending order regardless of which bound is given first.

### <a id="xlcalc-livebook"></a>`liveBook`

```ts
function liveBook(session: Session): LiveBook
```

The live workbook of a synchronous session ([`LiveBook`](#livebook)): the same object for the same session. An asynchronous
session throws [`InputError`](#xlcalc-inputerror): its writes cannot return what they calculated.

### <a id="xlcalc-locales"></a>`LOCALES`

```ts
const LOCALES: ReadonlyMap<string, Locale>
```

The loaded locales by tag, en-US and the registered ones, in the order they were loaded. The map is for reading only.

### <a id="xlcalc-max_cols"></a>`MAX_COLS`

```ts
const MAX_COLS: 16384
```

The number of columns on a sheet, as in Excel: 16,384, A to XFD.

### <a id="xlcalc-max_formula_depth"></a>`MAX_FORMULA_DEPTH`

```ts
const MAX_FORMULA_DEPTH: 64
```

Excel's limit on nested function calls in a formula: 64.

### <a id="xlcalc-max_formula_length"></a>`MAX_FORMULA_LENGTH`

```ts
const MAX_FORMULA_LENGTH: 8192
```

Excel's limit on the length of a formula's text: 8,192 characters.

### <a id="xlcalc-max_outline"></a>`MAX_OUTLINE`

```ts
const MAX_OUTLINE: 7
```

The deepest outline of rows or columns Excel allows: seven levels of groups below the sheet's own level.

### <a id="xlcalc-max_row_height"></a>`MAX_ROW_HEIGHT`

```ts
const MAX_ROW_HEIGHT: 409.5
```

The maximum row height Excel allows, in points.

### <a id="xlcalc-max_rows"></a>`MAX_ROWS`

```ts
const MAX_ROWS: 1048576
```

The number of rows on a sheet, as in Excel: 1,048,576.

### <a id="xlcalc-mergedarea"></a>`mergedArea`

```ts
function mergedArea(r: CellRect, merges: readonly CellRect[]): CellRect
```

The rectangle a merge covers (`mergeCells`, Merge Cells): `rect` extended to include each merged area of `merges`
that it overlaps, repeated for each area it then overlaps, until it overlaps no other.

### <a id="xlcalc-nameat"></a>`nameAt`

```ts
function nameAt(text: string, caret?: number, opts?: ScanOptions): NameAt | null
```

The name of a function being typed at `caret` (the end of the text by default) in formula text `text`, for the
Formula AutoComplete list: the name characters around the caret, starting with a letter or `_`, where an operand may
begin (after the formula's first character, an operator, the range's `:` included, `(`, the locale's list separator or a space); null where no
function can appear: no name at the caret, inside a string, a quoted sheet name or a table's brackets, after a
sheet's `!` or a `$`, and for text that is not a formula. `opts.locale`: the locale the text is typed in (its list
separator); without it, the text is treated as en-US.

### <a id="xlcalc-namekey"></a>`nameKey`

```ts
function nameKey(name: string): string
```

The key by which names are compared: sheet, table and defined names are the same name in any case (`Sales`,
`SALES`), by Unicode uppercase in every locale (`straße` is `STRASSE`).

### <a id="xlcalc-needsapostrophe"></a>`needsApostrophe`

```ts
function needsApostrophe(text: string, opts?: EntryTextOptions): boolean
```

Whether text typed into a cell would be parsed as something other than that text, so the cell needs an apostrophe
before it to keep it as text: a number, a date, a boolean, an error or a formula in the locale (`007`, `1.5` in
ru-RU, `=A1`), a number in a percentage cell, text that begins with an apostrophe, or empty text.

### <a id="xlcalc-numberformatcategory"></a>`numberFormatCategory`

```ts
function numberFormatCategory(format: NumberFormat, locale?: string | Locale): NumberFormatCategory
```

The category of a number format in `locale` (a loaded locale's tag or a locale object; en-US without one), as the
ribbon's Number Format box displays it, with the category's name there: for a format code as stored in an xlsx file,
or a built-in format by its number.

### <a id="xlcalc-numberformatchoices"></a>`numberFormatChoices`

```ts
function numberFormatChoices(locale?: string | Locale): FormatChoice[]
```

The items of the Number Format list in `locale`, in list order: each item's key, name, the format it applies and that format's code in the locale.

### <a id="xlcalc-outlineedit"></a>`outlineEdit`

```ts
function outlineEdit(sheet: Pick<SheetInput, "columns" | "rows" | "outline">, op: OutlineEditOp): OutlineEditResult
```

Excel's outline commands on a standalone sheet, as [`editStructure`](#xlcalc-editstructure) performs them in a workbook ([`OutlineEditOp`](#outlineeditop):
Group, Ungroup, Show and Hide Detail, the level buttons): the sheet's rows and columns after the command, taking its
summary settings into account, or why Excel rejects the command, in English
(`{ ok: false, code: 'detail', message: 'the detail of row 11 is already shown' }`). The sheet is checked as
[`outlineGroups`](#xlcalc-outlinegroups) checks it, and the operation as [`editStructure`](#xlcalc-editstructure) checks it.

### <a id="xlcalc-outlinegroups"></a>`outlineGroups`

```ts
function outlineGroups(sheet: Pick<SheetInput, "columns" | "rows" | "outline">): SheetOutline
```

A sheet's groups of rows and columns as Excel draws them ([`SheetOutline`](#sheetoutline)), from its lines' outline levels, hidden
state and collapsed markers and its summary settings: a [`SheetInput`](#sheetinput), or a [`SheetLayout`](#sheetlayout) as `getLayout()` returns it.
Each axis's groups are ordered by first line, the outer before the inner when two start together; a group is
collapsed when Excel would keep the marker on its summary line ([`RowInput.collapsed`](#rowinput)), whatever markers the sheet
holds.

### <a id="xlcalc-parsecell"></a>`parseCell`

```ts
function parseCell(cell: string): CellPosition | null
```

A cell's row and column (`B12`, `$B$12`, `b12`), or null when the text is not a cell of the sheet (`ZZZ1`, `A0`, `B2:C3`).

### <a id="xlcalc-parseclipboard"></a>`parseClipboard`

```ts
function parseClipboard(text: string): string[][]
```

Text from the clipboard as the rectangle Excel pastes ([`Session.pasteText`](#session)): rows split at line breaks, made of fields
split at tabs, each row as wide as the longest; a field that begins with a quote runs to its closing quote, with line
breaks and tabs inside it, and a doubled quote stands for one; an empty field clears its cell, and a line break after
the last row is a row of empty fields.

### <a id="xlcalc-parseentry"></a>`parseEntry`

```ts
function parseEntry(text: string, opts?: EntryOptions): EntryResult
```

Text typed into a cell, as Excel interprets it in the workbook's locale: the cell's new input, with the cell's style,
meta, note and link and the format Excel leaves it, or a formula Excel would reject. Empty text clears the value and
keeps the rest. Typed over the cell, or edited in place with `opts.inPlace`.

### <a id="xlcalc-parselines"></a>`parseLines`

```ts
function parseLines(key: string, axis: "columns" | "rows"): LineSpan | null
```

The lines that a key of [`SheetInput.columns`](#sheetinput) or `rows` refers to (`C`, `B:D`; `5`, `5:9`), or null for a key a
workbook does not accept: lowercase, a span of a single line (`C:C`) or in reverse order, or beyond the sheet.

### <a id="xlcalc-parserange"></a>`parseRange`

```ts
function parseRange(range: string): RangeRect | null
```

A range's rectangle and kind: `B2:D9` and `B2` are cells, `B:D` entire columns, `2:9` entire rows, with `$` and
letters in any case; corners given in reverse order are swapped, as Excel reorders such a typed reference. Null for
anything else, or for a range that extends beyond the sheet.

### <a id="xlcalc-patchlines"></a>`patchLines`

```ts
function patchLines<T extends object>(map: Record<string, T> | undefined, span: string, patch: LinePatch<T>): Record<string, T>
```

A map of [`SheetInput.columns`](#sheetinput) or `rows` with `patch` applied to the lines `span` (`C`, `B:D`; `5`, `5:9`; the axis
follows from the key's shape), as `setLayout` applies it: a value sets a field, null removes it, and an omitted field
is left unchanged; spans that the patch cuts through are split, and lines left without fields are dropped. Returns a
new map; the given one is not modified. The map is checked as a workbook's would be: an object for each span of the
axis, with no two spans overlapping.

### <a id="xlcalc-quotesheetifneeded"></a>`quoteSheetIfNeeded`

```ts
function quoteSheetIfNeeded(name: string, locale?: Pick<Locale, "booleans"> | undefined): string
```

A sheet's name in a reference as Excel writes it: quoted, with any apostrophe inside doubled, when it is not a plain
name (`'Q1 data'`, `'2024'`, `'O''Brien'`), when it is a cell address on the sheet (`'A1'`, `'XFD1048576'`; not
`XFE1`), when it begins with an R1C1 reference (`'R'`, `'R1C1'`, `'R2D2'`) or when it is one of the locale's booleans
(`'TRUE'`; `'ИСТИНА'` with ru-RU). `locale` provides the booleans, en-US's (the stored text, `Range.Formula`) without
one: a locale object ([`ruRU`](#xlcalclocalesru-ru-ruru)), not a tag, so that `xlcalc/a1` carries no list of locales; a tag throws an
[`InputError`](#xlcalc-inputerror). The engine's stored formulas quote every sheet name: this is the text a user reads.

### <a id="xlcalc-rangeaddress"></a>`rangeAddress`

```ts
function rangeAddress(rect: CellRect): string
```

A rectangle's address as Excel writes it: `B2` for a single cell, `B2:D9`, `B:D` when it spans every row, `2:9` when it spans every column.

### <a id="xlcalc-rectcontains"></a>`rectContains`

```ts
function rectContains(outer: CellRect, inner: CellRect): boolean
```

Whether `inner` lies within `outer`.

### <a id="xlcalc-rectsmeet"></a>`rectsMeet`

```ts
function rectsMeet(a: CellRect, b: CellRect): boolean
```

Whether two rectangles share a cell.

### <a id="xlcalc-refat"></a>`refAt`

```ts
function refAt(text: string, caret?: number, opts?: ScanOptions): PointAction
```

What an arrow key does at `caret` (the end of the text by default) in `text` typed over a cell: insert a reference at
the caret, enter the text, or nothing ([`PointAction`](#pointaction)). `opts.locale`: the locale the text is typed in, a loaded
locale's tag or a locale object; without it, the text is treated as en-US.

### <a id="xlcalc-registerlocale"></a>`registerLocale`

```ts
function registerLocale(locale: Locale): void
```

Loads a locale for workbooks in that locale ([`BookInput.locale`](#bookinput)) and for its tag wherever a locale is accepted: one
of the package's locale modules (`import { ruRU } from 'xlcalc/locales/ru-RU'`, then `registerLocale(ruRU)`). Loading
the same object again does nothing; a different object with an already loaded tag throws an [`InputError`](#xlcalc-inputerror). A pool's
threads load it too.

### <a id="xlcalc-registerregex"></a>`registerRegex`

```ts
function registerRegex(l: RegexLevel): void
```

Loads a level of `xlcalc/regex` for match mode 3 of `XLOOKUP` and `XMATCH` in the calculations from now on, a pool's
threads too: `import { pcreEngine } from 'xlcalc/regex'`, then `registerRegex(pcreEngine)` (or [`pcreSyntax`](#xlcalcregex-pcresyntax)). The
last one registered is used.

### <a id="xlcalc-samename"></a>`sameName`

```ts
function sameName(a: string, b: string): boolean
```

Whether two names of sheets, tables or defined names are the same name ([`nameKey`](#xlcalc-namekey)).

### <a id="xlcalc-samevalue"></a>`sameValue`

```ts
function sameValue(a: CellValue | undefined, b: CellValue | undefined): boolean
```

Whether two values are the same for the values hash: numbers equal to 15 significant digits (`0.1 + 0.2` is `0.3`,
`-0` is 0), errors by their code; an absent value (`undefined`) is an empty cell, like null.

### <a id="xlcalc-scanreferences"></a>`scanReferences`

```ts
function scanReferences(formula: string, opts?: ScanOptions): FormulaReference[]
```

The references to cells, entire columns and rows, and tables in formula text, in order, with their positions in it
(`at` to `end`), for coloring them as they are typed. `formula` is a stored formula, or text typed in `opts.locale`;
a leading `=` is skipped. A token that does not parse (a stray character, a half-typed error literal, a sheet name
with nothing after it) is skipped and parsing continues past it; an unclosed string, quoted sheet name or bracket
takes up the rest of the text. Strings, function names and defined names are not references.

### <a id="xlcalc-serialtodate"></a>`serialToDate`

```ts
function serialToDate(serial: number, opts?: DateSystemOptions): string
```

The date and time of a serial number in `opts.dateSystem`, in the form [`CellInput.date`](#cellinput) accepts, without a time zone:
just the date for a whole number (`2026-01-31`), otherwise with the time to the millisecond (`2026-01-31T18:00`,
`2026-01-31T18:00:30.5`). Serial 0 of the 1900 system is `1900-01-00` and 60 is `1900-02-29`, as Excel displays them.

### <a id="xlcalc-stepdecimals"></a>`stepDecimals`

```ts
function stepDecimals(format: NumberFormat | undefined, value: CellValue, step: 1 | -1, opts?: { locale?: string | undefined; width?: number | undefined; }): NumberFormat | undefined
```

The number format Excel's Increase Decimal (`step` 1) or Decrease Decimal (-1) gives a cell of `format` (General
without one) holding `value`, in `opts.locale` (a built-in format's code in it) and a column of `opts.width` (Excel's
standard without one), as Excel for Mac does: a General cell of a number takes its decimals one more or one less
(`0.00` for 1.5), all of them when General shows the number whole, else those it shows in the column; a code takes a
decimal more or less in each section of digits, a copy of its last one (`#.##` to `#.###`); a date, a time, a
fraction, text, accounting's zero (`"-"??`) stay. It returns `format` itself when the step changes nothing: no number
in General, no decimal to take away. Excel applies the active cell's result to the whole selection.

### <a id="xlcalc-toinvariant"></a>`toInvariant`

```ts
function toInvariant(text: string, where: string | Locale, opts?: TranslateOptions): Translation
```

Text typed in a locale (a loaded locale's tag, or a locale object) as a stored formula: `=СУММ(A1:A3;1,5)` becomes
`SUM(A1:A3,1.5)`. A leading `=` is removed, since stored formulas have none; English function names, TRUE/FALSE and
error texts are also accepted, and a sheet name with or without quotes is kept as typed. Error positions refer to
`text`, and error messages are in the locale's language. `opts.table`, the table containing the formula's cell, adds
that table's name before a structured reference without one (`[@Qty]` becomes `Sales[@Qty]`), the reverse of
[`toLocal`](#xlcalc-tolocal) dropping it, so the formula also works outside the table.

### <a id="xlcalc-tolocal"></a>`toLocal`

```ts
function toLocal(formula: string, where: string | Locale, opts?: TranslateOptions): Translation
```

A stored formula as displayed in a locale (a loaded locale's tag, or a locale object): `SUM(A1:A3,1.5)` becomes
`СУММ(A1:A3;1,5)` in ru-RU, and a sheet name is quoted only where Excel quotes it (`'Data'!B2` becomes `Data!B2`, in
en-US too; [`quoteSheetIfNeeded`](#xlcalc-quotesheetifneeded) with the locale's booleans). Structured references are displayed as Excel displays
them, the formula's own row as `@` (`Sales[[#This Row],[Qty]]` becomes `Sales[@Qty]`) and an entire table by its name
alone (`Sales[]` becomes `Sales`); `opts.table`, the table containing the formula's cell, drops that table's name
before its own columns (`[@Qty]`, `SUM([Qty])`), as Excel displays a formula inside a table.

### <a id="xlcalc-validateformula"></a>`validateFormula`

```ts
function validateFormula(text: string, ctx?: ValidateContext): ValidateResult
```

Checks formula text as a cell stores it (without `=`, en-US syntax): returns its canonical text and the functions,
sheets, names, external workbooks and tables it refers to, or why and where it fails to parse ([`ValidateResult`](#validateresult)).
With `ctx`, the sheets, names and tables it refers to are checked against the workbook's.

### <a id="xlcalc-validatelocalformula"></a>`validateLocalFormula`

```ts
function validateLocalFormula(text: string, where: string | Locale, ctx?: ValidateContext): ValidateResult
```

[`validateFormula`](#xlcalc-validateformula) for text typed in a locale (a loaded locale's tag, or a locale object): the same result, with the
canonical stored text, error positions referring to `text` (a leading `=` is allowed) and messages in the locale's
language.

### <a id="xlcalc-validatename"></a>`validateName`

```ts
function validateName(name: string, opts?: NameOptions): NameCheck
```

Checks a name ([`NameInput.name`](#nameinput)) against Excel's rules in a workbook of `opts.locale` (en-US without one; a locale's
words for TRUE and FALSE are also rejected, ИСТИНА and ЛОЖЬ in ru-RU).

### <a id="xlcalc-validatenames"></a>`validateNames`

```ts
function validateNames(names: NameInput[], ctx: NamesContext): NameIssue[]
```

The issues of a workbook's names, found by the same check a workbook gets (`prepareNames`), without the workbook.

### <a id="xlcalc-validatesheetname"></a>`validateSheetName`

```ts
function validateSheetName(name: string, opts?: NameOptions): SheetNameCheck
```

Whether Excel accepts a sheet name, by the rules a workbook's sheets are checked against: 1 to 31 characters, none of
`: \ / ? * [ ]`, no apostrophe at either end, not `History` in any case. A rejection's message is in `opts.locale`'s
language, English without one; `at` is the position of the character Excel rejects.

### <a id="xlcalc-valuecomparator"></a>`valueComparator`

```ts
function valueComparator(opts?: CompareValuesOptions): (a: CellValue, b: CellValue) => number
```

[`compareValues`](#xlcalc-comparevalues) with its options checked once, for `Array.prototype.sort` over many values: the order of two cells'
values, which it accepts as [`CellValue`](#cellvalue)s without checking them (`values.sort(valueComparator({ descending: true }))`).

### <a id="xlcalc-viewportof"></a>`viewportOf`

```ts
function viewportOf(saved: BookSnapshot, sheet: string, range: string, opts?: SavedViewportOptions): SavedViewport
```

A viewport over a saved workbook and its result, as [`Session.getViewport`](#session) returns it over a session of that workbook,
without calculating its formulas: the result's values, text and color, and the workbook's formats, styles and
layout; with `appearance`, its conditional formats calculated over the result's values.

## xlcalc/a1

Described under an entry above: [`cellAddress`](#xlcalc-celladdress), [`colLetters`](#xlcalc-colletters), [`colNumber`](#xlcalc-colnumber), [`linesKey`](#xlcalc-lineskey), [`MAX_COLS`](#xlcalc-max_cols), [`MAX_ROWS`](#xlcalc-max_rows), [`mergedArea`](#xlcalc-mergedarea), [`nameKey`](#xlcalc-namekey), [`outlineEdit`](#xlcalc-outlineedit), [`outlineGroups`](#xlcalc-outlinegroups), [`parseCell`](#xlcalc-parsecell), [`parseLines`](#xlcalc-parselines), [`parseRange`](#xlcalc-parserange), [`patchLines`](#xlcalc-patchlines), [`quoteSheetIfNeeded`](#xlcalc-quotesheetifneeded), [`rangeAddress`](#xlcalc-rangeaddress), [`rectContains`](#xlcalc-rectcontains), [`rectsMeet`](#xlcalc-rectsmeet), [`sameName`](#xlcalc-samename).

## xlcalc/text

Described under an entry above: [`callAt`](#xlcalc-callat), [`columnWidthFromChars`](#xlcalc-columnwidthfromchars), [`columnWidthFromPoints`](#xlcalc-columnwidthfrompoints), [`columnWidthToChars`](#xlcalc-columnwidthtochars), [`columnWidthToPoints`](#xlcalc-columnwidthtopoints), [`compareValues`](#xlcalc-comparevalues), [`DEFAULT_ROW_HEIGHT`](#xlcalc-default_row_height), [`escapeColumn`](#xlcalc-escapecolumn), [`findMatcher`](#xlcalc-findmatcher), [`findReplacer`](#xlcalc-findreplacer), [`formatClipboard`](#xlcalc-formatclipboard), [`FUNCTION_CATEGORIES`](#xlcalc-function_categories), [`MAX_OUTLINE`](#xlcalc-max_outline), [`MAX_ROW_HEIGHT`](#xlcalc-max_row_height), [`nameAt`](#xlcalc-nameat), [`numberFormatCategory`](#xlcalc-numberformatcategory), [`numberFormatChoices`](#xlcalc-numberformatchoices), [`parseClipboard`](#xlcalc-parseclipboard), [`quoteSheetIfNeeded`](#xlcalc-quotesheetifneeded), [`refAt`](#xlcalc-refat), [`scanReferences`](#xlcalc-scanreferences), [`stepDecimals`](#xlcalc-stepdecimals), [`valueComparator`](#xlcalc-valuecomparator).

## xlcalc/entry

### <a id="xlcalcentry-entrytext"></a>`entryText`

```ts
function entryText(value: CellValue, opts?: EntryTextOptions): string
```

The text a cell holding `value` opens with in edit mode, as Excel's formula bar displays it, in `opts.locale` and
according to the cell's `opts.numberFormat`: a number to 15 significant digits with the locale's decimal separator
(`1234.5`, `1234,5`), a percentage cell's number as a percentage (`12%`), a date or time cell's value in the locale's
short date and long time formats (`1/31/2025  6:00:00 PM`, `31.01.2025  18:00:00`), a boolean or an error in the
locale's words, and text with an apostrophe where typing it would otherwise be parsed as something else (`'007`).
Parsing the text as edited in place gives back the value ([`parseEntry`](#xlcalcentry-parseentry) with `inPlace: true`), but for a negative
percentage, which Excel's in-place editor reads as a formula (`-5%` is `=-5%%`).

### <a id="xlcalcentry-needsapostrophe"></a>`needsApostrophe`

```ts
function needsApostrophe(text: string, opts?: EntryTextOptions): boolean
```

Whether text typed into a cell would be parsed as something other than that text, so the cell needs an apostrophe
before it to keep it as text: a number, a date, a boolean, an error or a formula in the locale (`007`, `1.5` in
ru-RU, `=A1`), a number in a percentage cell, text that begins with an apostrophe, or empty text.

### <a id="xlcalcentry-parseentry"></a>`parseEntry`

```ts
function parseEntry(text: string, opts?: EntryOptions): EntryResult
```

Text typed into a cell, as Excel interprets it in the workbook's locale: the cell's new input, with the cell's style,
meta, note and link and the format Excel leaves it, or a formula Excel would reject. Empty text clears the value and
keeps the rest. Typed over the cell, or edited in place with `opts.inPlace`.

## xlcalc/xlsx

Types: see [Types](#types).

### <a id="xlcalcxlsx-fromlegacy"></a>`fromLegacy`

```ts
function fromLegacy(formula: string, ctx?: LegacyContext): string
```

A legacy formula (text without `=`, as a file stores it, with the prefixes of newer functions removed) as the dynamic
array formula Excel 365 displays for it, in canonical text: `A1:A10*2` becomes `@A1:A10*2`, `INDEX(A1:C3,0,1)`
becomes `@INDEX(A1:C3,0,1)`, `SUM(A1:A10)` stays. Excel puts `@` before a name depending on what the name's formula
returns, so the workbook's names are needed for formulas that use names. A formula the engine cannot parse is
returned as is. The API function (`strict-args.ts`), from `xlcalc/xlsx`.

### <a id="xlcalcxlsx-readxlsx"></a>`readXlsx`

```ts
function readXlsx(data: ArrayBuffer | Uint8Array<ArrayBufferLike>, opts?: ReadXlsxOptions): Promise<XlsxRead>
```

Reads an xlsx file (the file's bytes) into a workbook, the values Excel stored, and what the workbook could not hold,
with messages in `opts.locale` (en-US by default).

### <a id="xlcalcxlsx-writexlsx"></a>`writeXlsx`

```ts
function writeXlsx(snapshot: XlsxSnapshot, opts?: WriteXlsxOptions): Promise<XlsxWritten>
```

Writes a workbook as an xlsx file: its formulas with the values from `snapshot.result` that `opts.values` requests,
or only its values (`opts.mode: 'values'`).

## xlcalc/xlsx/keep

Types: see [Types](#types).

### <a id="xlcalcxlsxkeep-keptrefusal"></a>`KeptRefusal`

```ts
class KeptRefusal extends Error
```

[`writeKept`](#xlcalcxlsxkeep-writekept)'s refusal: the change would leave a part of the file the model does not hold pointing at the wrong cells; `code` says why.

### <a id="xlcalcxlsxkeep-openkept"></a>`openKept`

```ts
function openKept(data: ArrayBuffer | Uint8Array<ArrayBufferLike>, opts?: ReadXlsxOptions): Promise<XlsxRead>
```

Reads an xlsx file into a workbook as [`readXlsx`](#xlcalcxlsx-readxlsx) does, its carriers added to the names for [`writeKept`](#xlcalcxlsxkeep-writekept).

### <a id="xlcalcxlsxkeep-writekept"></a>`writeKept`

```ts
function writeKept(original: ArrayBuffer | Uint8Array<ArrayBufferLike>, snapshot: XlsxSnapshot, opts?: WriteKeptOptions): Promise<XlsxWritten>
```

Writes a book over the xlsx file it was read from ([`openKept`](#xlcalcxlsxkeep-openkept)): what the book changed is written, the rest is the
file's as it was, the parts the model does not have among it. Refuses ([`KeptRefusal`](#xlcalcxlsxkeep-keptrefusal), its message in the language of
`opts.locale`) a change that would leave references the model does not hold pointing at other cells.

## xlcalc/worker

Described under an entry above: [`CalcError`](#xlcalc-calcerror), [`MAX_FORMULA_DEPTH`](#xlcalc-max_formula_depth), [`MAX_FORMULA_LENGTH`](#xlcalc-max_formula_length).

### <a id="xlcalcworker-createcalcpool"></a>`createCalcPool`

```ts
function createCalcPool(opts?: CalcPoolOptions): Calc
```

A pool of threads. In Node it runs `worker_threads` with `maxHeapMb` as each thread's heap; in a browser it runs Web
Workers, started by `spawn` or from the thread module next to this one, without a memory limit.

## xlcalc/function-text

### <a id="xlcalcfunction-text-functiondescription"></a>`functionDescription`

```ts
function functionDescription(fn: string, locale?: string | Locale): string | null
```

What a function does, looked up by its English name ([`FunctionSignature.fn`](#functionsignature), [`CallAt.fn`](#callat)) in any case: one line in
the locale's language when the locale has descriptions (`{ ...ruRU, ...ruRUFunctionText }`), otherwise in en-US; null
for a name outside the whitelist.

## xlcalc/function-text/ru-RU

### <a id="xlcalcfunction-textru-ru-rurufunctiontext"></a>`ruRUFunctionText`

```ts
const ruRUFunctionText: Required<Pick<Locale, "argumentWords" | "functionText">>
```

## xlcalc/locales

### <a id="xlcalclocales-loadlocale"></a>`loadLocale`

```ts
function loadLocale(tag: string): Promise<Locale>
```

Loads one of the package's locales by its tag (`'ru-RU'`), as `import { ruRU } from 'xlcalc/locales/ru-RU'` and
`registerLocale(ruRU)` do: a workbook in that locale is taken from then on, and a pool's threads load it too. A locale
loaded already is given back without importing its module again; a tag that is not one of the package's locales is an
[`InputError`](#xlcalc-inputerror) (tags are compared as written).

## xlcalc/locales/ru-RU

### <a id="xlcalclocalesru-ru-ruru"></a>`ruRU`

```ts
const ruRU: Locale
```

ru-RU, for [`registerLocale`](#xlcalc-registerlocale), or to pass as is to [`toLocal`](#xlcalc-tolocal) and the other functions that work without a workbook.

## xlcalc/regex

### <a id="xlcalcregex-pcreengine"></a>`pcreEngine`

```ts
const pcreEngine: RegexLevel
```

The PCRE2 engine: Excel's syntax and match limit, its steps counted on the calculation's budget.

### <a id="xlcalcregex-pcresyntax"></a>`pcreSyntax`

```ts
const pcreSyntax: RegexLevel
```

Excel's syntax by JavaScript's regular expressions: fast, without Excel's match limit, a heavy pattern not stopped.

## xlcalc/contract

Described under an entry above: [`MAX_FORMULA_DEPTH`](#xlcalc-max_formula_depth), [`MAX_FORMULA_LENGTH`](#xlcalc-max_formula_length).

## Types

### `AbortSignalLike`

What a calculation needs from an AbortSignal; the platform's `AbortSignal` qualifies.

```ts
interface AbortSignalLike {
  readonly aborted: boolean
  readonly reason?: unknown
  addEventListener(type: 'abort', listener: () => void, options?: { once?: boolean }): void
  removeEventListener(type: 'abort', listener: () => void): void
}
```

### `AddressMove`

How a structure operation moved a workbook's addresses ([`EditResult.moved`](#editresult)). `lines`: the rows or columns from `at`
onward (a row number, or a column number with A as 1) move by `count`: down or right when lines were inserted before
`at`, up or left when `-count` lines starting at `at` were deleted. `range`: the cells of `from` moved so that its
top-left cell is at `to`, replacing what was there. A sheet renamed, deleted, inserted or moved: `index` is its
position among the sheets, from 0.

```ts
type AddressMove =
  | { kind: 'lines'; sheet: string; axis: 'rows' | 'columns'; at: number; count: number }
  | { kind: 'range'; from: RangeRef; to: CellRef }
  | { kind: 'renameSheet'; sheet: string; name: string }
  | { kind: 'deleteSheet'; sheet: string }
  | { kind: 'insertSheet'; sheet: string; index: number }
  | { kind: 'moveSheet'; sheet: string; index: number }
```

### `AlignmentStyle`

A cell's alignment ([`CellStyle.alignment`](#cellstyle)), like xlsx `<alignment>`: horizontal and vertical alignment, text wrapping, indent, orientation.

```ts
interface AlignmentStyle {
  horizontal?: 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed'
  vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed'
  wrap?: boolean
  /** Indent level, 0 to 250. */
  indent?: number
  /**
   * The text's orientation (Format Cells > Alignment > Orientation, xlsx `textRotation`): degrees from -90 to 90, up for
   * positive ones, or `vertical`, its letters one above the other. AutoFit of rows and columns takes the turned text's
   * room as Excel for Mac does (`fitRows`, `fitColumns`), and a number typed into a turned cell widens no column.
   */
  rotation?: number | 'vertical'
}
```

### `AsyncComputeOptions`

Options of [`compute`](#xlcalc-compute) on a pool or on [`createInlineCalc`](#xlcalc-createinlinecalc): those of [`compute`](#xlcalc-compute), plus the call's deadline and signal.

```ts
type AsyncComputeOptions = ComputeOptions & CalcOptions
```

### `AsyncEditOptions`

Options of an asynchronous call that changes a session's workbook.

```ts
type AsyncEditOptions = CalcOptions & EditOptions
```

### `AsyncSession`

A session calculated in slices or in a thread: the methods of [`Session`](#session), returning promises. Calls are processed in
order, each after the previous one has finished. A call that fails (a deadline, an abort, a limit) changes nothing.

```ts
interface AsyncSession {
  /** Mixed cell, name and formatting operations applied atomically, with one final recalculation. */
  applyChanges(operations: SessionOperation[], opts?: ChangesOptions & CalcOptions): Promise<ChangesResult>
  /** The changes to dependent values and the errors, without changing the session; use the preview's clock when applying. */
  previewChanges(operations: SessionOperation[], opts?: PreviewChangesOptions & CalcOptions): Promise<ChangesPreview>
  /** Reads multiple ranges at a single revision, with only the requested fields and number of cells; range by range, each range's cells row by row. */
  readCells(ranges: RangeRef[], opts?: ReadCellsOptions & CalcOptions): Promise<ReadCellsResult>
  /** Searches where the workbook lives, literally or like Excel's Find; the output is paged, in Find's order. */
  findCells(ranges: RangeRef[], query: string, opts?: FindCellsOptions & CalcOptions): Promise<ReadCellsResult>
  /** Excel's Replace All, like `Session.replaceInCells`. */
  replaceInCells(ranges: RangeRef[], query: string, replacement: string, opts?: ReplaceOptions & CalcOptions): Promise<ReplaceResult>
  /** Direct formula dependents; source ranges stay ranges. */
  getDependents(ranges: RangeRef[], opts?: DependentsOptions & CalcOptions): Promise<DependentsResult>
  /** Totals over nonblank stored values and spills, without returning their cells. */
  summarizeRanges(ranges: RangeRef[], opts?: SummaryOptions & CalcOptions): Promise<RangeSummary>
  /** The session's clock as of its last finished call. */
  readonly clock: Clock
  /** The session's revision as of its last finished call. */
  readonly revision: number
  setCells(edits: CellEdit[], opts?: AsyncEditOptions): Promise<EditResult>
  replaceCells(edits: CellReplace[], opts?: AsyncEditOptions): Promise<EditResult>
  clearCells(ranges: RangeRef[], what: ClearKind, opts?: AsyncEditOptions): Promise<EditResult>
  patchCells(patches: CellPatch[], opts?: AsyncEditOptions): Promise<EditResult>
  setCellText(sheet: string, cell: string, text: string, opts?: CalcOptions & CellTextOptions): Promise<CellTextResult>
  previewCellText(sheet: string, cell: string, text: string, opts?: CalcOptions & CellTextOptions): Promise<CellTextPreview>
  setCellTexts(entries: CellTextEntry[], opts?: CalcOptions & CellTextOptions): Promise<CellTextsResult>
  previewCellTexts(entries: CellTextEntry[], opts?: CalcOptions & CellTextOptions): Promise<CellTextsPreview>
  pasteText(sheet: string, cell: string, text: string, opts?: AsyncEditOptions): Promise<PasteTextResult>
  previewPasteText(sheet: string, cell: string, text: string, opts?: AsyncEditOptions): Promise<PasteTextPreview>
  fillCellText(sheet: string, range: string, text: string, opts?: CalcOptions & FillTextOptions): Promise<FillTextResult>
  previewFillCellText(sheet: string, range: string, text: string, opts?: CalcOptions & FillTextOptions): Promise<FillTextPreview>
  setNames(names: NameInput[], opts?: AsyncEditOptions): Promise<EditResult>
  setExternalBooks(books: ExternalBook[], opts?: AsyncEditOptions): Promise<EditResult>
  setClock(clock: Clock, opts?: AsyncEditOptions): Promise<EditResult>
  editStructure(op: StructureOp, opts?: AsyncEditOptions): Promise<AsyncStructureResult>
  setStyles(styles: CellStyle[], opts?: AsyncEditOptions): Promise<EditResult>
  setLayout(edits: LayoutEdit[], opts?: AsyncEditOptions): Promise<EditResult>
  /** The session's workbook; from a thread, a copy. The order of its objects' keys is not part of it (`Session.getBook`). */
  getBook(): Promise<Book>
  /** Like `Session.getLayout`; from a thread, a copy, without the cells. */
  getLayout(): Promise<BookLayout>
  getUsedRange(sheet: string): Promise<SheetExtent>
  getAutoFilterRange(sheet: string): Promise<string | null>
  getEnd(sheet: string, cell: string, direction: EndDirection): Promise<string>
  getAutoSum(sheet: string, cell: string): Promise<AutoSumGuess>
  getNamesFromSelection(sheet: string, range: string, labels: SelectionLabels): Promise<NameInput[]>
  getSpecialCells(sheet: string, range: string, kind: SpecialCells): Promise<string[]>
  getWorkbookStatistics(sheet: string): Promise<WorkbookStatistics>
  explain(sheet: string, cell: string, opts?: ExplainOptions & CalcOptions): Promise<Explanation>
  /** Like `Session.getCells`; from a thread, only the rectangles are transferred back, in one call; `signal` and `timeoutMs` stop it as they stop a calculation. */
  getCells(sheet: string, range: string | string[], opts?: CalcOptions): Promise<Record<string, CellView>>
  /** Like `Session.copyText`; the text is prepared in a single read in the session's thread. */
  copyText(sheet: string, range: string, opts?: CopyTextOptions & CalcOptions): Promise<string>
  /** Like `Session.getValues`; from a thread, the rows are returned in one call. */
  getValues(sheet: string, range: string, opts?: ValuesOptions & CalcOptions): Promise<CellValue[][]>
  getCell(sheet: string, cell: string): Promise<CellView | undefined>
  /** Like `Session.getViewport`; from a thread, the viewport is returned in one call. */
  getViewport(sheet: string, range: string, opts?: ViewportOptions): Promise<Viewport>
  getCellIssue(sheet: string, cell: string): Promise<CellIssueDetail | null>
  getValidation(sheet: string, cell: string): Promise<CellValidation | null>
  validateEntry(sheet: string, cell: string, input: CellInput | null, opts?: CalcOptions): Promise<EntryCheck>
  getResult(opts?: CalcOptions): Promise<ComputeResult>
  /** Like `Session.getSnapshot`: an independent copy in one call, so no edit can come between the workbook and its result. */
  getSnapshot(opts?: CalcOptions): Promise<BookSnapshot>
  /** Like `Session.getHash`: from a thread, only the hash is transferred back. */
  getHash(opts?: CalcOptions): Promise<string>
  undo(opts?: AsyncEditOptions): Promise<UndoResult | null>
  redo(opts?: AsyncEditOptions): Promise<UndoResult | null>
  /** The session's history as of its last finished call. */
  readonly history: SessionHistory
  /**
   * Like `Session.subscribe`, synchronously: a listener runs when a call's result has arrived, before the call's promise
   * settles. From a pool thread, a change arrives with the call's result, so it reaches the listeners of the session
   * that made the call.
   */
  subscribe(listener: (change: SessionChange) => void): () => void
  close(): Promise<void>
}
```

### `AsyncSessionOptions`

`signal` and `timeoutMs` apply to opening the session (the workbook's first calculation).

```ts
type AsyncSessionOptions = SessionOptions & CalcOptions
```

### `AsyncStructureResult`

[`AsyncSession.editStructure`](#asyncsession): the same result as [`Session.editStructure`](#session) returns.

```ts
type AsyncStructureResult = SessionStructureResult
```

### `AutoFilter`

A sheet's AutoFilter ([`SheetInput.autoFilter`](#sheetinput)).

```ts
interface AutoFilter {
  /** The filtered cells with the header row first: a range from its top-left cell to its bottom-right one (`A1:D20`), or a cell (`A1`). */
  range: string
}
```

### `AutoSumGuess`

Excel's AutoSum at a cell ([`Session.getAutoSum`](#session)): the range it offers to sum and the formula it enters.

```ts
interface AutoSumGuess {
  /**
   * The range to sum, from the run of numbers nearest above the cell (past the cells right above it without a number)
   * to the cell, else the same to the left (`B2:B9`, `B2` for one), ending at a subtotal (a `SUM` or `SUBTOTAL`
   * formula); when the nearest is a subtotal, the subtotals alone, nearest first (`B9,B5`); null without numbers. A date
   * or a time is no number here.
   */
  range: string | null
  /** `SUM` of the range, without `=` (`SUM(B2:B9)`, `SUM()` without a range). */
  formula: string
  /** The formula as Excel shows it in the workbook's locale (`СУММ(B2:B9)`). */
  formulaLocal: string
}
```

### `AverageFormat`

Numbers above the average of the cells' numbers (below it with `below`, or also equal to it with `equal`), or `stdDev` standard deviations beyond it.

```ts
interface AverageFormat extends ConditionalBase {
  type: 'aboveAverage'
  below?: boolean
  equal?: boolean
  /** 1 to 3. */
  stdDev?: number
  style?: ConditionalStyle
}
```

### `BarLook`

A data bar in a cell: where it is drawn, as fractions of the cell's width from the side where its direction starts.

```ts
interface BarLook {
  /** From 0 to 1; `start` < `end`. */
  start: number
  end: number
  color: Color
  gradient: boolean
  border?: Color
  /** The axis between negative and positive bars, as a fraction of the width, for a rule with negative numbers. */
  axis?: number
  axisColor?: Color
  /** The bars run from right to left: the fractions are measured from the cell's right edge. */
  rightToLeft?: true
}
```

### `BlanksFormat`

The cell is blank (or contains only spaces), is not blank, contains an error, or does not. Blanks are determined by
the value, as Excel displays it (an error is not blank); errors by `formula`, or by Excel's (`ISERROR(A1)`) without
it. Excel writes a formula alongside both kinds of rule.

```ts
interface BlanksFormat extends ConditionalBase {
  type: 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors'
  formula?: string
  style?: ConditionalStyle
}
```

### `Book`

A workbook as returned by the calls that return one ([`Session.getBook`](#session), `getSnapshot`, `getLayout`, [`editStructure`](#xlcalc-editstructure),
[`readXlsx`](#xlcalcxlsx-readxlsx)): a [`BookInput`](#bookinput), which can be passed back to any call that accepts a workbook.

```ts
type Book = BookInput
```

### `BookInput`

A workbook as the engine accepts it: its sheets with their cells, names, styles, locale and date system; what [`compute`](#xlcalc-compute), [`createSession`](#xlcalc-createsession) and [`writeXlsx`](#xlcalcxlsx-writexlsx) accept and [`readXlsx`](#xlcalcxlsx-readxlsx) returns.

```ts
interface BookInput {
  /** The saved active visible sheet, by name; the writer defaults to the first visible sheet. Not part of inputHash. */
  activeSheet?: string
  /** Excel's rules for sheet names: 1 to 31 characters, none of `: \ / ? * [ ]`, no apostrophe at either end; at least one visible sheet. */
  sheets: SheetInput[]
  names?: NameInput[]
  /** The cell styles that cells, rows and columns refer to by index. */
  styles?: CellStyle[]
  /**
   * The locale the workbook is calculated in: the BCP 47 tag of a loaded locale, `LOCALES` (`'en-US'` by default; the
   * package's other locales are loaded with `registerLocale`). It changes calculated values as Excel's regional settings
   * do (text parsed as a number or a date, numbers and booleans converted to text, `TEXT` format codes, currency), but
   * not the formula text: formulas stay invariant, and `toLocal` displays them in a locale. A tag that is not loaded
   * throws an `InputError` that names the module to import.
   */
  locale?: string
  /**
   * The date system, like xlsx `workbookPr date1904`: 1900 (the default) counts days from 1900-01-00, including Excel's
   * phantom 1900-02-29; 1904 counts from 1904-01-01 as serial 0. It changes the serial numbers of typed and calculated
   * dates, TODAY and NOW, date text and how dates are displayed: a 1904 workbook displays a negative date or time with a
   * minus sign (`-12:00`). It is included in both hashes only when it is 1904.
   */
  dateSystem?: DateSystem
  /**
   * Iterative calculation of circular references, like Excel's setting and xlsx `calcPr iterate`. Without it, the cells
   * of a cycle and the cells that depend on it are not calculated and keep their values (`CellInput.held`). With it, the
   * cells of every cycle and the cells that depend on them are calculated from 0, repeatedly, sheet by sheet in order of
   * sheet name (case-insensitively), row by row. Iteration stops after `maxIterations` rounds (100 by default, 1 to
   * 32,767) or when no value changes by `maxChange` or more (0.001 by default). Each calculation starts again from 0, so
   * a session returns what a full calculation returns, and `held` has no effect. Excel instead continues from the values
   * it holds.
   */
  iteration?: { maxIterations?: number; maxChange?: number }
  /**
   * Excel's calculation option (Formulas > Calculation Options, xlsx `calcPr calcMode`): `manual`, or `autoNoTable`,
   * automatic except for data tables; automatic without it. The engine calculates every edit whatever it says: a host
   * that offers manual calculation keeps showing the values it showed until the user calculates. Not part of inputHash.
   */
  calculation?: 'manual' | 'autoNoTable'
  /**
   * The other workbooks that the workbook's formulas refer to (`[1]Sheet1!A1`, `'[1]My Sheet'!A1:B9`, `[1]!Rate`):
   * element `i` is `[i+1]`. References to them use the values stored here, as Excel calculates a file it opens without
   * updating its links; the other workbook itself is never read. The list is included in `inputHash` when it is not
   * empty.
   */
  externalBooks?: ExternalBook[]
  /**
   * The workbook's default font, its Normal style's (in xlsx the first font), in points: what a cell shows whose style
   * gives no font name or size. Without it, Calibri 11. It is kept, read and written, and the engine measures as for
   * Calibri 11 whatever it is: column widths in characters, fitted row heights, `####` in a column too narrow and
   * `entryWidth`, which Excel measures in the default font, so for another font they can differ from Excel's. It is
   * included in `inputHash` when given.
   */
  font?: { name: string; size: number }
}
```

### `BookLayout`

A workbook without the cells of its sheets ([`Session.getLayout`](#session)): its styles, names and settings, and each sheet's layout.

```ts
type BookLayout = Omit<Book, 'sheets'> & { sheets: SheetLayout[] }
```

### `BookSnapshot`

A workbook and its result: a session's, taken at a single point in time ([`Session.getSnapshot`](#session)), or a workbook and
what [`compute`](#xlcalc-compute) returned for it, in the form [`writeXlsx`](#xlcalcxlsx-writexlsx) accepts and [`viewportOf`](#xlcalc-viewportof) reads from a version a host saved.

```ts
interface BookSnapshot {
  book: Book
  result: ComputeResult
}
```

### `BorderLine`

One side of a cell's border; the line styles are those of xlsx.

```ts
interface BorderLine {
  style: 'thin' | 'medium' | 'thick' | 'dotted' | 'dashed' | 'double' | 'hair' | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot'
  color?: Color
}
```

### `CachedValues`

The values a file stores for its workbook's formulas, in the shape of a result: [`readXlsx`](#xlcalcxlsx-readxlsx)'s `cached`, or a result
for the workbook. They are looked up by sheet name and address, and only a cell's `value` is read.

```ts
interface CachedValues {
  sheets: { name: string; cells: Record<string, { value: CellValue }> }[]
}
```

### `Calc`

A calculator: [`createInlineCalc`](#xlcalc-createinlinecalc) (the same thread, in slices) or [`createCalcPool`](#xlcalcworker-createcalcpool) (threads).

```ts
interface Calc {
  compute(book: BookInput, opts?: AsyncComputeOptions): Promise<ComputeResult>
  createSession(book: BookInput, opts?: AsyncSessionOptions): Promise<AsyncSession>
  /** Rejects further calls (`closed`), including waiting and running ones, and closes the sessions. */
  close(): Promise<void>
}
```

### `CalcErrorCode`

Why a call failed ([`CalcError.code`](#xlcalc-calcerror)): an asynchronous call rejects with it, and a synchronous session throws `limit`,
`closed` and `stale_revision`. `soft_timeout`: its `timeoutMs` elapsed, as detected by a check
during the calculation. `aborted`: its `signal` fired. `hard_timeout`: it ran past the pool's `hardTimeoutMs` and its
thread was terminated. `out_of_memory`: its thread reached `maxHeapMb`. `thread_failed`: its thread exited or threw
outside the calculation. `restore_failed`: the session could not be restored on a thread and has been closed. `limit`:
the workbook or the edit exceeds [`CalcLimits`](#calclimits). `queue_full`: `maxQueue` tasks are already waiting. `too_many_sessions`:
the pool already holds `maxSessions` sessions. `closed`: the calculator or the session is closed. `stale_revision`: an
edit's `baseRevision` does not match the session's revision ([`EditOptions`](#editoptions)); nothing was changed.

```ts
type CalcErrorCode =
  | 'soft_timeout'
  | 'aborted'
  | 'hard_timeout'
  | 'out_of_memory'
  | 'thread_failed'
  | 'restore_failed'
  | 'limit'
  | 'queue_full'
  | 'too_many_sessions'
  | 'closed'
  | 'stale_revision'
```

### `CalcLimits`

Size limits for a calculation and for an edit. A workbook that exceeds them is not calculated, and an edit that
exceeds them is not made: the call throws [`CalcError`](#xlcalc-calcerror) with code `limit`, so a result is either exact or absent, and
the values hash never depends on the limits.

```ts
interface CalcLimits {
  /** Formulas in the workbook. */
  formulas?: number
  /** Values in a single range read as a whole, or in a single array a formula creates, checked before they are created. */
  cells?: number
  /**
   * Values a single spill may fill (its stored cells: the cells after a large array's last distinct value are not
   * stored), checked before they are written; `cells` when not given.
   */
  spill?: number
  /**
   * Cells a single session edit writes a value or a formula to (formatting alone is stored in the sheet's `looks`),
   * counted before anything is written: a copy's source cells that hold one, multiplied by the number of tiles in its
   * target; a move's source cells that hold one; the cells of `fillCellText`'s range and of `pasteText`'s rectangle; the
   * cells `replaceInCells` re-enters; the edits of `setCells`, `replaceCells`, `setCellTexts` and `applyChanges`; and
   * their previews. Operations that write no more than the sheet already holds (inserting and deleting rows and columns,
   * `clearCells`) are not limited by it.
   */
  writes?: number
}
```

### `CalcOptions`

Options of one asynchronous call.

```ts
interface CalcOptions {
  /** Stops the call: a waiting call immediately, a running one at its next check. */
  signal?: AbortSignalLike
  /** A soft deadline: milliseconds of calculation, counted from when the call starts calculating (not while it waits). */
  timeoutMs?: number
}
```

### `CalcPoolOptions`

Options of [`createCalcPool`](#xlcalcworker-createcalcpool).

```ts
interface CalcPoolOptions {
  /** Number of threads; 2 by default. */
  threads?: number
  /** Node: the heap of each thread (`resourceLimits.maxOldGenerationSizeMb`); not available in a browser. */
  maxHeapMb?: number
  /** A task that runs longer than this, counted from when it reaches its thread, terminates the thread (`hard_timeout`). */
  hardTimeoutMs?: number
  /** The soft deadline for every call that does not set its own. */
  timeoutMs?: number
  /** Maximum number of tasks waiting for a thread, across all threads; one more is rejected (`queue_full`). */
  maxQueue?: number
  /** Maximum number of open sessions in the pool; one more is rejected (`too_many_sessions`). */
  maxSessions?: number
  /** Maximum number of sessions a thread holds; beyond it, the least recently called session leaves the thread and is restored on its next call. */
  maxSessionsPerThread?: number
  /**
   * The limits for every calculation; a calculation's own `limits` override them field by field, and the other fields
   * still apply. With `maxHeapMb`, `cells` defaults to `maxHeapMb · 2^20 / 64`, and `spill` to the smaller of `cells`
   * and `maxHeapMb · 2^20 / 1024`: a single allocation larger than what remains of a thread's heap (a large array, the
   * tables of a large spill's cells) crashes the whole process, not just the thread.
   */
  limits?: CalcLimits
  /** Number of calls kept in a session's journal before the pool takes a new snapshot of its workbook; 64 by default. */
  snapshotEvery?: number
  /** Milliseconds of calculation between two yields inside a thread, at which a cancellation is checked; 50 by default. */
  sliceMs?: number
  /** Starts a thread in a browser, for a bundler that serves the thread's script from another location. */
  spawn?: () => WorkerLike
}
```

### `CallAt`

The function call containing the caret in formula text being typed ([`callAt`](#xlcalc-callat)): the innermost function whose brackets
enclose the caret, and the argument the caret is in, for a tooltip of the function's arguments ([`functionSignature`](#xlcalc-functionsignature)).

```ts
interface CallAt {
  /** The function's name as typed (`СуммЕсли`). */
  name: string
  /** Its English name in uppercase, as `functionSignature` and stored formulas use it: a locale's name translated (`SUMIF`), any other name as is. */
  fn: string
  /** The argument the caret is in, from 0: the number of list separators before it at the call's own level. */
  index: number
  /** The positions of the call's `(` and of its `)`, or null while the call is still open. */
  open: number
  close: number | null
}
```

### `Cell`

A cell of a [`Book`](#book).

```ts
type Cell = CellInput
```

### `CellChange`

A cell that a call changed ([`EditResult.changes`](#editresult)): its address, its new value and, when they changed, its text, formula, formatting, note or hyperlink.

```ts
interface CellChange extends CellRef, Partial<CellDisplay> {
  value: CellValue
  spill?: string
  spillFill?: CellValue
  spillOf?: string
  /** The cell's formula, when it changed: canonical text, as in `CellResult.formula`; null when the cell no longer has one. */
  formula?: string | null
  /** The cell's own style, when it changed, as in `CellView.style`; null when it no longer has one (it uses its row's or column's). */
  style?: number | null
  /** The cell's own number format, when it changed; null when it no longer has one. */
  numberFormat?: NumberFormat | null
  /** A changed `numberFormat`, unless null, as a format code, as in `CellView.numberFormatCode`. */
  numberFormatCode?: string
  /** The cell's note, when its text or author changed; null when it was removed. */
  note?: CellNote | null
  /** The cell's hyperlink, when its url, location or tooltip changed; null when it was removed. */
  link?: CellLink | null
}
```

### `CellDisplay`

A value as a cell displays it: its text and the color of its format's section.

```ts
interface CellDisplay {
  /**
   * What Excel displays in the cell (VBA's `Range.Text`): the value in its number format and the workbook's locale,
   * with General shortened to fit the column's width, `#` across the column for a number that does not fit, and fills
   * (`*x`) repeated across it; text as is (it overflows into the adjacent cells). Derived from the value, the format, the
   * locale and the width, so it is not part of the values hash.
   */
  text: string
  /** The color of the format's section (`[Red]`, `[Color10]`), `#RRGGBB`; absent when the section has none. */
  color?: Color
}
```

### `CellEdit`

A cell edit ([`Session.setCells`](#session)): the given fields replace the cell's, null removes a field, and omitted fields are
left unchanged, as Excel keeps a cell's formatting when a value is typed over it. A cell has one kind of content: a
new `value` removes the formula, its `array` and its `held`, and a new `formula` removes the value and the `held` (a
formula entered where a circular reference prevents its calculation holds 0). `input: null` clears the content and
the hyperlink and keeps the rest, like Excel's Delete key.

```ts
interface CellEdit extends CellRef {
  input: CellFields | null
}
```

### `CellFields`

Fields of a cell to set ([`Session.setCells`](#session)): a value sets the field, null removes it, and an omitted field is left unchanged.

```ts
type CellFields = { [K in keyof CellInput]?: CellInput[K] | null }
```

### `CellImpact`

A cell that a previewed change would change ([`ChangesPreview`](#changespreview)): its value before and after.

```ts
interface CellImpact extends CellRef {
  before: CellValue
  after: CellValue
  /** The error after the operation, or the circular reference that then prevents the cell's calculation, if any. */
  issue?: CellIssueDetail
}
```

### `CellInput`

An input cell. A date is passed explicitly (`{ value: { date: '2025-01-31' } }`) and becomes an Excel serial number; a
number must be finite (NaN and infinities throw [`InputError`](#xlcalc-inputerror)).

```ts
interface CellInput {
  value?: Scalar | { date: string }
  /**
   * Formula text without `=`; a single leading `=` is removed on input (`=SUM(A1:A2)` becomes `SUM(A1:A2)`, with the
   * same input and `inputHash`; `==1` remains an error). When present, `value` is ignored.
   */
  formula?: string
  /**
   * Number format: an Excel format code as stored in an xlsx file (en-US syntax: `#,##0.00`, `dd.mm.yyyy`, `[Red]`),
   * displayed in the workbook's locale (`CellResult.text`), or the number of a built-in format (`NumberFormat`).
   * Without it, a cell uses its row's format, or else its column's.
   */
  numberFormat?: NumberFormat
  /** An index into the workbook's `styles`. Without it, a cell uses its row's style, or else its column's. */
  style?: number
  /** Arbitrary cell data (for example ids of source rows), passed through untouched. */
  meta?: JsonValue
  /**
   * A legacy array formula (entered with Ctrl+Shift+Enter; in xlsx, `<f t="array" ref="B2:C5">` without dynamic array
   * metadata), set on the top-left cell of its range: the range (`B2:C5`, or just the cell, `B2`). The formula is
   * calculated as an array, without implicit intersection, and fills its range without ever extending past it: a
   * smaller array is repeated when it has one row or one column, otherwise the cells beyond it are #N/A; a larger one
   * is truncated; a single value fills every cell. Nothing spills, and nothing blocks it. `ROW()` and `COLUMN()` return
   * the rows and columns of the range. The other cells of the range hold no input, and a session does not allow
   * changing some of them without the rest (`Session.setCells`). Only with `formula`.
   */
  array?: string
  /** A note on the cell (what Excel used to call a comment). Like a style, it stays with the cell and plays no part in calculation. */
  note?: CellNote
  /** A hyperlink on the cell; the cell displays its value. Like a style, it stays with the cell and plays no part in calculation. */
  link?: CellLink
  /**
   * The value a formula keeps while a circular reference prevents its calculation. In a workbook without `iteration`, a
   * formula in a cycle, or one that depends on a cycle (including through other cells), is not calculated, just as Excel
   * calculates none of them: it displays `held`, or 0 without it (a formula entered there displays 0 in Excel).
   * Elsewhere the field has no effect and is dropped on input, so it counts toward `inputHash` only in such cells. A
   * session maintains it the way Excel treats such cells: a cell that another edit puts in a cycle's path keeps the
   * value it displayed, a formula entered there keeps 0, and a cell no longer in the path is recalculated and keeps
   * nothing; `getBook()` returns it (`docs/cycles.md`). The values stored in a file are loaded as `held` through
   * `cached` (`ComputeOptions.cached`).
   */
  held?: Exclude<CellValue, null>
}
```

### `CellIsFormat`

The cell's value compared with the value of one formula, or two for `between` and `notBetween` (both bounds
inclusive, in either order), the way `=`, `<` and the other operators compare; a blank cell that a formula refers to
counts as 0 there.

```ts
interface CellIsFormat extends ConditionalBase {
  type: 'cellIs'
  operator: 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual' | 'between' | 'notBetween'
  /** Formulas as a cell stores them, without `=`: `100`, `"done"`, `$B$1`. */
  formulas: string[]
  style?: ConditionalStyle
}
```

### `CellIssue`

A cell with an error and its reason: an invalid formula (with the check code and position) or a calculation error; or
a formula stopped by a circular reference (`error: 'CYCLE'`), whose value is the one it keeps ([`CellInput.held`](#cellinput)).

```ts
interface CellIssue extends CellRef {
  error: ErrorCode
  reason: string
  check?: CheckCode
  /** Position in the formula text for `check`. */
  at?: number
  /** The function, upper-case, for the checks `unknown_function` and `unsupported` (`ValidateResult`). */
  fn?: string
}
```

### `CellIssueDetail`

Why a cell holds an error, or why a circular reference stops it ([`Session.getCellIssue`](#session)): its code, reason, check code,
position and function as the result's `diagnostics` reports them, and `origin`, the cell where the error originated, found by
following the cells its reason names (`from 'S'!A2`) to one whose error is its own; in a cycle, the first cell of the
cycle that the path reaches, and likewise for a formula stopped by a cycle, through the cells it reads that the cycle
also stops.

```ts
interface CellIssueDetail {
  error: ErrorCode
  reason: string
  check?: CheckCode
  at?: number
  fn?: string
  origin: CellRef
}
```

### `CellLink`

A hyperlink, with at least one of `url` and `location`. `url`: an address outside the workbook, like Excel's Address
(`https://example.com/`, `mailto:a@example.com`, a file). `location`: a place in the workbook, like Excel's
SubAddress (`'Sheet 2'!B12`, `Data!A1:C3`, a defined name), or, together with `url`, a place within the target (the
part after `#`).

```ts
interface CellLink {
  url?: string
  location?: string
  /** The text Excel displays when the pointer hovers over the cell (ScreenTip). */
  tooltip?: string
}
```

### `CellNote`

A note: its author and its text, lines separated by `\n`.

```ts
interface CellNote {
  author?: string
  text: string
  /** Shown all the time (Excel's Show/Hide Note, xlsx `<x:Visible/>`), not only while the pointer rests on its cell. */
  shown?: boolean
}
```

### `CellPatch`

A formatting change to a cell (`cell`) or a range (`range`: `B2:D9`, entire columns `B:D`, entire rows `2:5`)
([`Session.patchCells`](#session)): each given field sets the cells' own, null removes it, and an omitted field is left
unchanged; values and formulas stay as they are. A style given by its fields ([`StylePatch`](#stylepatch)) is applied over the style
each cell uses, its own or else its row's or column's, as Excel formats a cell; a style given by its index in the
workbook's `styles` replaces it.

```ts
type CellPatch = { sheet: string; set: PatchFields } & ({ cell: string; range?: never } | { range: string; cell?: never })
```

### `CellPosition`

A cell's row and column, numbered from 1 as in A1 notation: `B12` is row 12, column 2.

```ts
interface CellPosition {
  row: number
  col: number
}
```

### `CellProtection`

A cell's protection ([`CellStyle.protection`](#cellstyle)), like the Protection tab of Excel's Format Cells and xlsx `<protection>`:
on a protected sheet ([`SheetInput.protection`](#sheetinput)) a locked cell takes no edit, and a hidden cell's formula is not
displayed. A cell is locked and not hidden by default, and a style holds only what differs: `{ locked: false }`,
`{ hidden: true }`; `{ protection: { locked: true } }` is the style `{}`. A session does not enforce it: the host does.

```ts
interface CellProtection {
  locked?: boolean
  hidden?: boolean
}
```

### `CellReadField`

A cell field that `readCells` and `findCells` return: `editText` is the text the cell opens with in edit mode, as in a viewport ([`ViewportCell.editText`](#viewportcell)).

```ts
type CellReadField = 'value' | 'text' | 'formula' | 'editText' | 'style' | 'numberFormat' | 'note' | 'link' | 'issue'
```

### `CellRect`

A rectangle on a sheet: rows and columns numbered from 1, bounds inclusive; an entire column runs from row 1 to [`MAX_ROWS`](#xlcalc-max_rows).

```ts
interface CellRect {
  top: number
  left: number
  bottom: number
  right: number
}
```

### `CellRef`

A cell of a workbook: its sheet's name and its A1 address (`B2`).

```ts
interface CellRef {
  sheet: string
  cell: string
}
```

### `CellReplace`

A cell replaced as a whole ([`Session.replaceCells`](#session)): its new input, or null to remove the cell.

```ts
interface CellReplace extends CellRef {
  input: CellInput | null
}
```

### `CellResult`

A cell's value and, by default (`display`), what the cell displays: `text` and `color`, as in [`CellDisplay`](#celldisplay).

```ts
interface CellResult extends Partial<CellDisplay> {
  /** The value: numbers remain numbers whatever the format displays. */
  value: CellValue
  /**
   * Canonical formula text (sheet names always quoted, function names in uppercase), with sheets, tables and defined
   * names spelled as the workbook defines them, as Excel displays them (`ставка*2` with the name `Ставка` becomes
   * `Ставка*2`).
   */
  formula?: string
  /** The cell's own number format, as given in its input. */
  numberFormat?: NumberFormat
  /** The cell's `meta` as given; on the session's thread the workbook's own object, which must not be modified. */
  meta?: JsonValue
  /**
   * A formula whose array spilled: the range it fills on its sheet (`B2:D9`, with this cell at the top left); for a
   * legacy array formula (`CellInput.array`), its range.
   */
  spill?: string
  /**
   * The value of the cells of `spill` that the result does not list: only for an array of more than 65,536
   * values whose cells beyond some row and column all hold the same value (`A:A*2`: 0 below the data). Those
   * cells are then left out of `cells`, and out of the hash's lines except for a single line for the range.
   */
  spillFill?: CellValue
  /**
   * A cell that an array spilled into, or a cell in a legacy array formula's range: the address of the array's formula
   * on the same sheet (an empty element is 0).
   */
  spillOf?: string
}
```

### `CellStyle`

A cell style in the workbook's style table ([`BookInput.styles`](#bookinput)), like an xlsx cell format without its number format
(which is the cell's `numberFormat`). Styles play no part in calculation or in the values hash.

```ts
interface CellStyle {
  font?: FontStyle
  /** A solid fill. */
  fill?: { color: Color }
  border?: { top?: BorderLine; right?: BorderLine; bottom?: BorderLine; left?: BorderLine }
  alignment?: AlignmentStyle
  protection?: CellProtection
}
```

### `CellTextEntry`

Text typed into a cell, one of several typed at once ([`Session.setCellTexts`](#session)).

```ts
interface CellTextEntry extends CellRef {
  text: string
}
```

### `CellTextOptions`

Options of [`Session.setCellText`](#session) and [`Session.setCellTexts`](#session).

```ts
interface CellTextOptions extends EditOptions {
  /** False: the width to which a typed number would widen its column is returned (`width`) but not applied; true by default, as Excel widens the column. */
  widen?: boolean
  /**
   * True: a value that a cell's data validation rule does not allow, whatever the rule's style, is rejected
   * (`CellValidationRefusal`) and nothing changes, so that a host can display Excel's alert before the value is written
   * and call again without the option when the user keeps the value (Yes on a warning, OK on an information alert).
   * False by default: the value is written and returned with its alert.
   */
  refuseInvalid?: boolean
  /**
   * True: the text was edited within the cell's existing text (F2, a double-click, the formula bar) rather than typed
   * over the cell. Excel's percent entry then adds `%` only to a number that begins with a sign: in a cell formatted
   * `0%`, `6%` is 0.06, `6` is 6 (displayed as 600%) and `-5` is -0.05. False by default.
   */
  inPlace?: boolean
}
```

### `CellTextPreview`

What [`Session.setCellText`](#session) would do ([`Session.previewCellText`](#session)): the cell's input, its value, the width and the alert, without changing anything.

```ts
type CellTextPreview = { ok: true; input: CellInput | null; value: CellValue; width?: number; validation?: ValidationAlert } | CellTextRefused | CellValidationRefusal
```

### `CellTextRefusal`

Why [`Session.setCellText`](#session) rejected the text: a formula Excel would reject ([`CheckCode`](#checkcode)) or one that refers to a workbook not among the links (`external_ref`), part of an array formula, or a formula in a table's header cell.

```ts
type CellTextRefusal = 'syntax' | 'empty' | 'too_long' | 'too_deep' | 'external_ref' | 'array' | 'header'
```

### `CellTextRefused`

Text that [`Session.setCellText`](#session) rejected, and the correction Excel offers for the formula, as in [`parseEntry`](#xlcalcentry-parseentry)'s `correction`.

```ts
type CellTextRefused = Refusal<CellTextRefusal> & { correction?: string }
```

### `CellTextResult`

Text typed into a cell ([`Session.setCellText`](#session)): the edit's result, with the width the column was widened to (`width`)
and the data validation rule the value violates (`validation`); or why the text was rejected, with the position in
the text, or, with `refuseInvalid`, the rule the value violates.

```ts
type CellTextResult = (EditResult & { ok: true; width?: number; validation?: ValidationAlert }) | CellTextRefused | CellValidationRefusal
```

### `CellTextsPreview`

What [`Session.setCellTexts`](#session) would do ([`Session.previewCellTexts`](#session)): each cell's input and value, the widths and the alerts, without changing anything.

```ts
type CellTextsPreview = { ok: true; cells: PreviewCell[]; widths?: ColumnWidth[]; validations?: CellValidationAlert[] } | CellTextsRefusal
```

### `CellTextsRefusal`

Why [`Session.setCellTexts`](#session) wrote none of the texts: the first one rejected, as `setCellText` rejects it, with its
cell; with `refuseInvalid`, the cell whose value violates the strictest data validation rule once all are written
(stop, then warning, then information; the first cell of that style).

```ts
type CellTextsRefusal = (CellTextRefused & CellRef) | CellValidationRefusal
```

### `CellTextsResult`

Texts typed into several cells at once ([`Session.setCellTexts`](#session)): the edit's result, with the widths the columns were
widened to (`widths`) and the cells whose values their data validation rules do not allow (`validations`); or the
first rejected text, with its cell, and nothing written.

```ts
type CellTextsResult = (EditResult & { ok: true; widths?: ColumnWidth[]; validations?: CellValidationAlert[] }) | CellTextsRefusal
```

### `CellValidation`

A cell's data validation ([`Session.getValidation`](#session)): the rule that covers it, the rule's formulas as they apply at the
cell (relative references shifted from the rule's first cell), and, for a list, the items its dropdown offers.

```ts
interface CellValidation {
  rule: DataValidation
  formula1?: string
  formula2?: string
  /**
   * A list's items in order, calculated in the workbook's current state: a literal list's items as written (`text`, with
   * `value` the same text), or the cells of the reference its formula returns, each with its value and the text it
   * displays in its number format (blank cells after the last value are left out). An error when the formula returns
   * neither a reference nor a list. `input` is what Excel writes into the cell when the item is picked from the
   * dropdown, for `setCells`, which keeps the cell's format: a cell's value as is (`001` stays text, a number displayed
   * as `50%` is 0.5), a literal item parsed as typed in en-US, the locale in which a file stores the list (`1/2` is the
   * 2nd of January, `d-mmm`), or null for a blank cell. Excel checks no rule on a picked item.
   */
  list?: { value: CellValue; text: string; input: CellInput | null }[] | { error: ErrorCode }
}
```

### `CellValidationAlert`

A cell whose value its data validation rule does not allow ([`Session.setCellTexts`](#session)), with the rule's alert.

```ts
type CellValidationAlert = CellRef & ValidationAlert
```

### `CellValidationRefusal`

Typed text rejected because a cell's data validation rule does not allow its value ([`CellTextOptions.refuseInvalid`](#celltextoptions)): the cell, the rule, and its alert's style and message.

```ts
type CellValidationRefusal = ValidationRefusal & CellRef
```

### `CellValue`

A calculated value. Dates are returned as serial numbers, and numbers are rounded to 15 significant digits.

```ts
type CellValue = Scalar | { error: ErrorCode }
```

### `CellView`

A cell returned by [`Session.getCells`](#session): like [`CellResult`](#cellresult), plus the cell's own `style` (an index into the workbook's
`styles`), `note` and `link`, and the `appearance` that its sheet's conditional formats give it (when they set a
number format, `text` and `color` display the value in that format).

```ts
interface CellView extends CellResult {
  style?: number
  note?: CellNote
  link?: CellLink
  appearance?: ConditionalAppearance
  /** On the top-left cell of a legacy array formula, its range (`CellInput.array`); `spill` is also that range. */
  array?: string
  /** `numberFormat` as a format code, as `ViewportCell.numberFormat` returns it: a built-in format as its code in the workbook's locale (14 is `dd.mm.yyyy` in ru-RU). */
  numberFormatCode?: string
}
```

### `ChangesOptions`

An explicit clock makes a preview and its application reproducible, including in a live session.

```ts
interface ChangesOptions extends EditOptions { clock?: Clock }
```

### `ChangesPreview`

A complete calculation with bounded output; the session, its revision and its history are unchanged.

```ts
interface ChangesPreview extends ChangesResult {
  impacts: CellImpact[]
  totalChanges: number
  truncated: boolean
}
```

### `ChangesResult`

What `applyChanges` returns: the edit's result and the clock it calculated with.

```ts
interface ChangesResult extends EditResult { clock: Clock }
```

### `ChangingMethod`

The calls of a session that change its workbook, by name ([`SessionChange.method`](#sessionchange)).

```ts
type ChangingMethod =
  | 'applyChanges'
  | 'replaceInCells'
  | 'setCells'
  | 'replaceCells'
  | 'clearCells'
  | 'patchCells'
  | 'setCellText'
  | 'setCellTexts'
  | 'pasteText'
  | 'fillCellText'
  | 'setNames'
  | 'setExternalBooks'
  | 'setClock'
  | 'editStructure'
  | 'setStyles'
  | 'setLayout'
  | 'undo'
  | 'redo'
```

### `CheckCode`

Formula check codes. `unsupported`: an external function called with arguments that are not literals.
`unknown_name` also covers tables: an unknown table or column, or a structured reference without a table name outside
a table.

```ts
type CheckCode =
  | 'empty'
  | 'syntax'
  | 'too_long'
  | 'too_deep'
  | 'external_ref'
  | 'unknown_function'
  | 'unknown_sheet'
  | 'unknown_name'
  | 'unsupported'
```

### `ClearKind`

What [`Session.clearCells`](#session) clears: the content and hyperlinks (Excel's ClearContents), the formatting (ClearFormats), or both and everything else (Clear).

```ts
type ClearKind = 'contents' | 'formats' | 'all'
```

### `Clock`

The point in time and the random seed from which TODAY, NOW, RAND and RANDBETWEEN are calculated: ISO 8601 with Z or
an offset (`2026-10-04T12:00:00Z`), or a date alone, read as midnight UTC.

```ts
interface Clock {
  now: string
  seed: number
}
```

### `Color`

A color, `#RRGGBB`.

```ts
type Color = string
```

### `ColorScaleFormat`

Each number gets a color between those of the points on either side of it; other cells get none.

```ts
interface ColorScaleFormat extends ConditionalBase {
  type: 'colorScale'
  /** Two or three points, from the lowest. */
  points: { value: ConditionalValue; color: Color }[]
}
```

### `ColumnInput`

Columns of a sheet, like xlsx `<col>`: the width, and the style and number format for their cells that have none of their own.

```ts
interface ColumnInput {
  /**
   * In widths of the default font's digit, plus the cell padding, as xlsx stores it, 0 to 255. Without it, a column has
   * Excel's default width (8 digits plus padding; Excel for Windows displays it as 8.43).
   */
  width?: number
  /**
   * The width was set by Excel, by AutoFit or by widening the column for a number typed into it, rather than manually
   * (xlsx `bestFit`): a number typed into the column widens it again (`entryWidth`), as in a column of the default width.
   */
  bestFit?: boolean
  style?: number
  numberFormat?: NumberFormat
  /** Hidden (xlsx `hidden`); a column of width 0 is also hidden. SUBTOTAL and AGGREGATE still include the cells of hidden columns, as Excel does. */
  hidden?: boolean
  /** The outline level (Excel's Group, xlsx `outlineLevel`), 1 to 7; a column without one is not grouped. */
  outlineLevel?: number
  /** The collapsed marker on a group's summary column (xlsx `collapsed`), like `RowInput.collapsed`. */
  collapsed?: boolean
}
```

### `ColumnsFit`

Excel's AutoFit of columns `C` or `B:D` (a double-click on a column's border): each column with content to display
gets the width of its widest cell as Excel for Mac fits it ([`fitColumns`](#xlcalc-fitcolumns)), is marked `bestFit` and is unhidden if it
was hidden; a column with nothing to display keeps its width. The widths are calculated from the values and layout
left by the preceding edits (`docs/formats.md`, "Fitting columns").

```ts
interface ColumnsFit {
  sheet: string
  columns: string
  fit: true
}
```

### `ColumnWidth`

The width a column was widened to ([`Session.setCellTexts`](#session)): its sheet, its letters (`B`) and the width, as in [`ColumnInput.width`](#columninput).

```ts
interface ColumnWidth {
  sheet: string
  column: string
  width: number
}
```

### `CompareValuesOptions`

Options of [`compareValues`](#xlcalc-comparevalues) and [`valueComparator`](#xlcalc-valuecomparator).

```ts
interface CompareValuesOptions {
  /** Largest first; blanks stay last. */
  descending?: boolean
  /** A loaded locale's tag (`LOCALES`), or a locale object, loaded or not, whose text sort order is used (`ruRU`, Latin before Cyrillic, as Excel for Windows sorts); en-US's order without one. */
  locale?: string | Locale
}
```

### `ComputeOptions`

Options of [`compute`](#xlcalc-compute) and of a session: the clock, how formulas are executed, limits and the values a file stored.

```ts
interface ComputeOptions {
  /**
   * The clock for TODAY, NOW, RAND and RANDBETWEEN. By default, the current time and a random seed;
   * pass the `clock` of a result to reproduce that result exactly.
   */
  clock?: Clock
  externals?: Externals
  /** `closures` by default. */
  execution?: Execution
  /**
   * Each cell's displayed text and color (`CellResult.text`, `color`); true by default, false calculates values only.
   * `readCells` and `findCells` still return `text` when their fields ask for it.
   */
  display?: boolean
  /** None by default. */
  limits?: CalcLimits
  /**
   * The values stored in the workbook's file (`readXlsx`'s `cached`), used the way Excel opens a file with a circular
   * reference: a formula stopped by the cycle (`CellInput.held`) that has no `held` of its own keeps the value stored for
   * it, or empty text when the file stores none. Nothing else is read from them.
   */
  cached?: CachedValues
}
```

### `ComputeResult`

What [`compute`](#xlcalc-compute) returns: each sheet's cells with their values and displayed text, the issues, the values' `hash`, `inputHash` and the clock it calculated with.

```ts
interface ComputeResult {
  sheets: { name: string; cells: Record<string, CellResult> }[]
  /**
   * Every cell with an error and every formula stopped by a circular reference (`CellIssue`), sheet by sheet in the
   * workbook's order and row by row, then every name with a problem, in the order of the workbook's names.
   */
  diagnostics: Diagnostic[]
  /** Cells that depend on volatile functions (directly or through other cells). */
  volatile: CellRef[]
  /** The clock the workbook was calculated with (for a live session, the latest one). */
  clock: Clock
  /** sha256 of the canonical text of the values (including spilled cells) together with the engine configuration version. */
  hash: string
  /**
   * sha256 of the input (sheets, cells, names, and styles, columns, rows, merged areas, frozen panes, conditional formats,
   * data validation and outline settings when there are any). The order of the sheets counts, that of cells and names does
   * not. A workbook saved to xlsx and read back usually gets another `inputHash` (formulas in canonical text, dates as
   * numbers, styles renumbered): compare `hash` to tell whether its values changed.
   */
  inputHash: string
  configVersion: string
  ms: number
}
```

### `ConditionalAppearance`

The effect of a sheet's conditional formats on a cell ([`CellView.appearance`](#cellview)), based on the cells' current values. A
field is set only where the cell's display differs from what its own style and value produce.

```ts
interface ConditionalAppearance {
  /** The styles of the rules whose conditions are true, merged in order (the first rule's field wins); a color scale's color is in `fill`. */
  style?: ConditionalStyle
  bar?: BarLook
  icon?: Icon
  /** A data bar or an icon set hides the value: the cell displays only its bar or icon. */
  hideValue?: true
}
```

### `ConditionalBase`

What all kinds of conditional format have in common ([`ConditionalFormat`](#conditionalformat)): the cells it applies to and Excel's Stop If True.

```ts
interface ConditionalBase {
  /**
   * The cells it applies to: rectangles on the sheet (`B2:B20`, `D2`, entire columns `F:F`, entire rows `3:5`), like
   * Excel's "Applies to" box and xlsx `sqref`. Its formulas are written for the top-left cell of the rectangle enclosing
   * them: a reference without `$` shifts from there to each cell, as in a copied formula.
   */
  ranges: string[]
  /** Excel's Stop If True: the rules after this one do not apply to a cell for which this rule's condition is true. */
  stopIfTrue?: boolean
}
```

### `ConditionalFormat`

A conditional format: the cells it applies to, a condition, and the formatting of a cell that meets it. The rules
whose conditions are true for a cell apply in the sheet's order: a style field set by an earlier rule wins, and
`stopIfTrue` skips the rules after one whose condition is true. Its formulas are stored text, like a cell's: a
leading `=` is removed.

```ts
type ConditionalFormat =
  | CellIsFormat
  | ExpressionFormat
  | TextFormat
  | BlanksFormat
  | TimePeriodFormat
  | TopFormat
  | AverageFormat
  | DuplicatesFormat
  | ColorScaleFormat
  | DataBarFormat
  | IconSetFormat
```

### `ConditionalStyle`

The formatting a rule applies to the cells that meet its condition, like an xlsx differential format (`dxf`): each field it sets replaces the cell's own.

```ts
interface ConditionalStyle {
  /** `bold: false` removes bold from a cell that has it, and so on. */
  font?: Pick<FontStyle, 'bold' | 'italic' | 'underline' | 'strike' | 'color'>
  fill?: { color: Color }
  border?: CellStyle['border']
  /** The number format used to display the cell's value. */
  numberFormat?: NumberFormat
}
```

### `ConditionalValue`

A threshold of a color scale, a data bar or an icon set, over the numbers in the cells the rule applies to. `min`
and `max`: the lowest and the highest number; `num`: the number `value` returns; `percent`: `value` percent of the
way from the lowest to the highest; `percentile`: the `value`-th percentile of the numbers; `formula`: the result of
`value`. A data bar's `autoMin` and `autoMax` are Excel 2010's automatic bounds: 0, or the lowest number when it is
below 0 (the highest when it is above 0).

```ts
interface ConditionalValue {
  type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula' | 'autoMin' | 'autoMax'
  /** A formula as a cell stores it (`10`, `$B$1*2`), for `num`, `percent`, `percentile` and `formula`. */
  value?: string
}
```

### `CopyContext`

What [`copyFormula`](#xlcalc-copyformula) reads from a workbook: its sheet names and their tables; a [`BookInput`](#bookinput) qualifies.

```ts
interface CopyContext {
  sheets: { name: string; tables?: TableInput[] }[]
}
```

### `CopyTextOptions`

Options of [`Session.copyText`](#session), which copies a selection as plain text.

```ts
interface CopyTextOptions {
  /** A trailing CRLF, as on Windows; false for Mac. True by default. */
  trailingNewline?: boolean
  /** Maximum number of visible cells to copy, 20,000 by default. A larger copy is rejected, never truncated. */
  maxCells?: number
  /** Maximum length of the encoded text, 16 Mi characters by default, checked before its rows are assembled. */
  maxChars?: number
}
```

### `DataBarFormat`

Each number gets a bar across its cell, its length proportional to where the number lies between `min` and `max`.
The fields from `gradient` onward are Excel 2010's (xlsx `x14:dataBar`).

```ts
interface DataBarFormat extends ConditionalBase {
  type: 'dataBar'
  min: ConditionalValue
  max: ConditionalValue
  color: Color
  /** False: the cell displays the bar without its value. */
  showValue?: boolean
  /** The shortest and the longest bar as a percentage of the cell's width: 0 and 100 by default, 10 and 90 for an Excel 2007 bar. */
  minLength?: number
  maxLength?: number
  /** A gradient fill (true by default) or a solid one. */
  gradient?: boolean
  /** The bar's border color; no border without it. */
  border?: Color
  /** The direction of the bars; by default the sheet's direction (left to right). */
  direction?: 'leftToRight' | 'rightToLeft'
  /** The fill color of bars for negative numbers; the bar's color without it. */
  negativeColor?: Color
  /** The border color of those bars; `border` without it. */
  negativeBorder?: Color
  /** Where the axis between negative and positive bars is: placed according to the numbers (the default), at the middle of the cell, or none (negative bars run in the same direction as positive ones). */
  axis?: 'middle' | 'none'
  /** The axis color; black without it. */
  axisColor?: Color
  /** An Excel 2007 bar (from a file without the Excel 2010 part of the rule): a gradient without a border, negative numbers as shorter bars, no axis; the fields from `gradient` onward do not apply. */
  excel2007?: boolean
}
```

### `DataValidation`

A data validation rule, like xlsx `<dataValidation>`. Its formulas are stored formula text without `=` (en-US, like a
cell's; a leading `=` is removed), relative to the top-left cell of the rectangle enclosing the ranges in `range`,
whatever their order, as in a conditional format: a rule over `B2:B9` with `formula1: 'A2>0'` checks `A5>0` for B5,
and one over `H3:H6 J1:J3` with `$A1>2` checks `$A3>2` for H3. A literal list is a text constant of
comma-separated items (`'"Yes,No,Maybe"'`), as a file stores it in every locale; a list from cells is a reference
(`'$D$2:$D$9'`, `"'Lists'!$A$1:$A$20"`), a defined name, or a formula that returns a reference.

```ts
interface DataValidation {
  /** The cells: space-separated ranges (`B2:B20 D2`, `C:C`), like xlsx `sqref`; no other rule on the sheet overlaps them. */
  range: string
  type: ValidationType
  /** For whole, decimal, date, time and textLength: how the value is compared; `between` by default. */
  operator?: ValidationOperator
  /** The value, the lower bound, the list or the formula of a custom rule; every type except `any` has one. */
  formula1?: string
  /** The upper bound for `between` and `notBetween`. */
  formula2?: string
  /**
   * A blank cell passes, and so does any value while a bound or a list refers to a blank cell (Excel's Ignore blank,
   * xlsx `allowBlank`); true by default, as when Excel creates a rule.
   */
  allowBlank?: boolean
  /** A list displays its dropdown in the cell; true by default (false is xlsx `showDropDown="1"`). */
  dropdown?: boolean
  /** The input message Excel displays while the cell is selected: its title (Excel allows 32 characters) and text (255). */
  promptTitle?: string
  prompt?: string
  /** Whether the input message is displayed; true by default, as when Excel creates a rule. */
  showPrompt?: boolean
  /** The alert for a value the rule does not allow: stop (the default), warning or information, and its title (32) and text (225). */
  errorStyle?: 'stop' | 'warning' | 'information'
  errorTitle?: string
  error?: string
  /** Whether the alert is displayed; true by default, as when Excel creates a rule. Without it, Excel accepts any typed value. */
  showError?: boolean
}
```

### `DateSystemOptions`

The workbook's date system, as in [`BookInput.dateSystem`](#bookinput): 1900 without one.

```ts
interface DateSystemOptions {
  dateSystem?: DateSystem
}
```

### `DependentCell`

A cell whose formula refers to the ranges passed to `getDependents`, with the references that point to them.

```ts
interface DependentCell extends CellRef {
  /** Matching source references remain rectangles, never expanded into individual cells. */
  references: RangeRef[]
  referencesTruncated?: boolean
}
```

### `DependentsOptions`

Options of `getDependents`: the page options and how many references to return for each cell.

```ts
interface DependentsOptions extends PageOptions {
  /** Maximum matching source references per dependent, 64 by default, at most 1000. */
  maxReferences?: number
}
```

### `DependentsResult`

What `getDependents` returns: a page of the cells whose formulas refer to the ranges, sheet by sheet in the workbook's order, each sheet's by rows and in a row by columns.

```ts
interface DependentsResult extends QueryPage { dependents: DependentCell[] }
```

### `Diagnostic`

What a result reports about its workbook, by `kind`: a cell with an error and its reason (`cell`, a [`CellIssue`](#cellissue)), or a
name that the workbook dropped or keeps with an error (`name`, a [`NameIssue`](#nameissue)).

```ts
type Diagnostic = ({ kind: 'cell' } & CellIssue) | ({ kind: 'name' } & NameIssue)
```

### `DisplayOptions`

Options of [`displayValue`](#xlcalc-displayvalue): where the value is displayed.

```ts
interface DisplayOptions {
  /** The number format; General without it. */
  numberFormat?: NumberFormat
  /** A loaded locale's tag (`LOCALES`), or a locale object, loaded or not; en-US without one. */
  locale?: string | Locale
  /**
   * The column's width as in `ColumnInput.width` (the sum of several columns' widths for a merged area). Without it, the
   * width is not limited: General displays all the digits it would in a wide column, and no number turns into `#`.
   */
  width?: number
  /** The date system used to display dates, as in `BookInput.dateSystem`; 1900 without it. */
  dateSystem?: DateSystem
}
```

### `DuplicatesFormat`

Values that occur more than once among the cells (text compared case-insensitively), or only once.

```ts
interface DuplicatesFormat extends ConditionalBase {
  type: 'duplicateValues' | 'uniqueValues'
  style?: ConditionalStyle
}
```

### `EditOptions`

Options of a call that changes a session's workbook.

```ts
interface EditOptions {
  /**
   * The revision of the workbook that the caller's view reflects (`EditResult.revision`, `Session.revision`): when the
   * session's revision differs, the call changes nothing and throws `CalcError` with code `stale_revision`, so that a
   * host never applies an edit on top of one it has not seen.
   */
  baseRevision?: number
}
```

### `EditResult`

What a call that changes a session's workbook returns: the session's revision after the call, the cells it changed,
and anything else a host displaying the workbook needs to read again.

```ts
interface EditResult {
  /** The session's revision after the call: every call that changes the workbook increments it, including `undo` and `redo`. */
  revision: number
  /**
   * The cells whose value, displayed text, formula, formatting, note or hyperlink changed, as `CellChange`; a cell whose
   * value changed because of another cell is listed with its value only. After a structure operation, the cells whose
   * value or formula differs from that of the cell previously at their old position (`moved`). A rectangle of more than
   * 65,536 cells without an entry of their own whose formatting changed (`patchCells` or `clearCells` over formatted
   * ranges, a merge of entire columns, their undo and redo) is in `repaint` instead of here, spilled values included.
   */
  changes: CellChange[]
  /** The sheets whose layout changed (rows, columns, merged areas, tables, conditional formats, validation…): a host reads `getLayout()` again. */
  layout?: string[]
  /** The workbook's style table, when the call changed it. */
  styles?: CellStyle[]
  /** How addresses moved, in order, after a structure operation or its undo: a host moves its viewport accordingly. */
  moved?: AddressMove[]
  /**
   * Rectangles where cells may look different without being in `changes`: the cells covered by the workbook's
   * conditional formats, after a call that changed a value or the rules (a rule's formatting depends on values wherever
   * its formulas point), and a rectangle of more than 65,536 cells whose formatting changed (`patchCells`, `clearCells`, a
   * structure operation, their undo and redo), which a host reads again where it overlaps its viewport.
   */
  repaint?: RangeRef[]
  /**
   * The rows whose height the call fitted as AutoFit does, in runs of rows of one height (`5`, `2:9`), `height` null for
   * the default: the rows of cells that typed text wraps, the rows of cells whose style `patchCells` changed (a font
   * makes a row taller or lower), and the rows a copy laid its looks on (`editStructure`'s `copyRange` without `paste`,
   * with `all` or `formats`), where the height is the default or one Excel fitted and the row shows. Their sheets are
   * in `layout`.
   */
  fitted?: FittedRows[]
}
```

### `EndDirection`

The direction in which Control plus an arrow key moves from a cell ([`Session.getEnd`](#session)).

```ts
type EndDirection = 'up' | 'down' | 'left' | 'right'
```

### `EntryCheck`

Whether a value satisfies a cell's rule ([`Session.validateEntry`](#session)); when it does not, the rule whose alert Excel
displays, its style (`errorStyle`, stop by default) and its message: the rule's `error`, or else the text of Excel's
alert in the workbook's locale.

```ts
type EntryCheck = { ok: true } | ValidationRefusal
```

### `EntryOptions`

Options of [`parseEntry`](#xlcalcentry-parseentry): the workbook that the text is typed into, and the target cell.

```ts
interface EntryOptions {
  /**
   * The workbook's locale, a loaded locale's tag (`LOCALES`) or a locale object, loaded or not; en-US without one:
   * numbers, dates, booleans, errors and formulas are parsed as Excel parses them in that locale.
   */
  locale?: string | Locale
  /** The workbook's date system, as in `BookInput.dateSystem`; 1900 without it. */
  dateSystem?: DateSystem
  /**
   * The cell's entry context: its effective `numberFormat`, `style`, `meta`, `note` and `link`; other fields are not read.
   * A `CellInput`, or a `CellView` from `getCell` or `getCells`, carries only the cell's own format. To take a row's or
   * column's format into account, pass the resolved cell from `getViewport`, or set `numberFormat` to the effective
   * format yourself. The new input keeps the cell's style, meta, note and link, and its format where Excel keeps it.
   * The returned input stores that format on the cell; `setCellText` keeps an inherited format on its row or column.
   */
  cell?: { numberFormat?: NumberFormat; style?: number; meta?: JsonValue; note?: CellNote; link?: CellLink } | null
  /** The point in time whose year completes a date typed without a year (`15.3` in ru-RU, `3/15` in en-US), in ISO format as in `Clock.now`; the current time without it. */
  now?: string
  /**
   * The workbook's links to other workbooks (`BookInput.externalBooks`): a formula's reference to one of them, typed as
   * Excel displays it (`=[Book2.xlsx]Sheet1!A1`, `='C:\data\[Book2.xlsx]Sheet1'!A1`, `=Book2.xlsx!Rate`) or by number,
   * is matched by file name and replaced by its number (`'[1]Sheet1'!A1`); a reference to a workbook not among the links
   * is rejected (`external_ref`). Without them, such a formula is kept as typed.
   */
  externalBooks?: ExternalBook[]
  /**
   * True: the text was edited within the cell's existing text (F2, a double-click, the formula bar), as `entryText`
   * writes it, rather than typed over the cell; Excel's percent entry then adds `%` only to a number that begins with a
   * sign: in a cell formatted `0%`, `6%` is 0.06 and `6` is 6 (600%). False by default, as `Session.setCellText` has it.
   */
  inPlace?: boolean
}
```

### `EntryResult`

Text typed into a cell, as Excel interprets it ([`parseEntry`](#xlcalcentry-parseentry)): the cell's new input, `null` when nothing is left of the
cell; or a formula Excel would reject, with a check code (`syntax`, `empty`, `too_long`, `too_deep`, and `external_ref`
for a workbook not among `opts.externalBooks`), a message in the language of `opts.locale` (English for en-US) and the
position in the typed text. When a formula is typed with
brackets, quotes or array braces left open at its end, Excel closes them on Return if the formula begins with a
function, has no `)` and leaves only function brackets open (`=SUM(1,2` becomes `=SUM(1,2)`); any other such formula
it rejects and offers to correct, provided the closed formula parses and every function in it is known: `correction`
is that text, as typed with everything left open closed (`=(1+2` offers `=(1+2)`). Pasted text is not closed.

```ts
type EntryResult = { ok: true; input: CellInput | null } | (Refusal<CheckCode> & { at: number; correction?: string })
```

### `EntryTextOptions`

Options of [`entryText`](#xlcalcentry-entrytext) and [`needsApostrophe`](#xlcalcentry-needsapostrophe): the cell's number format (its row's or column's when it has none of its
own), the workbook's locale and date system, and, for text, the point in time whose year completes a date typed
without one, as [`parseEntry`](#xlcalcentry-parseentry) parses it.

```ts
interface EntryTextOptions {
  numberFormat?: NumberFormat
  /** A loaded locale's tag (`LOCALES`) or a locale object; en-US without one. */
  locale?: string | Locale
  dateSystem?: DateSystem
  /** ISO, as in `Clock.now`; the current time without it. */
  now?: string
}
```

### `EntryWidthOptions`

Options of [`entryWidth`](#xlcalc-entrywidth): the format used to display the value (the cell's format after the entry), the workbook's
locale and date system, and the `width` (the default without one) and `bestFit` of the column typed into. A cell
whose text is turned ([`AlignmentStyle.rotation`](#alignmentstyle)) widens no column, so a host does not call it for one.

```ts
interface EntryWidthOptions extends DisplayOptions {
  /** The column's width was fitted by Excel (`ColumnInput.bestFit`), so typing widens it as it widens the default width. */
  bestFit?: boolean
}
```

### `ErrorCode`

`CALC`: Excel's #CALC!, an empty array (`FILTER` that matches nothing) or text longer than 32,767 characters
(`CONCAT`). `SPILL`: an array that cannot spill from its formula's cell (the range is not empty, overlaps a merged
area or runs off the sheet). `NULL`: Excel's #NULL!, ranges that do not intersect under the intersection operator
(`A1:A3 C1:C3`). `ERROR`: a formula that cannot be calculated (for example, one that does not parse).

```ts
type ErrorCode = 'DIV0' | 'NA' | 'NAME' | 'NUM' | 'REF' | 'VALUE' | 'NULL' | 'CYCLE' | 'SPILL' | 'CALC' | 'ERROR'
```

### `Execution`

How formulas are executed; values, issues and hashes are the same either way. `closures` (the default) builds each
formula as a tree of JavaScript closures and generates no code, so it works under a Content Security
Policy without 'unsafe-eval'. `eval` generates JavaScript source and compiles it with `new Function`, which
such a policy blocks (calculation then throws an `EvalError`).

```ts
type Execution = 'closures' | 'eval'
```

### `ExplainNode`

A node of an explanation ([`Explanation.root`](#explanation)): a cell or a range, its formula and value, and the nodes it reads.

```ts
interface ExplainNode {
  /** The sheet and A1 cell or range of a node in this workbook, for navigation; absent for an external workbook. */
  location?: RangeRef
  /** `'Sheet'!B3` or `'Sheet'!D:D`. */
  ref: string
  value?: CellValue
  formula?: string
  /** A cell that an array spilled into: its formula's cell (`'Sheet'!B2`), which is the node's child. */
  spillOf?: string
  /** For a range: the number of non-empty cells. */
  cells?: number
  /**
   * A cell with an error, or a formula stopped by a circular reference: the cell where the error originated
   * (`'Sheet'!A2`, or the node's own `ref` when it originated there), as in `CellIssueDetail.origin`. The path to it is
   * in the tree even beyond `depth`, and it passes through a range by way of the cell the error came from.
   */
  origin?: string
  children: ExplainNode[]
}
```

### `ExplainOptions`

Options of `explain`: how deep the tree goes and how many nodes it holds.

```ts
interface ExplainOptions {
  /** Maximum depth of precedents, 3 by default. */
  depth?: number
  /** Node limit, 200 by default. */
  maxNodes?: number
  /** Expand ranges into their non-empty cells (by default a range is a single node). */
  expand?: boolean
  /** How many cells beyond `depth` the path from an error to the cell where it originated may continue, 64 by default. */
  errorDepth?: number
}
```

### `Explanation`

What `explain` returns: why a cell holds its value, as a tree of what it depends on and as text.

```ts
interface Explanation {
  root: ExplainNode
  nodes: number
  truncated: boolean
  /**
   * The tree as indented text (for logs, an LLM agent or a user) in the workbook's locale: in ru-RU its words, numbers,
   * formulas, booleans and errors are in Russian (`План!B2 = 480  ← ОКРУГЛ(B1*1,2;0)`), and a sheet name is quoted only
   * where Excel quotes it; the nodes keep the stored text. An error is written as the workbook's cells display it
   * (`#DIV/0!`), with the cell where it originated (`(error from S!A2)`).
   */
  text: string
}
```

### `ExpressionFormat`

A formula whose result is true, as IF evaluates its condition: TRUE, a number other than 0, or the text TRUE.

```ts
interface ExpressionFormat extends ConditionalBase {
  type: 'expression'
  formula: string
  style?: ConditionalStyle
}
```

### `ExternalBook`

An external workbook, as stored alongside the formulas that refer to it (xlsx `externalLinkN.xml`): its location, and
the values Excel saved from it when it last read it.

```ts
interface ExternalBook {
  /** Where the link points, as stored in the file: a file name, a relative or absolute path, or a `file:` or web address. */
  target: string
  /** Its sheets in the order the file lists them. */
  sheets: ExternalSheet[]
  /** Its defined names that the workbook's formulas use: workbook-level ones, or ones local to one of its sheets. */
  names?: ExternalName[]
}
```

### `ExternalName`

A defined name of an external workbook.

```ts
interface ExternalName {
  name: string
  /** What it refers to, as stored in the file without `=`: a reference into its workbook (`Sheet1!$B$1`) or `#REF!`. */
  ref: string
  /** The index in `sheets` of the sheet the name is local to; a workbook-level name has none. */
  sheet?: number
}
```

### `Externals`

External functions (host functions evaluated outside the workbook): their values are resolved in
advance, in a single request, and passed to the calculation. The key of a call is its canonical text
(`BALANCE("2025-01")`); arguments must be literals. They are options of the calculation, not part of the
workbook, so `inputHash` does not include them.

```ts
interface Externals {
  functions: string[]
  values?: Record<string, CellValue>
}
```

### `ExternalSheet`

A sheet of an external workbook and the values stored from it.

```ts
interface ExternalSheet {
  name: string
  /**
   * The stored values by A1 address, null for a cell stored as empty. A cell missing from them is blank, or `#REF!` with
   * `refreshError`. Without `cells`, nothing from the sheet is stored, and every cell on it is `#REF!`.
   */
  cells?: Record<string, ExternalValue | null>
  /**
   * Excel's flag that the last update of the sheet's values failed (`refreshError`): a cell missing from them is `#REF!`
   * instead of blank. Excel sets it on every sheet of a closed workbook when a formula referring to that workbook is
   * entered, and so does a session.
   */
  refreshError?: true
}
```

### `ExternalValue`

A stored value of an external workbook's cell.

```ts
type ExternalValue = number | string | boolean | { error: ErrorCode }
```

### `FillTextOptions`

Options of [`Session.fillCellText`](#session).

```ts
interface FillTextOptions extends CellTextOptions {
  /** The selection's active cell, which the text is typed into: a cell of the range (`B2`); its top-left cell by default. */
  active?: string
}
```

### `FillTextPreview`

What [`Session.fillCellText`](#session) would do ([`Session.previewFillCellText`](#session)): each cell's input and value, the widths and the alert, without changing anything.

```ts
type FillTextPreview = { ok: true; cells: PreviewCell[]; widths?: ColumnWidth[]; validation?: ValidationAlert } | CellTextsRefusal
```

### `FillTextResult`

Text entered into every cell of a range at once ([`Session.fillCellText`](#session), Excel's Control+Return): the edit's result,
with the widths the columns were widened to (`widths`) and the alert of the active cell's data validation rule
(`validation`); or why the text was rejected, with the cell, and nothing written.

```ts
type FillTextResult = (EditResult & { ok: true; widths?: ColumnWidth[]; validation?: ValidationAlert }) | CellTextsRefusal
```

### `FillType`

What a fill takes from its source ([`StructureOp`](#structureop) `fillRange`), like Excel's `XlAutoFillType`. `default` is dragging
the fill handle (`xlFillDefault`). Each line's source is split into runs of cells of the same kind, skipping empty
cells: numbers continue by their step or trend; dates by days, by their step, or by months when they share a day of
the month or are month ends; times by hours or by their step; text with a number at the start (followed by a space)
or at the end continues by the number, including quarters and en-US ordinals; names of days and months continue
through their lists; dates, times and text without a constant step are copied, as are plain text, booleans and
formulas. A single number continues by 1 unless the source holds only plain numbers. `series` (`xlFillSeries`,
dragging with Ctrl over a number) always continues a single number by 1. `copy` (`xlFillCopy`) copies every cell.

```ts
type FillType = 'default' | 'series' | 'copy'
```

### `FindCellsOptions`

Options of `findCells`: what to search and how, in addition to those of `readCells`.

```ts
interface FindCellsOptions extends ReadCellsOptions {
  /**
   * What to search: the text a cell displays (`text`, the default; Excel's Look in: Values), the text it opens with in
   * edit mode (`edit`: a formula in the workbook's locale, a number in its format, text as is, without an apostrophe;
   * Excel's Look in: Formulas), its stored formula (`formula`) or its value (`value`: the same in every locale, a number
   * as JavaScript writes it, `TRUE` or `FALSE`, an error by its name such as `#DIV/0!`). Regular expressions are not
   * supported.
   */
  in?: 'text' | 'edit' | 'formula' | 'value'
  /** Excel's Match case; case-insensitive by default (`straße` matches `STRAẞE`, `i` does not match `İ`). */
  caseSensitive?: boolean
  /** The query is interpreted as Excel's Find interprets it: `*` matches any text, `?` one character (a character outside the BMP takes two), and `~` before a character makes it literal. Without it, the query is literal. */
  wildcards?: boolean
  /** Excel's Match entire cell contents: the entire text must match the query, not just a part of it. */
  wholeCell?: boolean
  /** Excel's Search: By Columns: each range's cells are searched column by column; row by row without it. */
  byColumns?: boolean
  /**
   * Cells in hidden rows and columns are skipped as Excel's Find skips them: always when searching displayed text
   * (`text`, Look in: Values), and in every kind of search on a filtered sheet, where the cells are visited block by
   * block, one block of cells between hidden rows and columns at a time (`docs/live-books.md`). Without it, every cell
   * in the ranges is searched.
   */
  skipHidden?: boolean
  /**
   * The cell after which Find starts, like `Range.Find`'s `After` (a user's Find Next from the active cell): cells are
   * visited in Find's order starting from the one after it, wrapping around to the first cell of the ranges and on to
   * this cell, which comes last; it must be a cell within the ranges.
   */
  after?: CellRef
}
```

### `FindTextOptions`

Options of [`findMatcher`](#xlcalc-findmatcher) and [`findReplacer`](#xlcalc-findreplacer): Find's options, as `findCells` accepts them, applied to a host's own texts.

```ts
type FindTextOptions = Pick<FindCellsOptions, 'caseSensitive' | 'wildcards' | 'wholeCell'>
```

### `FittedRows`

Rows given one height by a call ([`EditResult.fitted`](#editresult)).

```ts
interface FittedRows {
  sheet: string
  rows: string
  height: number | null
}
```

### `FontStyle`

A font; whatever a style omits comes from the workbook's default font.

```ts
interface FontStyle {
  name?: string
  /** In points, 1 to 409. */
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: 'single' | 'double'
  strike?: boolean
  color?: Color
}
```

### `FormatCategory`

A number format category, as the Number Format box on Excel's Home tab displays it for the active cell
([`numberFormatCategory`](#xlcalc-numberformatcategory)): the category of Excel's Format Cells dialog that produces the format, or `custom` for a
format that none of them produces.

```ts
type FormatCategory = 'general' | 'number' | 'currency' | 'accounting' | 'date' | 'time' | 'percentage' | 'fraction' | 'scientific' | 'text' | 'special' | 'custom'
```

### `FormatChoice`

An item of the Number Format list in a locale: its key, its name in the list (`Short Date`), the format it applies (a
built-in format by number, or a format code as Excel writes it to a file), and that format's code in the locale, in
en-US syntax as a cell's `numberFormat` accepts it.

```ts
interface FormatChoice {
  key: FormatChoiceKey
  name: string
  format: NumberFormat
  code: string
}
```

### `FormatChoiceKey`

An item in the Number Format box's list on the Home tab, in list order ([`numberFormatChoices`](#xlcalc-numberformatchoices)).

```ts
type FormatChoiceKey = 'general' | 'number' | 'currency' | 'accounting' | 'shortDate' | 'longDate' | 'time' | 'percentage' | 'fraction' | 'scientific' | 'text'
```

### `FormulaReference`

A reference found in formula text: to cells, rows or columns of a sheet, or to a table.

```ts
type FormulaReference = GridReference | TableReference
```

### `FrozenPanes`

A sheet's frozen panes ([`SheetInput.frozen`](#sheetinput), Excel's Freeze Panes): the rows above the split and the columns to the left of it.

```ts
interface FrozenPanes {
  /** 1 to 1,048,575. */
  rows?: number
  /** 1 to 16,383. */
  columns?: number
}
```

### `FunctionArgument`

An argument in a function's signature: its name in the locale and whether a call can omit it.

```ts
interface FunctionArgument {
  name: string
  optional: boolean
}
```

### `FunctionCategory`

A function category, as Excel's Insert Function dialog lists them; `compatibility` holds the old names Excel keeps
alongside their newer versions (`NORMDIST` alongside `NORM.DIST`).

```ts
type FunctionCategory = 'financial' | 'date' | 'math' | 'statistical' | 'lookup' | 'database' | 'text' | 'logical' | 'information' | 'engineering' | 'compatibility'
```

### `FunctionSignature`

A function's signature for an argument tooltip while a formula is being typed ([`functionSignature`](#xlcalc-functionsignature)), as Excel's
tooltip lists the arguments: `SUM(number1, [number2], ...)`, in ru-RU `СУММ(число1;[число2];...)`.

```ts
interface FunctionSignature {
  /** The function's name in the locale (`СУММ`). */
  name: string
  /** Its English name, as stored formulas, `CallAt.fn` and `functionDescription` use it (`SUM`). */
  fn: string
  /** The arguments in the tooltip's order, with repeating ones listed twice (`number1`, `number2`), the second time as optional. */
  args: FunctionArgument[]
  /**
   * How many of the last arguments repeat beyond the list (the tooltip's `...`): 1 for SUM, 2 for SUMIFS, 0 when none
   * do. The argument at index `i` beyond the list is named like the one `repeat` positions before it, with its number
   * incremented.
   */
  repeat: number
  /** The tooltip's text, with the locale's list separator. */
  text: string
  /** The function's category, as Excel's Insert Function dialog lists it (`FUNCTION_CATEGORIES`). */
  category: FunctionCategory
  /** The category's name in the locale (`Math & Trig`, `Математические`). */
  categoryName: string
}
```

### `GridReference`

A reference to cells, entire columns or entire rows, with its rectangle ([`CellRect`](#cellrect)) and its sheet as written
(unquoted): `sheet` alone for a single sheet, `sheet` and `sheetLast` for a reference that spans sheets
(`Jan:Mar!A1`), neither for the formula's own sheet. `spill`: a spill reference (`E2#`), whose rectangle is the
formula's cell.

```ts
interface GridReference extends ReferenceSpan, CellRect {
  kind: 'cells' | 'columns' | 'rows'
  sheet?: string
  sheetLast?: string
  spill?: true
}
```

### `Icon`

An icon from a set: the set and the icon's index in it, from 0, in Excel's order (xlsx `iconId`: for `3Arrows`, 0 is the red down arrow and 2 the green up arrow).

```ts
interface Icon {
  set: IconSetName
  index: number
}
```

### `IconSetFormat`

Each number gets an icon from the set according to the thresholds it reaches.

```ts
interface IconSetFormat extends ConditionalBase {
  type: 'iconSet'
  icons: IconSetName
  /**
   * One threshold per icon of the set, from the lowest; a number gets the icon of the last threshold it reaches (`gte`:
   * at or above it, true by default; false: above it), and always reaches the first one. `icon` substitutes a different
   * icon, `null` no icon (Excel 2010's custom icons).
   */
  points: { value: ConditionalValue; gte?: boolean; icon?: Icon | null }[]
  /** The icons in reverse order: the last one for the lowest numbers. */
  reverse?: boolean
  /** False: the cell displays the icon without its value. */
  showValue?: boolean
}
```

### `IconSetName`

Excel's icon sets by their xlsx names; `3Stars`, `3Triangles` and `5Boxes` are Excel 2010's.

```ts
type IconSetName =
  | '3Arrows'
  | '3ArrowsGray'
  | '3Flags'
  | '3TrafficLights1'
  | '3TrafficLights2'
  | '3Signs'
  | '3Symbols'
  | '3Symbols2'
  | '3Stars'
  | '3Triangles'
  | '4Arrows'
  | '4ArrowsGray'
  | '4RedToBlack'
  | '4Rating'
  | '4TrafficLights'
  | '5Arrows'
  | '5ArrowsGray'
  | '5Rating'
  | '5Quarters'
  | '5Boxes'
```

### `InlineCalcOptions`

Options of [`createInlineCalc`](#xlcalc-createinlinecalc).

```ts
interface InlineCalcOptions {
  /** Milliseconds of calculation between two yields to the event loop; 10 by default. */
  sliceMs?: number
  /** The soft deadline for every call that does not set its own. */
  timeoutMs?: number
  /** The limits for every calculation; a calculation's own `limits` override them field by field, and the other fields still apply. */
  limits?: CalcLimits
}
```

### `JsonValue`

Data a host keeps on a cell ([`CellInput.meta`](#cellinput)), as JSON.

```ts
type JsonValue = Scalar | JsonValue[] | { [key: string]: JsonValue }
```

### `KeptRefusalCode`

Why [`writeKept`](#xlcalcxlsxkeep-writekept) refuses, for a program ([`KeptRefusal.code`](#xlcalcxlsxkeep-keptrefusal)): `file`, a file it does not write over (a Strict Open XML
workbook, a sheet name Excel does not take, a worksheet in other namespaces); `carriers`, a hidden name [`openKept`](#xlcalcxlsxkeep-openkept)
gave the workbook is gone; `unread`, a sheet whose rows, columns or name changed has a part with references writeKept
does not read; `deleted`, cells, a sheet or a table a part of the file refers to or belongs to were deleted; `moved`, a
part would move or be cut by the change, which writeKept does not write; `links`, a link to another workbook was
removed, moved or pointed at another file, or one of a kind other than a workbook (DDE, OLE) changed; `font`, the default font changed; `threads`, the notes of a sheet with comment threads changed; `rules`,
the data validation of a sheet changed and the file has rules of it the workbook could not read (its [`readXlsx`](#xlcalcxlsx-readxlsx) issues
`dataValidation`), which writing them again would lose, or its conditional formats changed and the file has rules of
it that cannot go with them as Excel would (a pivot table's, one read over part of its ranges, one whose formulas
would need making for a new first cell); `layout`, the workbook's carriers were laid out by another version of
openKept or for another file, or their layout's name is gone while carriers stay.

```ts
type KeptRefusalCode = 'file' | 'carriers' | 'layout' | 'unread' | 'deleted' | 'moved' | 'links' | 'font' | 'threads' | 'rules'
```

### `LayoutEdit`

A change of a sheet's layout.

```ts
type LayoutEdit =
  /** Columns `C` or `B:D`; the patch applies to each. Setting or removing the width clears bestFit unless it is given explicitly. */
  | { sheet: string; columns: string; set: LinePatch<ColumnInput> }
  | ColumnsFit
  | RowsFit
  /** Rows `5` or `5:9`; the patch applies to each. Setting or removing the height clears bestFit unless it is given explicitly. */
  | { sheet: string; rows: string; set: LinePatch<RowInput> }
  /** All the merged areas of the sheet. */
  | { sheet: string; merges: string[] }
  /** The sheet's filter mode (`SheetInput.filtered`). */
  | { sheet: string; filtered: boolean }
  /** The sheet's AutoFilter (`SheetInput.autoFilter`); null removes it. */
  | { sheet: string; autoFilter: AutoFilter | null }
  /** The sheet's protection (`SheetInput.protection`); null unprotects it. */
  | { sheet: string; protection: SheetProtection | null }
  /** The sheet's frozen panes (`SheetInput.frozen`); null unfreezes them. */
  | { sheet: string; frozen: FrozenPanes | null }
  /** The color of the sheet's tab (`SheetInput.tabColor`); null removes it. */
  | { sheet: string; tabColor: Color | null }
  /** All the conditional formats of the sheet (`SheetInput.conditionalFormats`), in priority order. */
  | { sheet: string; conditionalFormats: ConditionalFormat[] }
  /** All the data validation rules of the sheet (`SheetInput.validations`). */
  | { sheet: string; validations: DataValidation[] }
  /** Where the sheet's summary rows and columns are (`SheetInput.outline`); null for Excel's default, below and to the right. */
  | { sheet: string; outline: OutlineSettings | null }
```

### `LegacyContext`

What [`fromLegacy`](#xlcalcxlsx-fromlegacy) knows about the workbook: its names, and the formula's sheet (whose local names it can see).

```ts
interface LegacyContext {
  names?: NameInput[]
  sheet?: string
}
```

### `LinePatch`

Fields to change on columns or rows: a value sets the field, null removes it, and an omitted field is left unchanged.

```ts
type LinePatch<T> = { [K in keyof T]?: T[K] | null }
```

### `LineSpan`

The columns or rows a key of [`SheetInput.columns`](#sheetinput) or `rows` refers to (`C`, `B:D`, `5`, `5:9`), numbered from 1.

```ts
interface LineSpan {
  first: number
  last: number
}
```

### `LiveBook`

Objects for a synchronous session's sheets and cells ([`liveBook`](#xlcalc-livebook)): a cell reads what the session calculated, an
assignment to it is an edit of the session (recalculated, one undo step), and a listener is notified of changes to a
cell, a sheet or the workbook, whoever makes them. One per session: [`liveBook`](#xlcalc-livebook) returns the same object for the same
session.

```ts
interface LiveBook {
  /** The session the objects read and edit. */
  readonly session: Session
  /**
   * The sheet with this name, in any case: the same object for as long as the sheet exists, even if renamed or moved.
   * Throws `InputError` for a name that matches no sheet in the workbook; a sheet deleted and restored by `undo` is a new
   * object.
   */
  sheet(name: string): LiveSheet
  /**
   * Runs `fn` and applies the assignments it makes to cells as a single `applyChanges`: one recalculation, one revision,
   * one undo step, one change notification. Until `fn` returns, the cells read as they were before it started. A batch
   * inside another batch is part of the outer one; when `fn` throws, none of its assignments is applied and the error
   * propagates. Typed text (`LiveCell.type`) cannot be part of a batch, and neither can an `fn` that returns a promise.
   */
  batch(fn: () => void): void
  /** Calls `listener` after each change to the workbook, as `Session.subscribe` does; returns the function that unsubscribes. */
  watch(listener: (change: SessionChange) => void): () => void
}
```

### `LiveCell`

A cell of a [`LiveSheet`](#livesheet), identified by its address: after rows are inserted above it, it refers to the cell now at
that address. Each property reads the session's current state; an assignment edits the session immediately, or as
part of a `batch`, and throws whatever the session's call throws, leaving the workbook unchanged.

```ts
interface LiveCell {
  readonly sheet: LiveSheet
  /** The cell's address, `B12`. */
  readonly address: string
  /** The calculated value, null for an empty cell. */
  get value(): CellValue
  /** Like `setCells` with a `value`: the content is replaced, the formula is removed, the formatting stays. Null clears the content and the hyperlink, like `clearCells` with `contents`. */
  set value(value: Scalar | { date: string })
  /** The text the cell displays (`CellResult.text`); undefined for an empty cell and in a session without `display`. */
  get text(): string | undefined
  /** The formula without `=`, as in `CellResult.formula`; undefined for a cell without one. */
  get formula(): string | undefined
  /** Like `setCells` with a `formula`; null removes it. */
  set formula(formula: string | null)
  /** The cell's own number format, as in `CellView.numberFormat`. */
  get numberFormat(): NumberFormat | undefined
  /** Like `patchCells`: the value and formula stay; null removes the cell's own format. */
  set numberFormat(format: NumberFormat | null)
  /** The cell's own style, an index into the workbook's `styles`, as in `CellView.style`. */
  get style(): number | undefined
  /** Like `patchCells`: a style by its index, or by its fields applied over the style the cell uses; null removes the cell's own style. */
  set style(style: number | StylePatch | null)
  /** Text typed into the cell, as `Session.setCellText` accepts it, with the same result. */
  type(text: string, opts?: CellTextOptions): CellTextResult
  /**
   * Calls `listener` after each change that affects the cell: a change to its value, displayed text, formula or
   * formatting (`EditResult.changes`, including cells changed by recalculation), or a structure operation that moved
   * other cells to its address (`moved`). Returns the function that unsubscribes; after the sheet is deleted, the
   * listener is no longer called.
   */
  watch(listener: (change: SessionChange) => void): () => void
}
```

### `LiveSheet`

A sheet of a [`LiveBook`](#livebook). After the sheet is deleted, every use of the object throws [`InputError`](#xlcalc-inputerror).

```ts
interface LiveSheet {
  /** The sheet's name now: a rename changes it. */
  readonly name: string
  /**
   * The sheet's cells by address, in any case, without `$` (`cells.B12`, `cells.b12`), each created when it is read and
   * the same object while a reference to it is held; any other key (a range, `then`, `toJSON`, a symbol) is undefined.
   */
  readonly cells: { readonly [address: string]: LiveCell }
  /** The cell at `address`, as `cells` returns it; throws `InputError` for anything other than a cell address. */
  cell(address: string): LiveCell
  /**
   * Calls `listener` after each change that affects the sheet: its cells, its layout, a structure operation on it, its
   * renaming or moving, the workbook's style table. Returns the function that unsubscribes; after the sheet is deleted,
   * the listener is no longer called.
   */
  watch(listener: (change: SessionChange) => void): () => void
}
```

### `Locale`

A locale ([`ruRU`](#xlcalclocalesru-ru-ruru) from `xlcalc/locales/ru-RU`): how formulas are written in it and, for a workbook calculated in it, its calculation rules and messages.

```ts
interface Locale {
  /** BCP 47 tag: `en-US`, `ru-RU`. A locale without messages of its own uses those of the loaded locale with its tag (English for any other tag). */
  readonly tag: string
  /** Decimal separator of number literals: `.` in en-US, `,` in ru-RU. */
  readonly decimalSeparator: string
  /** Separator of function arguments, also the union operator: `,` in en-US, `;` in ru-RU. */
  readonly listSeparator: string
  /** Separator of columns in array constants: `,` in en-US (`{1,2;3,4}`). */
  readonly arrayColumnSeparator: string
  /** Separator of rows in array constants: `;` in en-US. */
  readonly arrayRowSeparator: string
  /** The words for TRUE and FALSE. */
  readonly booleans: { readonly TRUE: string; readonly FALSE: string }
  /** Error literals by their invariant text (`'#N/A': '#Н/Д'`); an error not listed keeps its English text. */
  readonly errors: Readonly<Record<string, string>>
  /** Function names by invariant name (`SUM: 'СУММ'`); a function not listed keeps its English name. */
  readonly functions: Readonly<Record<string, string>>
  /**
   * The item specifiers of structured references by their invariant text (`'#Data': '#Данные'`); without them, or for a
   * specifier not listed, the English one is used. A reference's items are separated by the list separator
   * (`[#Данные];[Qty]` in ru-RU).
   */
  readonly tableItems?: Readonly<Partial<Record<TableItem, string>>>
  /**
   * The words used in argument names in function tooltips (`functionSignature`), by their English word (`number:
   * 'число'`, so `number1` becomes `число1`); a word not listed stays in English. The package's locales keep them, and
   * `functionText`, in separate modules (`xlcalc/function-text/ru-RU`), which a host merges into the locale.
   */
  readonly argumentWords?: Readonly<Record<string, string>>
  /**
   * The text of Excel's Insert Function dialog: the category names (`functionSignature`) and what each function does
   * (`functionDescription`), by its English name; anything not listed stays in English.
   */
  readonly functionText?: { readonly categories: Readonly<Partial<Record<FunctionCategory, string>>>; readonly descriptions: Readonly<Record<string, string>> }
  /**
   * The calculation rules as data: number and date text, names, `TEXT` format codes, text sort order, the code page.
   * Every locale a workbook is calculated in has them, except en-US, whose rules are built into the engine's code; a
   * locale without them parses and displays values as en-US does.
   */
  readonly regional?: RegionalData
  /**
   * Messages for users in the locale's language (`messages.ts`): why a formula does not parse, why a structure operation
   * is rejected, the words of `explain`, what was lost from an xlsx file. A message the locale lacks is in English. They
   * are functions, the only part of a locale that a structured clone cannot carry: a pool's thread imports the package's
   * module for the tag itself.
   */
  readonly messages?: Partial<Messages>
}
```

### `LookRange`

A rectangle of a sheet's formatting ([`SheetInput.looks`](#sheetinput)): a cell (`B2`), a range (`B2:D9`), entire columns (`B:D`)
or entire rows (`2:9`), and the style and number format used by its cells that have no entry; with neither field,
its cells have none of their own. The workbooks the engine returns hold the same formatting in a single canonical
form: rows in blocks, the runs shared by a group of rows as rectangles across the group, then the rows of the group
that differ, and its rows without formatting as rectangles with neither field.

```ts
interface LookRange {
  range: string
  /** An index into the workbook's `styles`. */
  style?: number
  numberFormat?: NumberFormat
}
```

### `NameAt`

The name of a function being typed at the caret, in formula text being typed ([`nameAt`](#xlcalc-nameat)), for the Formula AutoComplete
list: where the name begins and ends, and the part of it before the caret.

```ts
interface NameAt {
  from: number
  to: number
  typed: string
}
```

### `NameCheck`

What [`validateName`](#xlcalc-validatename) returns: the name is valid, or why it is not.

```ts
type NameCheck = { ok: true } | Refusal<NameRefusal>
```

### `NameInput`

A defined name of a workbook ([`BookInput.names`](#bookinput)): the name, what it refers to and its scope.

```ts
interface NameInput {
  /**
   * Excel's rules for names, the same as for tables: a letter of any script, `_` or `\` first, then letters, digits,
   * `_`, `.`, `\` and `?`, at most 255 characters; not a cell address (`Q1`) or an R1C1 reference (`R`, `C`, `R2C3`),
   * not TRUE or FALSE, nor the workbook locale's words for them, in any case (ИСТИНА and ЛОЖЬ in ru-RU, which Excel
   * ru-RU rejects; `Ставка`, `НДС_2025`, `АБ12` and `Став?ка` are valid names). Names are compared case-insensitively:
   * `ставка` is the name `Ставка`.
   */
  name: string
  /**
   * What the name refers to: a formula without `=` (a leading one is removed), as Excel's "Refers to" box and xlsx
   * `<definedName>` hold it. A reference with a sheet name (`'Data'!$D$2:$D$3401`, `'Data'!$D:$D`), a constant (`0.2`,
   * `"text"`, `{1,2,3}`), a formula (`'Data'!$A$1*2`, `SUM('Data'!$A:$A)`, `INDEX('Data'!$A$1:$C$9,2,0)`, which returns
   * a reference), or a union of references on one sheet (`'Data'!$A$1:$A$3,'Data'!$C$1:$C$3`). Every reference includes
   * its sheet. A reference without `$` is relative to A1, as a file stores it: used in a cell, it shifts by that cell's
   * offset from A1, wrapping around the sheet's edges (`'Data'!B1` used in C5 is `'Data'!D5`, and `'Data'!A1048576` is
   * the row above). Names in it are global names (a sheet-level name is written as `'Sheet'!name`), as a file stores a
   * name's formula. The formula is calculated where the name is used, as part of the formula that uses it.
   */
  ref: string
  /** The sheet where the name is visible; a local name shadows a global one. Without it the name is global. */
  scope?: string
  /**
   * Hidden (xlsx `hidden="1"`): Excel's Name Manager does not list the name, which a host's list of names should follow.
   * Not part of calculation: formulas use a hidden name like any other.
   */
  hidden?: boolean
}
```

### `NameIssue`

Why a defined name returns no value: the name, its scope (none for a global name) and the reason.

```ts
interface NameIssue {
  name: string
  scope?: string
  reason: string
  /**
   * The workbook has no such name: one Excel would reject, a duplicate, an unknown scope sheet, a formula that does not
   * parse, or a reference without its sheet; a formula that uses the word finds no name. False: the name is in the
   * workbook and returns an error where it is used (an unknown sheet, table or function in its formula).
   */
  dropped: boolean
}
```

### `NameOptions`

The workbook whose rules [`validateName`](#xlcalc-validatename) checks a name against.

```ts
interface NameOptions {
  /** The workbook's locale (`BookInput.locale`): a loaded locale's tag, or a locale object; en-US without one. */
  locale?: string | Locale
}
```

### `NameRefusal`

Why Excel would reject a name ([`validateName`](#xlcalc-validatename)). `chars`: it does not start with a letter of any script, `_` or `\`,
contains a character other than letters, digits, `_`, `.`, `\` and `?`, or is empty or longer than 255 characters.
`reference`: it is parsed as a cell address or an R1C1 reference (`Q1`, `R`, `R2C3`). `boolean`: TRUE or FALSE, or
the workbook locale's word for either, in any case (`ИСТИНА` in ru-RU).

```ts
type NameRefusal = 'chars' | 'reference' | 'boolean'
```

### `NamesContext`

The parts of a workbook that [`validateNames`](#xlcalc-validatenames) checks names against.

```ts
interface NamesContext {
  /** Sheets of the workbook: a name's scope must be one of them, and references in its formula to other sheets return `#REF!`. */
  sheets: string[]
  /** Tables of the workbook with their columns, for structured references in the names' formulas. */
  tables?: { name: string; columns: string[] }[]
  /** The workbook's links to other workbooks, for references to them in the names (`ValidateContext.externalBooks`). */
  externalBooks?: ExternalBook[]
  externalFunctions?: string[]
  /** The workbook's locale (`BookInput.locale`): a loaded locale's tag, or a locale object; en-US without one. */
  locale?: string | Locale
}
```

### `NumberFormat`

A number format: a format code as stored in an xlsx file, or the number of one of Excel's built-in formats, 0 to
49, which is how a file stores the format of a cell that has no code of its own (`numFmtId`). Excel displays some
built-in formats according to the locale: 14 is the locale's short date (`m/d/yyyy` in en-US, `dd.mm.yyyy` in
ru-RU), 5 to 8 and 37 to 44 its currency and number formats, 15 to 22 its other date and time formats. A format code
is displayed as written, regardless of the locale.

```ts
type NumberFormat = string | number
```

### `NumberFormatCategory`

What the Number Format box displays for a format: its category and the category's name in the locale (`Number`, `Числовой`).

```ts
interface NumberFormatCategory {
  category: FormatCategory
  name: string
}
```

### `OutlineEditOp`

An outline operation of [`StructureOp`](#structureop) without its sheet, applied to a single sheet ([`outlineEdit`](#xlcalc-outlineedit)): `{ op: 'group', rows: '5:7' }`.

```ts
type OutlineEditOp = WithoutSheet<Extract<StructureOp, { op: 'group' | 'ungroup' | 'showDetail' | 'hideDetail' | 'showLevels' }>>
```

### `OutlineEditResult`

A sheet's rows and columns after an outline command ([`outlineEdit`](#xlcalc-outlineedit)), each axis in full as [`SheetInput`](#sheetinput) holds it, and
omitted when it has no entries; or why Excel rejects the command, as [`editStructure`](#xlcalc-editstructure) rejects it, in English.

```ts
type OutlineEditResult = { ok: true; rows?: Record<string, RowInput>; columns?: Record<string, ColumnInput> } | Refusal<'outline_deep' | 'not_grouped' | 'detail'>
```

### `OutlineGroup`

A group of rows or columns as Excel draws it ([`outlineGroups`](#xlcalc-outlinegroups)): the longest run of lines from `first` to `last`
(columns by number) at outline level `level` or deeper, so a group of level 2 lies inside one of level 1.

```ts
interface OutlineGroup extends LineSpan {
  level: number
  /**
   * The summary line, which holds the group's button: the line just after the group, below or to the right (just before
   * it with summaries above or to the left, `OutlineSettings`); null when that line is beyond the sheet's edge, where
   * Excel draws no button. Groups that end together share one.
   */
  summary: number | null
  /** Collapsed, with the button showing `+`: its summary line carries the collapsed marker as Excel keeps it (`RowInput.collapsed`); without a summary line, all its lines are hidden. */
  collapsed: boolean
}
```

### `OutlineSettings`

A group of rows has its summary row below it (`summaryBelow`, true by default) or above it; a group of columns, to the right or to the left.

```ts
interface OutlineSettings {
  summaryBelow?: boolean
  summaryRight?: boolean
}
```

### `PageOptions`

Options of a query that returns results in pages (`readCells`, `findCells`, `getDependents`): how many results a page holds and where the previous page stopped.

```ts
interface PageOptions extends EditOptions {
  /** Maximum number of results returned, 1000 by default, at most 10000. */
  maxCells?: number
  /** An opaque continuation of the same query at the same revision; once the session is edited, it is rejected. */
  cursor?: string
  /** Count all matching cells, even after the page is full; false by default. */
  count?: boolean
}
```

### `PaneName`

A pane of a window split in two or four: the first is the top left one, the others lie right of or below a line.

```ts
type PaneName = 'topLeft' | 'topRight' | 'bottomLeft' | 'bottomRight'
```

### `PasteKind`

What a copy pastes of its source ([`StructureOp`](#structureop) `copyRange`), as Excel's Paste Special. `all`, the default, pastes
everything. `values`: each cell's value as computed, a formula's and a spilled cell's too, an error as an error
constant; `formulas`: the formulas as a copy pastes them and the constants. Both leave the target's styles, number
formats, merged areas, notes, hyperlinks, conditional formats and data validation, a blank of the source clears the
target's cell, and both are refused onto a merged area the copy does not bring. `formats`: styles, number formats,
merged areas and conditional formats, the target's contents, notes, hyperlinks and data validation left.
`validation`: the data validation alone. `valuesAndNumberFormats` and `formulasAndNumberFormats`: as `values` and
`formulas` with the copy's number formats, and not refused onto merged areas. `columnWidths`: the widths of the
copy's columns alone. `allExceptBorders`: as `all`, each cell keeping the destination's borders. `notes`: the copy's
notes alone, a cell without one clearing the destination's (Excel's Comments). `allMergingConditionalFormats`: as
`all`, the destination's conditional formats kept whole after the copy's, a rule copied on its own sheet growing
where it makes one rectangle with the pasted cells; with `skipBlanks` it is `all`.
Excel's Format Painter is `formats`; its All Using Source Theme is `all` within one workbook.

```ts
type PasteKind = 'all' | 'values' | 'formulas' | 'formats' | 'validation' | 'valuesAndNumberFormats' | 'formulasAndNumberFormats' | 'columnWidths' | 'allExceptBorders' | 'notes' | 'allMergingConditionalFormats'
```

### `PasteOperation`

Paste Special's operation ([`StructureOp`](#structureop) `copyRange`): each number of the copy is added to, subtracted from,
multiplied into or divided into the destination cell's, a blank counting as 0; a formula on either side makes a
formula of both (`10+(A1*2)`). Text, booleans and errors of the copy leave the destination's cell. It applies to the
kinds that paste contents and is ignored by the others; onto an array spilled over more than one cell or a legacy
array formula, of one cell or its whole range, it is refused (`array`).

```ts
type PasteOperation = 'add' | 'subtract' | 'multiply' | 'divide'
```

### `PasteTextPreview`

What [`Session.pasteText`](#session) would do ([`Session.previewPasteText`](#session)): the range, and the input and value of each of its cells, without changing anything.

```ts
type PasteTextPreview = { ok: true; range: string; cells: PreviewCell[] } | PasteTextRefusal
```

### `PasteTextRefusal`

Why [`Session.pasteText`](#session) pasted nothing, with the cell: part of an array formula, a formula into a table's header cell
(as `setCellText` rejects them), or a rectangle that runs beyond the sheet's last row or column (`off_sheet`).

```ts
type PasteTextRefusal = Refusal<'array' | 'header' | 'off_sheet'> & CellRef
```

### `PasteTextResult`

Text pasted from the clipboard ([`Session.pasteText`](#session)): the edit's result with the range the clipboard's rectangle filled (`B1:D2`); or why nothing was pasted.

```ts
type PasteTextResult = (EditResult & { ok: true; range: string }) | PasteTextRefusal
```

### `PatchFields`

The fields [`CellPatch`](#cellpatch) sets; on entire rows and columns, `meta`, `note` and `link` can only be removed.

```ts
interface PatchFields {
  style?: number | StylePatch | null
  numberFormat?: NumberFormat | null
  meta?: JsonValue | null
  note?: CellNote | null
  link?: CellLink | null
}
```

### `PointAction`

What an arrow key does at a position in formula text typed over a cell (Excel's Enter mode), as [`refAt`](#xlcalc-refat) determines
it: `insert`, a reference to the cell the arrow points to is inserted at `at` (Point mode), preceded by its own `=`
with `equals` (at the start of the text, where Excel starts a new formula: `=A1` becomes `=D3=A1`); `commit`, the text
is entered into the cell and the arrow moves the selection, as after a value; `none`, nothing happens (the text cannot
take a reference there and is not a formula Excel would enter: inside a string, or after a character that does not
occur in the locale's formulas).

```ts
type PointAction = { action: 'insert'; at: number; equals?: true } | { action: 'commit' } | { action: 'none' }
```

### `PreviewCell`

A cell of a preview of text written into cells ([`Session.previewCellTexts`](#session), `previewPasteText`, `previewFillCellText`): its input and its value as the call would leave them.

```ts
interface PreviewCell extends CellRef {
  input: CellInput | null
  value: CellValue
}
```

### `PreviewChangesOptions`

Options of `previewChanges`: those of `applyChanges`, and how many changed cells to return.

```ts
interface PreviewChangesOptions extends ChangesOptions {
  /** Maximum number of changed cells returned, 1000 by default, at most 10000. The calculation itself is still complete. */
  maxChanges?: number
}
```

### `QueryPage`

A page of a query's results: the revision and clock it read, whether more results remain and where to continue.

```ts
interface QueryPage {
  revision: number
  clock: Clock
  truncated: boolean
  cursor?: string
  /** With count: the number matched over the complete query, not just this page. */
  total?: number
}
```

### `RangeRect`

A range's rectangle and what its text refers to: cells (`B2:D9`, `B2`), entire columns (`B:D`) or entire rows (`2:9`).

```ts
interface RangeRect extends CellRect {
  kind: 'cells' | 'columns' | 'rows'
}
```

### `RangeRef`

A rectangle on a sheet: `B2`, `B2:D9`, entire columns `B:D` or entire rows `2:9` (no `$`, no sheet name).

```ts
interface RangeRef {
  sheet: string
  range: string
}
```

### `RangeSummary`

What `summarizeRanges` returns, as Excel's status bar shows it for a selection: how many cells are not empty, how many hold numbers and errors, the sum of the numbers, and the smallest and largest number.

```ts
interface RangeSummary {
  revision: number
  clock: Clock
  cells: number
  numbers: number
  errors: number
  /** Compensated sum; NUM when it overflows the finite number range. */
  sum: number | { error: 'NUM' }
  min: number | null
  max: number | null
}
```

### `ReadCell`

A cell as `readCells` and `findCells` return it: its address and the requested fields ([`CellReadField`](#cellreadfield)), no others.

```ts
interface ReadCell extends CellRef, Partial<Pick<CellView, 'value' | 'text' | 'formula' | 'style' | 'numberFormat' | 'note' | 'link'>> {
  issue?: CellIssueDetail | null
  /** With the field `editText`: the text the cell opens with in edit mode, as in `ViewportCell.editText`; this is what `findCells` searches with `edit`, except for the apostrophe before text that would otherwise be parsed as something else, which the search leaves out. */
  editText?: string
  // the fields of CellView no CellReadField asks for, kept so that host code reading them compiles until a breaking release
  /** @deprecated Never returned: no `CellReadField` asks for it; `getCell` and a viewport give it. */
  color?: CellView['color']
  /** @deprecated Never returned: no `CellReadField` asks for it; `getCell` and a viewport give it. */
  spill?: CellView['spill']
  /** @deprecated Never returned: no `CellReadField` asks for it; `getCell` and a viewport give it. */
  spillFill?: CellView['spillFill']
  /** @deprecated Never returned: no `CellReadField` asks for it; `getCell` and a viewport give it. */
  spillOf?: CellView['spillOf']
  /** @deprecated Never returned: no `CellReadField` asks for it; `getCell` and a viewport give it. */
  meta?: CellView['meta']
  /** @deprecated Never returned: no `CellReadField` asks for it; `getCell` and a viewport give it. */
  array?: CellView['array']
  /** @deprecated Never returned: no `CellReadField` asks for it; `getCell` and a viewport give it. */
  appearance?: CellView['appearance']
  /** @deprecated Never returned: no `CellReadField` asks for it; `getCell` and a viewport give it. */
  numberFormatCode?: CellView['numberFormatCode']
}
```

### `ReadCellsOptions`

Options of `readCells`: the fields to read and which rows and columns.

```ts
interface ReadCellsOptions extends PageOptions {
  /**
   * value, text and formula by default. Address and sheet are always returned. `text` is the displayed text in a session
   * opened with `display: false` too; `['value', 'formula']` reads without formatting.
   */
  fields?: CellReadField[]
  /**
   * Only the cells in these rows (for example, the rows a host's filter displays), as in `ViewportOptions.rows`: their
   * row numbers on the sheet, from 1, in ascending order, each listed once; the cells of each range that fall in these
   * rows, with `total` counting only those. A row outside a range contributes nothing to that range.
   */
  rows?: number[]
}
```

### `ReadCellsResult`

What `readCells` and `findCells` return: a page of cells.

```ts
interface ReadCellsResult extends QueryPage { cells: ReadCell[] }
```

### `ReadXlsxOptions`

Options of [`readXlsx`](#xlcalcxlsx-readxlsx).

```ts
interface ReadXlsxOptions {
  /**
   * The language of the issue messages: a loaded locale's tag (`'ru-RU'` once `registerLocale(ruRU)` has run) or a
   * locale object (`ruRU`), en-US by default. An issue's code and other fields are the same in every language. A file
   * does not say what language its reader uses, so the host passes the locale the user chose.
   */
  locale?: string | Locale
  /**
   * How large the file's parts may be once inflated: a single part, 512 MiB by default, and all the parts read together,
   * 1 GiB by default. A file that exceeds them is rejected, and so is a part that inflates to more than the file says it
   * holds. A host that reads files uploaded by users should set them lower.
   */
  limits?: XlsxLimits
}
```

### `Refusal`

A check or an edit that failed: a code for programs, a message for users (in the workbook's locale when the call has
one, otherwise in English) and, for typed text, the position in that text. Every check returns either
`{ ok: true, … }` or a [`Refusal`](#refusal) with its own codes.

```ts
interface Refusal<Code extends string = string> {
  ok: false
  code: Code
  message: string
  at?: number
}
```

### `RegexLevel`

A level of match mode 3 (`XLOOKUP`, `XMATCH`) beyond the one built in, which [`registerRegex`](#xlcalc-registerregex) loads: [`pcreSyntax`](#xlcalcregex-pcresyntax)
of `xlcalc/regex`, Excel's syntax translated into JavaScript's regular expressions, or [`pcreEngine`](#xlcalcregex-pcreengine), a PCRE2 engine
of the package's own with Excel's match limit (`#VALUE!`), stopped by a call's `timeoutMs`. Without one, match
mode 3 uses JavaScript's regular expressions in JavaScript's syntax, Unicode `\d \s \w \b`, `\A \z \Z`, `{,m}`
and a leading `(?i)` as PCRE2 takes them.

```ts
interface RegexLevel {
  readonly kind: 'xlcalc-pcre-syntax' | 'xlcalc-pcre-engine'
  /** A pattern's test of a text: true, false or `'limit'` past Excel's match limit; null for a pattern refused (`#VALUE!`). */
  compile(pattern: string): { test(text: string): boolean | 'limit' } | null
}
```

### `ReplaceOptions`

Options of [`Session.replaceInCells`](#session): Find's options as `findCells` accepts them, except what to search, which is always the text a cell opens with in edit mode (Look in: Formulas), and the search order.

```ts
interface ReplaceOptions extends EditOptions, Pick<FindCellsOptions, 'caseSensitive' | 'wildcards' | 'wholeCell' | 'skipHidden' | 'rows'> {}
```

### `ReplaceResult`

What [`Session.replaceInCells`](#session) did: the edit's result and the number of cells whose text it replaced (`replaced`,
counting a legacy array formula as one); the cells whose new text Excel would reject, left unchanged, with the
reasons (`refused`).

```ts
interface ReplaceResult extends EditResult {
  replaced: number
  refused?: (Refusal<CellTextRefusal> & CellRef)[]
}
```

### `RowInput`

Rows of a sheet, like xlsx `<row>`: the height, and the style and number format for their cells that have none of their own.

```ts
interface RowInput {
  /** In points, 0 to 409.5. Without it, a row has Excel's default height, 15 points (one line of Calibri 11). */
  height?: number
  /**
   * The height was set by Excel, by AutoFit or when wrapping text was typed into the row, rather than manually (xlsx
   * writes it without `customHeight`): text typed into the row fits it again (`fitRows`), as for a row of the default
   * height.
   */
  bestFit?: boolean
  style?: number
  numberFormat?: NumberFormat
  /**
   * Hidden (xlsx `hidden`), manually or by a filter: SUBTOTAL with codes 101–111 and AGGREGATE with options 1, 3, 5
   * and 7 ignore the row's cells, and on a `filtered` sheet so does SUBTOTAL with codes 1–11, as in Excel. A row with
   * `height` 0 is also hidden, as Excel treats it. The rows of a collapsed group are hidden rows.
   */
  hidden?: boolean
  /** The outline level (Excel's Group, xlsx `outlineLevel`), 1 to 7; a row without one is not grouped. */
  outlineLevel?: number
  /**
   * The collapsed marker on a group's summary row (xlsx `collapsed`). Excel sets it on every summary row whose group has
   * all its rows hidden, however they were hidden, and clears it when one of them is shown; a summary row inside a group
   * whose rows are all hidden keeps the marker it had, so an inner group stays collapsed when the outer one is expanded
   * again. A workbook may hold other markers (`setLayout` sets the field as given): `outlineGroups` and the outline
   * commands read the markers as Excel would keep them, and `writeXlsx` writes them that way.
   */
  collapsed?: boolean
}
```

### `RowsFit`

Excel's AutoFit of rows `5` or `2:9` (a double-click on the border below a row's number): each row gets the height of
its tallest cell as Excel for Mac fits it ([`fitRows`](#xlcalc-fitrows)), is marked `bestFit` and is unhidden if it was hidden; a row
whose fitted height equals the default gets no height. The heights are calculated from the values and layout left by
the preceding edits (`docs/formats.md`, "Fitting rows").

```ts
interface RowsFit {
  sheet: string
  rows: string
  fit: true
}
```

### `SavedViewport`

A viewport over a saved result ([`viewportOf`](#xlcalc-viewportof)): like [`Viewport`](#viewport), but without a session revision, and with [`SavedViewportCell`](#savedviewportcell) cells.

```ts
interface SavedViewport extends Omit<Viewport, 'revision' | 'cells'> {
  cells: Record<string, SavedViewportCell>
}
```

### `SavedViewportCell`

A cell of a viewport over a saved result ([`viewportOf`](#xlcalc-viewportof)): like [`ViewportCell`](#viewportcell), with its value, `text` and `color` as
the result has them; without `text` and `color` when the result was calculated without display text. The
`appearance` of conditional formats is present only with [`SavedViewportOptions.appearance`](#savedviewportoptions): without it, the cell's
`text` displays the value in the cell's own format, not a rule's.

```ts
type SavedViewportCell = ViewportCell
```

### `SavedViewportOptions`

Options of [`viewportOf`](#xlcalc-viewportof): those of `getViewport`, and `appearance`.

```ts
interface SavedViewportOptions extends ViewportOptions {
  /**
   * Also return, for each cell, the `appearance` its sheet's conditional formats give it, and its `text` and `color` in a
   * number format they set, as `getViewport` returns them in a session of the workbook on the result's clock; false by
   * default. A result holds no appearance: the rules are calculated over the result's values, but its formulas are not
   * recalculated (TODAY in a rule is the result's day). The first viewport with this option over a workbook and a result
   * reads all their values; later ones cost only what the viewport and its rules read.
   */
  appearance?: boolean
}
```

### `Scalar`

A plain value: a number, text, a boolean, or null for an empty cell.

```ts
type Scalar = number | string | boolean | null
```

### `ScanOptions`

Options of [`scanReferences`](#xlcalc-scanreferences).

```ts
interface ScanOptions {
  /** The locale the text is typed in, a loaded locale's tag (`LOCALES`) or a locale object; without it, the text is treated as a stored formula. */
  locale?: string | Locale
}
```

### `SelectionLabels`

The sides of a range whose cells label the others ([`Session.getNamesFromSelection`](#session)), as the boxes of Excel's Create
Names from Selection: the top row, the left column, the bottom row, the right column; at least one.

```ts
interface SelectionLabels {
  top?: boolean
  left?: boolean
  bottom?: boolean
  right?: boolean
}
```

### `Session`

A session over a workbook: edits recalculate only the cells that depend on them. The session's current clock
(`clock`) is what a saved result should record: a full calculation with it produces the same hash.

```ts
interface Session {
  /** Mixed cell, name and formatting operations applied atomically, with one final recalculation. */
  applyChanges(operations: SessionOperation[], opts?: ChangesOptions): ChangesResult
  /** The changes to dependent values and the errors, without changing the session; use the preview's clock when applying. */
  previewChanges(operations: SessionOperation[], opts?: PreviewChangesOptions): ChangesPreview
  /** Reads multiple ranges at a single revision, with only the requested fields and number of cells; range by range, each range's cells row by row. */
  readCells(ranges: RangeRef[], opts?: ReadCellsOptions): ReadCellsResult
  /** Searches where the workbook lives, literally or like Excel's Find; the output is paged, in Find's order. */
  findCells(ranges: RangeRef[], query: string, opts?: FindCellsOptions): ReadCellsResult
  /**
   * Excel's Replace All (`Range.Replace`): in each cell that `findCells` finds by the text it opens with in edit mode
   * (Look in: Formulas), the query's matches in that text are replaced with `replacement` as is, and the edited text is
   * entered back into the cell: a formula in the workbook's language, a number in its format without the `%` of Excel's
   * percent entry, text that would be parsed as something else kept as text, and a legacy array formula as a whole, from
   * its first cell, wherever one of its cells was found. Each `*` matches as little as possible, except the query's last
   * one, which matches the rest. Text Excel would reject (a formula that no longer parses) leaves its cell unchanged
   * (`refused`). No column is widened and no data validation rule is checked. One recalculation, one revision and one
   * undo step; none when nothing is found (`docs/live-books.md`).
   */
  replaceInCells(ranges: RangeRef[], query: string, replacement: string, opts?: ReplaceOptions): ReplaceResult
  /** Direct formula dependents; source ranges stay ranges. */
  getDependents(ranges: RangeRef[], opts?: DependentsOptions): DependentsResult
  /** Totals over nonblank stored values and spills, without returning their cells. */
  summarizeRanges(ranges: RangeRef[], opts?: SummaryOptions): RangeSummary
  readonly clock: Clock
  /** The workbook's revision: 0 when opened, incremented by each call that changes the workbook (`EditResult.revision`). */
  readonly revision: number
  /**
   * Cell edits, applied atomically with one recalculation (`CellEdit`): each sets the fields it gives and keeps the
   * cell's other fields, as Excel keeps a cell's formatting when a value is typed over it; null clears the content, like
   * the Delete key. An edit that keeps the cell's value and formula (a new `style`, `numberFormat`, `meta`, `note` or
   * `link`) recalculates nothing. A value or formula entered in a spill range blocks the array from spilling (#SPILL!);
   * clearing it lets the array spill again. Edits that change some, but not all, cells of a legacy array formula's range
   * (`CellInput.array`) throw `InputError` and change nothing, as Excel does not allow changing part of an array; an edit
   * of its top-left cell that keeps an array formula over the same range changes all of it. An edit of a table's header
   * cell renames its column, as in Excel: the structured references to it follow, a number or a boolean becomes its
   * text, a cleared cell names the column `Column1` (the first free name of this kind, using the word of the workbook's
   * locale: `Столбец1` in ru-RU), and a name already used by another column gets a number appended for the column
   * further right (`Price2`). A formula in a header cell throws `InputError`.
   */
  setCells(edits: CellEdit[], opts?: EditOptions): EditResult
  /** Cells replaced as a whole, otherwise applied as `setCells` applies edits: the cell loses whatever the input omits, and null removes the cell. */
  replaceCells(edits: CellReplace[], opts?: EditOptions): EditResult
  /**
   * Clears the cells of rectangles, like Excel's ClearContents (`contents`: values, formulas and hyperlinks), ClearFormats
   * (`formats`: styles and number formats, and the conditional formats over them) and Clear (`all`: both, plus notes,
   * meta and the data validation rules over them); entire rows and columns also lose their own style and format
   * (`docs/edits.md`).
   */
  clearCells(ranges: RangeRef[], what: ClearKind, opts?: EditOptions): EditResult
  /**
   * Formatting of cells and ranges (`CellPatch`), applied in order and atomically, leaving the cells' values and formulas
   * as they are: an empty cell in a rectangle gets a cell entry with the formatting, a cell left with nothing is removed,
   * and a cell in a legacy array formula's range or in a spill takes the formatting while the array stays. A style given
   * by its fields is applied over each cell's own (`docs/edits.md`); identical styles share one entry of the workbook's
   * `styles`. Recalculates nothing.
   */
  patchCells(patches: CellPatch[], opts?: EditOptions): EditResult
  /**
   * Text typed into a cell, interpreted as Excel interprets it in the workbook's locale, date system and clock, according
   * to the format the cell displays (`parseEntry`): the cell keeps its style, meta, note and link, and a format it
   * inherits from its row or column is not written into it. A number that does not fit a column of the default width
   * widens the column in the same step (`width`), as Excel does, unless `widen` is false. Text with a line break wraps
   * the cell (its style gets `alignment.wrap`, using a matching style in the workbook's table or adding one), and the row
   * of a wrapping cell is fitted to it (`fitRows`) in the same step, unless its height was set manually; a value that
   * the cell's data validation rule does not allow is written and returned with the rule (`validation`), or, with
   * `refuseInvalid`, rejected with the rule and nothing is written. A formula Excel would reject, or a change to part of
   * an array, is rejected with a message in the workbook's locale.
   */
  setCellText(sheet: string, cell: string, text: string, opts?: CellTextOptions): CellTextResult
  /** What `setCellText` would do, without changing anything. */
  previewCellText(sheet: string, cell: string, text: string, opts?: CellTextOptions): CellTextPreview
  /**
   * Texts typed into several cells at once (pasted text, an entry into a selection), each interpreted as `setCellText`
   * interprets it in its cell, and written with one recalculation, one revision and one undo step; a cell given twice
   * takes the later text, interpreted against what the earlier one wrote. Each column widens once, to the largest width
   * that a text typed on its own into one of its cells would give it (`widths`), unless `widen` is false; cells wrap and
   * rows are fitted as `setCellText` does it; the values that the cells' data validation rules do not allow once all
   * the texts are written are returned with their cells (`validations`). A text that `setCellText` would reject rejects
   * the whole call, with its cell, and nothing is written; with `refuseInvalid`, so does the cell whose value violates
   * the strictest rule once all the texts are written, the first one with a rule of that style.
   */
  setCellTexts(entries: CellTextEntry[], opts?: CellTextOptions): CellTextsResult
  /** What `setCellTexts` would do, without changing anything. */
  previewCellTexts(entries: CellTextEntry[], opts?: CellTextOptions): CellTextsPreview
  /**
   * Text pasted from the clipboard at a cell, as Excel pastes it: the rows split at its line breaks and the fields at its
   * tabs (`parseClipboard`; a field that begins with a quote runs to its closing quote) fill a rectangle starting at the
   * cell, and the missing fields of a short row, and a line break after the last row, clear their cells; a field with a
   * line break wraps its cell, as `setCellText` wraps it, but no row is fitted. Each field is interpreted as
   * `setCellText` interprets it, except that an apostrophe stays part of the text, a formula Excel would reject is
   * pasted as text, a percentage cell has no percent entry (`5` becomes 500%) and `€5` gets the euro format; no column
   * is widened and no data validation rule is checked. One recalculation, one revision and one undo step. Rejected, with
   * the cell, where Excel would not paste: part of an array formula, a formula into a table's header cell, and a
   * rectangle beyond the sheet's edge.
   */
  pasteText(sheet: string, cell: string, text: string, opts?: EditOptions): PasteTextResult
  /** What `pasteText` would do, without changing anything. */
  previewPasteText(sheet: string, cell: string, text: string, opts?: EditOptions): PasteTextPreview
  /**
   * Text entered into every cell of a range at once, as Excel's Control+Return enters it from the selection's active
   * cell (`active`, the range's top-left cell by default): interpreted once, as `setCellText` interprets it in the active
   * cell, it gives every cell what the active cell then holds, a formula with its relative references shifted as if
   * copied from the active cell (`copyFormula`), and the active cell's number format, if it has one. The columns widen
   * as typing widens them (`widths`), unless `widen` is false, and text with a line break wraps every cell and fits
   * their rows, as `setCellText` does; a value that the active cell's data validation rule does not allow is written
   * and returned with that rule's alert (`validation`), the only rule checked, or, with `refuseInvalid`, rejected with
   * the active cell. One recalculation, one revision and one undo step. A formula Excel would reject, part of an array
   * formula and a formula in a table's header cell are rejected, with the cell, and nothing is written.
   */
  fillCellText(sheet: string, range: string, text: string, opts?: FillTextOptions): FillTextResult
  /** What `fillCellText` would do, without changing anything. */
  previewFillCellText(sheet: string, range: string, text: string, opts?: FillTextOptions): FillTextPreview
  setNames(names: NameInput[], opts?: EditOptions): EditResult
  /**
   * Replaces the workbook's links to other workbooks (`BookInput.externalBooks`), as a host gives them: a new link with
   * the values it read from the other workbook, more values of a link, values read again. Formulas read a link by its
   * number (`'[2]Sheet1'!A1`), so a new link goes after the others, which keep their places; typed text finds a link by
   * its file name (`[Prices.xlsx]Sheet1!A1`). The workbook is calculated again; one revision and one undo step.
   */
  setExternalBooks(books: ExternalBook[], opts?: EditOptions): EditResult
  /** Non-live sessions only: a live session throws `InputError`. */
  setClock(clock: Clock, opts?: EditOptions): EditResult
  /** A structure operation on the session's workbook, followed by recalculation; rejected wherever `editStructure` rejects it. Hiding or unhiding a sheet recalculates nothing. */
  editStructure(op: StructureOp, opts?: EditOptions): SessionStructureResult
  /** Replaces the workbook's style table; every style that cells, rows and columns use must remain in it. Recalculates nothing; displayed text does not depend on styles. */
  setStyles(styles: CellStyle[], opts?: EditOptions): EditResult
  /**
   * Column widths, row heights, row and column styles and formats, hidden rows and columns, outline levels, merged
   * areas, the filter mode, frozen panes, conditional formats, data validation rules and outline settings, applied
   * atomically. A merged area over a table is rejected with an `InputError`, as in a workbook. Merged areas block arrays
   * from spilling over them: the formulas whose spill range overlaps an added or removed area are recalculated. Hidden rows and the filter mode change what SUBTOTAL and AGGREGATE include: their
   * formulas over the affected rows are recalculated. Nothing else is recalculated. With `display`, the cells whose
   * displayed text or color changed because of a width, a format or a merged area are listed too. Setting or removing a
   * row height without bestFit clears its previous fitted marker; an explicitly given bestFit is honored.
   */
  setLayout(edits: LayoutEdit[], opts?: EditOptions): EditResult
  /**
   * The session's workbook with every change so far: what a host saves as a version. Its objects belong to the
   * session and later calls change them, so do not modify it; serialize or copy whatever must be preserved. The order
   * of its objects' keys (a sheet's fields, a sheet's cells) is not part of the workbook: it can differ between kinds
   * of session and after undo and redo, while the hash, the values and a file written from it do not.
   */
  getBook(): Book
  /**
   * The session's workbook without its cells: the styles, names and settings, and every field of each sheet except
   * `cells` and `looks` (columns, rows, merged areas, tables, frozen panes, conditional formats, data validation and the
   * rest). Its cost depends on the size of the layout, not of the cells; its objects belong to the session, like those
   * of `getBook()`.
   */
  getLayout(): BookLayout
  /**
   * Where a sheet's cells end: the last row and column reached by a cell in its input (a value, a formula, or just a
   * style, a format, meta, a note or a link), a spilled value (the entire spill range), a merged area or a table.
   * Settings of entire rows and columns, conditional formats and data validation do not count. Its cost depends on the
   * sheet's columns, spills and cells without a value; it does not scan all the cells.
   */
  getUsedRange(sheet: string): SheetExtent
  /**
   * The range the sheet's AutoFilter acts on, as Excel reads it (`AutoFilter.Range`): its range (`SheetInput.autoFilter`)
   * grown down over the rows right under it that hold a value or a formula in its columns or the column on either side,
   * one after another, an empty row ending it; null without an AutoFilter. The AutoFilter's own range stays as Excel
   * saves it; a host that applies criteria filters this range (`docs/structure.md` "The AutoFilter"). It reads the rows
   * under the range up to the first empty one.
   */
  getAutoFilterRange(sheet: string): string | null
  /**
   * The address of the cell that Control plus an arrow key moves to from `cell`, as Excel's `Range.End` returns it
   * (`docs/live-books.md`): from a non-empty cell next to another non-empty cell, the last cell of that block; otherwise
   * the next non-empty cell, or the sheet's edge (row 1 or 1,048,576, column A or XFD), where a cell already at the edge
   * stays. A value or a formula makes a cell non-empty, whatever it displays (empty text, zero, an error), and so does
   * being part of a spill or of a legacy array; formatting, a note or a hyperlink alone does not. Hidden rows and columns
   * are treated as absent: whether hidden manually, by a filter or by a size of 0, `End` never stops on them and a block
   * continues across them, and when every row or column beyond a cell is hidden, the edge is the last visible one. A
   * merged area counts as its individual cells, with its value in the top-left one. Its cost depends on the cells along
   * the row or column it moves along, up to where that line's values end; it does not scan the sheet.
   */
  getEnd(sheet: string, cell: string, direction: EndDirection): string
  /**
   * What Excel's AutoSum (Home > Editing, Alt+=) offers at a cell: the range of numbers it sums, above the cell, else to
   * its left, and the `SUM` formula it enters, in the workbook's locale too (`AutoSumGuess`). Without numbers the formula
   * is `SUM()`, which Excel opens for editing and refuses to enter. It reads the cells along the cell's column and row.
   */
  getAutoSum(sheet: string, cell: string): AutoSumGuess
  /**
   * The names Excel's Create from Selection (Formulas > Defined Names) makes over a range (`SelectionLabels`): each
   * label of the chosen sides, its text made a name as Excel makes it, refers to the cells of its column or row in the
   * range past the labels; with a row and a column of labels, the label in their corner refers to all of them. A label
   * without text makes none, and of two labels with one name the later one stays. A host passes the names to
   * `setNames`, asking first, as Excel does, about a name the workbook already has.
   */
  getNamesFromSelection(sheet: string, range: string, labels: SelectionLabels): NameInput[]
  /**
   * The cells of a kind (`SpecialCells`) Excel's Go To Special selects in a range, as rectangles (`B2:C9`, `E4`); none
   * where Excel says no cells were found. A range of one cell stands for the sheet's used range, as in Excel, and a
   * larger one is taken within it.
   */
  getSpecialCells(sheet: string, range: string, kind: SpecialCells): string[]
  /** What Excel's Workbook Statistics shows with `sheet` active: the sheet's end and counts, and the workbook's counts. */
  getWorkbookStatistics(sheet: string): WorkbookStatistics
  explain(sheet: string, cell: string, opts?: ExplainOptions): Explanation
  /**
   * The cells of a rectangle on a sheet (`B2:D9`, `B2`, entire columns `B:D`, entire rows `2:9`), or of several (a cell
   * in two of them is listed once), as `getResult()` returns them, by address: the same cells (including spilled ones,
   * without empty ones) with the same fields, plus each cell's own `style`, `note` and `link`, a legacy array formula's
   * `array`, and the `appearance` the sheet's conditional formats give it based on the current values (an empty cell
   * with an appearance is listed with the value null). For a viewport over a sheet: its cost depends on what the
   * rectangles hold, not on the workbook; the statistics of a rule's cells (lowest, highest, percentiles) are cached
   * until an edit changes them. Edits do not list cells whose conditional appearance changed: they return the rules'
   * cells as `repaint`, for a host to read again where they overlap its viewport.
   */
  getCells(sheet: string, range: string | string[]): Record<string, CellView>
  /**
   * Plain-text clipboard data for a range: calculated values in their effective formats and the workbook's locale,
   * not clipped to the column width, skipping hidden rows and columns. Explicit rectangles keep selected blanks;
   * entire rows or columns stop at the used range, and an empty sheet produces one empty field. Changes nothing.
   */
  copyText(sheet: string, range: string, opts?: CopyTextOptions): string
  /**
   * The values of a rectangle on a sheet (`B2:D9`, `B:B`, `2:9`), as rows from top to bottom, each from left to right:
   * every cell's value as `readCells` returns it in the field `value` (a number to 15 significant digits, text, a
   * boolean, an error as `{ error }`), or null where a cell holds none. The rectangle is exactly the requested one,
   * including its empty cells, not trimmed to where the sheet's cells end; one with more than `maxCells` cells is
   * rejected (`CalcError` `limit`). Each call returns new arrays. Its cost is the rectangle's cells, without a result
   * object and an address for each; changes nothing.
   */
  getValues(sheet: string, range: string, opts?: ValuesOptions): CellValue[][]
  /** A single cell as `getCells(sheet, cell)` returns it; undefined for a cell that `getCells` leaves out. */
  getCell(sheet: string, cell: string): CellView | undefined
  /**
   * A viewport over a sheet for a grid (`B2:D9`, `B:D`, `2:9`): its cells with their styles and number formats resolved,
   * including empty cells with a style or number format from their row or column, the widths and heights of its columns
   * and rows, the merged areas and tables that overlap it, and the sheet's frozen panes (`Viewport`); only some of its
   * rows and columns with `opts.rows` and `opts.columns`. Like `getCells`, its cost depends on what the viewport holds,
   * not on the workbook.
   */
  getViewport(sheet: string, range: string, opts?: ViewportOptions): Viewport
  /**
   * Why a cell holds an error, or why a circular reference stops it, and the cell where it originated
   * (`CellIssueDetail`), without building the workbook's result; null for a cell with neither.
   */
  getCellIssue(sheet: string, cell: string): CellIssueDetail | null
  /** A cell's data validation (`CellValidation`), with its list calculated now; null when no rule covers the cell. */
  getValidation(sheet: string, cell: string): CellValidation | null
  /**
   * Whether the cell would satisfy its rule if it held `input` (null: cleared), as Excel checks a cell's value
   * (`Range.Validation.Value`): the value the input calculates to, in the workbook with that input. Changes nothing; a
   * cell without a rule accepts anything. Text typed into the cell must first be converted to input by `parseEntry`.
   */
  validateEntry(sheet: string, cell: string, input: CellInput | null): EntryCheck
  getResult(): ComputeResult
  /** An independent copy of the workbook and its result at a single point in time; later edits and changes to the copy cannot affect each other. */
  getSnapshot(): BookSnapshot
  /** The values hash that `getResult()` returns, without building the result. */
  getHash(): string
  /**
   * Reverts the last change to the workbook, like Excel's Undo: everything one call changed (`UndoCall`), then
   * recalculates; null when there is nothing to undo. A change after an undo clears what `redo` could have restored.
   */
  undo(opts?: EditOptions): UndoResult | null
  /** Reapplies the last change that `undo` reverted; null when there is none. */
  redo(opts?: EditOptions): UndoResult | null
  /** The number of steps `undo` and `redo` can take now. */
  readonly history: SessionHistory
  /**
   * Calls `listener` after each call that changes the workbook, whoever makes it: an edit, a structure operation, `undo`,
   * `redo` (`SessionChange`: the call's `EditResult` and its name). Not after a preview, a read, a rejected call or a
   * call that changed nothing. A listener runs when the call has finished, before it returns, so the session it reads is
   * the one after the change; listeners run in the order they subscribed, and changes reach them in revision order: a
   * change made by a listener reaches every listener after the change it responds to. A listener that throws stops
   * neither the call nor the other listeners; its error is rethrown in a microtask. Returns the function that
   * unsubscribes; `close()` unsubscribes every listener.
   */
  subscribe(listener: (change: SessionChange) => void): () => void
  close(): void
}
```

### `SessionChange`

A change to a session's workbook as its subscribers receive it ([`Session.subscribe`](#session)): the [`EditResult`](#editresult) of the call
that made it (the revision after the call, the cells it changed, the sheets whose layout changed, the style table,
how addresses moved, where to repaint) and the call's name.

```ts
interface SessionChange extends EditResult {
  method: ChangingMethod
}
```

### `SessionHistory`

What a session can undo and redo now: the number of steps kept in each direction.

```ts
interface SessionHistory {
  undo: number
  redo: number
}
```

### `SessionOperation`

Serializable operations applied in order, with one final recalculation and one undo step.

```ts
type SessionOperation =
  | { op: 'setCells'; edits: CellEdit[] }
  | { op: 'replaceCells'; edits: CellReplace[] }
  | { op: 'clearCells'; ranges: RangeRef[]; what: ClearKind }
  | { op: 'patchCells'; patches: CellPatch[] }
  | { op: 'setNames'; names: NameInput[] }
  | { op: 'setStyles'; styles: CellStyle[] }
  | { op: 'setLayout'; edits: LayoutEdit[] }
```

### `SessionOptions`

Session options. A session keeps a single clock, `clock` or by default the time it was opened and a random
seed, until `setClock`. A live session (`live: true`) instead takes a fresh clock and a new seed on every
recalculation. Edits are compiled with the `execution` the session was opened with (`closures` by
default). With `display` (the default), results and changes include what the cells display. `limits` apply to every
call; a call that exceeds them throws [`CalcError`](#xlcalc-calcerror) and leaves the session unchanged. `undo` limits the steps
[`Session.undo`](#session) keeps. `cached`, as [`compute`](#xlcalc-compute) accepts it, gives the formulas stopped by a circular reference the values
stored in the workbook's file; from then on, the session's workbook holds them as `held`.

```ts
type SessionOptions = ({ clock?: Clock; live?: false } | { live: true }) & { externals?: Externals; execution?: Execution; display?: boolean; limits?: CalcLimits; undo?: UndoOptions; cached?: CachedValues }
```

### `SessionStructureResult`

[`Session.editStructure`](#session): the edit's result and where the cells went, without the workbook (`getBook()` returns it),
in every kind of session. `changes` are the cells whose value or formula differs from that of the cell previously at
their old position (`mapCell`), by new address, with a cell that no old cell moved to counting as empty before; with
`display`, also the cells whose displayed text or color differs. An interface that moves its displayed cells
according to `moved` (or `mapCell`) and applies these changes has the new cells.

```ts
type SessionStructureResult =
  | (EditResult & {
      ok: true
      /** Where a cell of the workbook before the operation is now; null if it was deleted or overwritten, or if its sheet is not in that workbook. */
      mapCell(ref: CellRef): CellRef | null
    })
  | Refusal<StructureRefusal>
```

### `Sheet`

A sheet of a [`Book`](#book).

```ts
type Sheet = SheetInput
```

### `SheetExtent`

Where a sheet's cells end ([`Session.getUsedRange`](#session)): all content lies within its first `rows` rows and `columns` columns; 0 and 0 for an empty sheet.

```ts
interface SheetExtent {
  rows: number
  columns: number
}
```

### `SheetInput`

A sheet of a workbook ([`BookInput.sheets`](#bookinput)): its name, its cells by address, and everything else a sheet holds (rows, columns, merged areas, tables, formatting, rules, frozen panes).

```ts
interface SheetInput {
  name: string
  /** Cells by A1 address (`B12`). */
  cells: Record<string, CellInput>
  /**
   * The formatting of the cells that have no entry in `cells`, as rectangles, each applied over those before it: a sheet
   * formatted across many empty cells (entire rows out to XFD) is stored without an object per cell. A cell without an
   * entry uses the style and the number format of the last rectangle that covers it, both from that rectangle (a
   * rectangle without a field gives the cell none of its own); a cell with an entry uses its own fields, whatever the
   * rectangles say. A field the cell lacks comes from its row, or else its column, as for an entry.
   * `[{ range: 'A20:XFD22', style: 0 }, { range: 'R20:R22', style: 1, numberFormat: '0.0%' }]` is the same sheet as an
   * entry `{ style: 0 }` in each cell of A20:XFD22 except R20:R22, which hold `{ style: 1, numberFormat: '0.0%' }`. On
   * input, every entry that has only a style or a number format is moved here, and the calls that return a workbook
   * return this list in a form of their own (`LookRange`).
   */
  looks?: LookRange[]
  /** Columns by letter, a single column (`C`) or a span (`B:D`) like xlsx `<col min max>`; spans do not overlap. */
  columns?: Record<string, ColumnInput>
  /** Rows by number, a single row (`5`) or a span (`5:9`); spans do not overlap. */
  rows?: Record<string, RowInput>
  /** Merged areas (`B2:D3`), like xlsx `<mergeCell>`: two or more cells each, none overlapping another. */
  merges?: string[]
  /**
   * A filter on the sheet has criteria, the sheet's AutoFilter or a table's (Excel's filter mode): every hidden row of
   * the sheet counts as filtered out, so SUBTOTAL ignores it with codes 1–11 too and a copy skips it, as Excel does
   * whatever hid the row.
   */
  filtered?: boolean
  /**
   * The sheet's AutoFilter (Excel's Filter over cells outside a table, xlsx `<autoFilter>`): the filter buttons on the
   * first row of its range. It does not overlap a table of the sheet, which has a filter of its own. Its criteria and
   * sort are not in the model: whether a filter hid rows is `filtered`. Not part of calculation or inputHash.
   */
  autoFilter?: AutoFilter
  /** The sheet's tables (Excel's list objects); they do not overlap, and no merged area or multi-cell array range overlaps one. */
  tables?: TableInput[]
  /**
   * Frozen panes (Excel's Freeze Panes): the first `rows` rows and the first `columns` columns stay visible while the
   * sheet scrolls; at least one of the two is set (`{ rows: 1 }` freezes the top row). Not part of calculation.
   */
  frozen?: FrozenPanes
  /** The saved sheet view from xlsx; it does not affect calculation or inputHash. */
  view?: SheetView
  /** The color of the sheet's tab (Excel's Tab Color, xlsx `<tabColor>`), `#RRGGBB`. Not part of calculation or inputHash. */
  tabColor?: Color
  /**
   * The sheet's conditional formats (Excel's conditional formatting rules) in priority order, highest first. They do not
   * change values: a session returns their effect on display in `CellView.appearance` (`docs/conditional-formats.md`).
   */
  conditionalFormats?: ConditionalFormat[]
  /**
   * Data validation (Excel's Data Validation): the rules for the sheet's cells, at most one rule per cell. Not part of
   * calculation: a session returns a cell's rule and its list (`Session.getValidation`) and checks a value against it
   * (`Session.validateEntry`).
   */
  validations?: DataValidation[]
  /** Where the summary rows and columns of the sheet's groups are placed (`RowInput.outlineLevel`); below and to the right by default. */
  outline?: OutlineSettings
  /**
   * The sheet is hidden (Excel's Hide, xlsx `state="hidden"`) or very hidden (`'very'`, xlsx `state="veryHidden"`, a
   * sheet that Excel's interface does not offer to unhide; only its object model can). Not part of calculation: formulas
   * read a hidden sheet like any other. A workbook always has at least one visible sheet.
   */
  hidden?: boolean | 'very'
  /**
   * The sheet is protected (Excel's Protect Sheet, xlsx `<sheetProtection>`), with what users may still do: its cells
   * that are not locked take edits (`CellStyle.protection`). Not part of calculation or `inputHash`, and a session does
   * not enforce it: the host does. The workbook holds no password.
   */
  protection?: SheetProtection
}
```

### `SheetLayout`

A sheet without its cells and their formatting ([`Session.getLayout`](#session)).

```ts
type SheetLayout = Omit<Sheet, 'cells' | 'looks'>
```

### `SheetNameCheck`

The result of [`validateSheetName`](#xlcalc-validatesheetname); `at` is the position of the character Excel rejects.

```ts
type SheetNameCheck = { ok: true } | Refusal<SheetNameCode>
```

### `SheetNameCode`

Why Excel rejects a sheet name: empty, longer than 31 characters, containing one of `: \ / ? * [ ]`, an apostrophe at either end, or `History`.

```ts
type SheetNameCode = 'empty' | 'too_long' | 'character' | 'apostrophe' | 'reserved'
```

### `SheetOutline`

A sheet's groups ([`outlineGroups`](#xlcalc-outlinegroups)), on each axis by first line, with the outer group before the inner one when two start together.

```ts
interface SheetOutline {
  rows: OutlineGroup[]
  columns: OutlineGroup[]
}
```

### `SheetProtection`

What users may do on a protected sheet ([`SheetInput.protection`](#sheetinput)), as the boxes of Excel's Protect Sheet dialog: each
field true allows it. Selecting locked and unlocked cells is allowed by default, the rest is not, and a protection
holds only what differs from that: `{}` is Excel's default protection, `{ formatCells: true }` also allows
formatting cells, `{ selectLockedCells: false }` allows selecting only the unlocked ones.

```ts
interface SheetProtection {
  selectLockedCells?: boolean
  selectUnlockedCells?: boolean
  formatCells?: boolean
  formatColumns?: boolean
  formatRows?: boolean
  insertColumns?: boolean
  insertRows?: boolean
  insertHyperlinks?: boolean
  deleteColumns?: boolean
  deleteRows?: boolean
  sort?: boolean
  /** Use AutoFilter. */
  autoFilter?: boolean
  /** Use PivotTable and PivotChart. */
  pivotTables?: boolean
  /** Edit objects: drawings, charts, notes. */
  objects?: boolean
  /** Edit scenarios. */
  scenarios?: boolean
}
```

### `SheetView`

The first saved sheet view; selection holds A1 cells or ranges, without a sheet name.

```ts
interface SheetView {
  activeCell?: string
  selection?: string[]
  topLeft?: string
  /** Zoom in percent, 10 to 400. */
  zoom?: number
  /** False hides the sheet's gridlines (View > Show > Gridlines, xlsx `showGridLines="0"`); shown by default. */
  gridlines?: boolean
  /** False hides the row and column headings (View > Show > Headings, xlsx `showRowColHeaders="0"`); shown by default. */
  headings?: boolean
  /** The window split into panes without freezing them (View > Split); frozen panes (`SheetInput.frozen`) take its place while the sheet has them. */
  split?: SplitPanes
}
```

### `SpecialCells`

The cells Excel's Go To Special (Home > Editing > Find & Select) selects ([`Session.getSpecialCells`](#session)): blank cells,
constants or formulas (of the kinds in `values`, every kind without it; a formula by the value it answers), or the
visible cells, those of no hidden row or column.

```ts
type SpecialCells = { type: 'blanks' } | { type: 'constants' | 'formulas'; values?: SpecialValue[] } | { type: 'visible' }
```

### `SpecialValue`

What a cell holds by the kinds of Excel's Go To Special: a number, text, a boolean, an error ([`SpecialCells.values`](#specialcells)).

```ts
type SpecialValue = 'numbers' | 'text' | 'logicals' | 'errors'
```

### `SplitPanes`

A window split into panes ([`SheetView.split`](#sheetview), xlsx `<pane>` without `state="frozen"`): a vertical line `x` points from
the window's left edge, a horizontal one `y` points from its top, at least one of them. As Excel saves them, the
points count the row and column headings in and are points on the screen at the view's zoom (xlsx `xSplit` and
`ySplit` in twentieths of a point): a grid draws a line at `x` less the width of its own row headings.

```ts
interface SplitPanes {
  x?: number
  y?: number
  /**
   * The top left cell of the pane after both lines (right of the vertical one, below the horizontal one), which also
   * gives the scroll of the other panes on its axes; the first pane's is the view's `topLeft`. Without it Excel scrolls
   * the panes after the lines to A1.
   */
  topLeft?: string
  /** The pane that holds the active cell and the selection; the first without it. */
  activePane?: PaneName
}
```

### `StatisticsCounts`

What Excel's Workbook Statistics (Review > Workbook Statistics) counts, on a sheet or in the workbook. The workbook
holds no PivotTables, charts, pictures, form controls, other objects or threads of comments (a thread read from a
file is a note): those counts are null.

```ts
interface StatisticsCounts {
  /** Cells holding a value or a formula; each cell of a spilled array or of a legacy array formula is one. */
  cells: number
  tables: number
  /** Cells with a formula; each cell of a spilled array or of a legacy array formula is one, as Excel counts them. */
  formulas: number
  notes: number
  pivotTables: null
  charts: null
  images: null
  formControls: null
  objects: null
  comments: null
}
```

### `StructureOp`

One structure operation ([`editStructure`](#xlcalc-editstructure)). Rows are 1-based numbers and columns are letters, as in A1
addresses; `count` is 1 by default; sheets and tables are matched case-insensitively. Tables move with their cells
and grow and shrink with the rows and columns inserted and deleted inside them; an inserted row gets the formulas of
the calculated columns, an inserted column a name (`Column1`), and references to a deleted column become #REF!.

```ts
type StructureOp =
  /** Inserts `count` rows before `row`: the rows from `row` onward move down; the new rows take the format of the row above. */
  | { op: 'insertRows'; sheet: string; row: number; count?: number }
  /** Deletes rows `row` … `row + count - 1`: the rows below move up. */
  | { op: 'deleteRows'; sheet: string; row: number; count?: number }
  /** Inserts `count` columns before `column`: the columns from `column` onward move right; the new ones take the format of the column to the left. */
  | { op: 'insertColumns'; sheet: string; column: string; count?: number }
  /** Deletes `count` columns starting at `column`: the columns to the right move left. */
  | { op: 'deleteColumns'; sheet: string; column: string; count?: number }
  /** Cut and paste: the range moves so that its top-left cell lands on `to`, replacing what was there; cells keep their `meta`. */
  | { op: 'moveRange'; from: RangeRef; to: CellRef }
  /**
   * Copy and paste: the range is pasted at `to`; a target that is a whole multiple of the source's size is
   * tiled with copies (fill down and fill right are copies of the first row or column). Copies have no
   * `meta` unless `withMeta`. `paste` pastes part of the copy, as Excel's Paste Special (`PasteKind`), `transpose`
   * turns it, rows into columns, `operation` combines the copy's numbers with the destination's, and `skipBlanks`
   * leaves the contents and look of the destination's cells under blank cells of the copy, and their notes under cells
   * without a note (`docs/structure.md` "Paste Special"); `operation` or `skipBlanks` without `paste` is Paste
   * Special's `all`.
   */
  | { op: 'copyRange'; from: RangeRef; to: RangeRef; withMeta?: boolean; paste?: PasteKind; transpose?: boolean; operation?: PasteOperation; skipBlanks?: boolean }
  /** A new name that follows Excel's rules (see `BookInput.sheets`) and is not used by another sheet. */
  | { op: 'renameSheet'; sheet: string; name: string }
  | { op: 'deleteSheet'; sheet: string }
  /**
   * Inserts a new empty sheet named `name` (following Excel's rules, not already used), before the sheet `before`, or
   * else after the last one. A reference that spans sheets includes a sheet inserted between its ends.
   */
  | { op: 'insertSheet'; name: string; before?: string }
  /**
   * Moves the sheet before the sheet `before`, or else after the last one; when `before` is hidden, the sheet goes before
   * the next visible one, as Excel moves it. A reference that spans sheets covers the sheets between its ends in the new
   * order; an end sheet that moves remains an end until it passes the other end, and then the next sheet inward takes
   * its place.
   */
  | { op: 'moveSheet'; sheet: string; before?: string }
  /** Hides the sheet (Excel's Hide Sheet), or makes it very hidden with `very` (see `SheetInput.hidden`); a workbook always keeps one visible sheet. */
  | { op: 'hideSheet'; sheet: string; very?: boolean }
  /** Unhides the sheet (Excel's Unhide Sheet). */
  | { op: 'unhideSheet'; sheet: string }
  /** Creates a table over cells of `sheet` (Excel's Format as Table): with a header row, its header cells take the column names. */
  | { op: 'addTable'; sheet: string; table: TableInput }
  /** A new name for the table, following the rules of `TableInput.name`; every structured reference to it is updated. */
  | { op: 'renameTable'; table: string; name: string }
  /** Deletes the table and its cells (Excel's Delete Table): structured references to it become #REF!. */
  | { op: 'deleteTable'; table: string }
  /**
   * Converts the table to plain cells (Excel's Convert to Range): structured references to it become references to its
   * cells with `$` and the sheet name (`'S'!$B$2:$B$9`; a reference to the formula's own row keeps the row relative,
   * `'S'!$B2`).
   */
  | { op: 'convertTable'; table: string }
  /**
   * Excel's Group and Ungroup of rows `5:7` or columns `B:D` (a single line: `5`, `C`). Group moves each line one
   * outline level deeper, leaving those already at the deepest level (7) there, and is rejected (`outline_deep`) when
   * all of them are at that level; adjacent lines at a level form one group, so a group added right after another of
   * its level merges with it (`outlineGroups`). Ungroup removes one level from each line that has one, and is rejected
   * (`not_grouped`) when none has. Hidden lines stay hidden.
   */
  | { op: 'group' | 'ungroup'; sheet: string; rows: string }
  | { op: 'group' | 'ungroup'; sheet: string; columns: string }
  /**
   * Excel's Show Detail and Hide Detail, a group's `+` and `−` buttons: at a summary line (`OutlineGroup.summary`), on
   * the lines it summarizes; at a line inside a group, on the group at that line's level that contains it, by way of the
   * group's summary line. Hiding hides the lines and marks the summary line `collapsed`; showing unhides them, except the
   * lines of the collapsed groups inside, and removes the marker. Rejected at a line in no group, or in a group whose
   * summary line is beyond the sheet's edge (`not_grouped`), and where the detail is already shown, or already hidden
   * (`detail`).
   */
  | { op: 'showDetail' | 'hideDetail'; sheet: string; row: number }
  | { op: 'showDetail' | 'hideDetail'; sheet: string; column: string }
  /**
   * Excel's outline level buttons (`Outline.ShowLevels`): hides the lines at outline level `rows` (or `columns`) and
   * deeper, shows the other grouped lines, and sets the collapsed markers as Excel keeps them; 8 shows every group. Lines
   * outside groups stay as they are.
   */
  | { op: 'showLevels'; sheet: string; rows?: number; columns?: number }
  /**
   * Excel's Merge Cells over `range` (`B2:D3`, including entire columns and rows): the range, extended by the merged
   * areas it overlaps, becomes one merged area (`SheetInput.merges`) that keeps its top-left cell; when that cell is
   * empty, the area's first cell holding a value or a formula, row by row, moves into it as cut and paste would move it.
   * The other cells lose their values and formulas and take the top-left cell's style and number format. With `across`,
   * each row of the range is merged separately (Merge Across); with `center`, the area is also centered (Merge &
   * Center). A range that overlaps a table is rejected (`table`), as is one that overlaps a legacy array formula
   * (`array`).
   */
  | { op: 'mergeCells'; sheet: string; range: string; across?: boolean; center?: boolean }
  /** Excel's Unmerge Cells: the merged areas that overlap `range` are split back into their cells. */
  | { op: 'unmergeCells'; sheet: string; range: string }
  /**
   * Excel's fill handle (`Range.AutoFill`): the cells of `range` are extended over `to`, which contains `range` and
   * extends beyond it along one axis, down, up, right or left (`A1:B2` over `A1:B9`, or over `A1:F2`). The added cells
   * take the source's cells in turn, as copies do (formulas, styles, number formats), with the part of the source
   * nearest to it at the far end when `to` is not a whole number of source lengths; the rejections are those of a copy.
   * Each column of a fill down or up, and each row of a fill right or left, is extended independently, and whatever
   * continues as a series takes its next values instead (`FillType`, docs/structure.md "The fill handle"). A `to` of any
   * other shape throws an `InputError`.
   */
  | { op: 'fillRange'; sheet: string; range: string; to: string; type?: FillType }
```

### `StructureRefusal`

Why Excel would reject an operation.

```ts
type StructureRefusal =
  /** Cells would go off the sheet (beyond XFD or row 1,048,576), or a paste does not fit. */
  | 'off_sheet'
  /** A new sheet name that breaks Excel's rules or is already used (sheet names are compared case-insensitively). */
  | 'sheet_name'
  /** A workbook keeps at least one visible sheet: its last visible sheet cannot be deleted or hidden. */
  | 'last_sheet'
  /** A very hidden sheet (`SheetInput.hidden` 'very') cannot be deleted or moved, as in Excel: unhide it first. */
  | 'very_hidden'
  /** A move or copy would split a merged area, at its source or at its destination. */
  | 'merged'
  /** A copy that skips rows hidden by a filter, pasted over hidden rows of a filtered sheet (multiple areas onto multiple areas). */
  | 'filtered'
  /** The operation would change part of a legacy array formula's range (`CellInput.array`), or Paste Special's operation a spilled array. */
  | 'array'
  /**
   * The operation would shift or split part of a table in a way Excel rejects: rows or columns inserted into or deleted
   * from two tables at once, a table's header row deleted without the rest of the table, a move that splits a table, a
   * table on top of another.
   */
  | 'table'
  /** A table name that breaks the rules of `TableInput.name` or is already used by a table or a defined name in the workbook. */
  | 'table_name'
  /** Group on lines that are all at the deepest outline level, 7. */
  | 'outline_deep'
  /** Ungroup on lines none of which has an outline level; Show or Hide Detail at a line in no group, or in a group without a summary line. */
  | 'not_grouped'
  /** Show Detail on a group already shown, Hide Detail on one already hidden (its summary line marked `collapsed`). */
  | 'detail'
```

### `StructureResult`

The result of [`editStructure`](#xlcalc-editstructure).

```ts
type StructureResult =
  | {
      ok: true
      /** The new workbook; the input workbook is not modified, and unchanged sheets and cells are shared. */
      book: Book
      /** Where a cell of the old workbook is now; null if it was deleted or overwritten, or if its sheet is not in the old workbook. */
      mapCell(ref: CellRef): CellRef | null
    }
  /** A rejection: its code, and a message in the language of the workbook's locale (English for en-US). */
  | Refusal<StructureRefusal>
```

### `StylePatch`

Fields of a style to change ([`CellPatch.set.style`](#cellpatch)): a value sets a field, null removes it, and an omitted field is left unchanged.

```ts
interface StylePatch {
  font?: { [K in keyof FontStyle]?: FontStyle[K] | null } | null
  fill?: { color: Color } | null
  border?: { [K in keyof NonNullable<CellStyle['border']>]?: BorderLine | null } | null
  alignment?: { [K in keyof AlignmentStyle]?: AlignmentStyle[K] | null } | null
  protection?: { [K in keyof CellProtection]?: CellProtection[K] | null } | null
}
```

### `SummaryOptions`

Options of `summarizeRanges`: only the cells in these rows (the rows a host's filter displays), as in [`ReadCellsOptions.rows`](#readcellsoptions).

```ts
interface SummaryOptions extends EditOptions {
  rows?: number[]
}
```

### `TableColumn`

A column of a table ([`TableInput.columns`](#tableinput)): its header name and what its cells and its totals row calculate.

```ts
interface TableColumn {
  /** 1 to 255 characters, unique within the table, compared case-insensitively; the header cell's text. */
  name: string
  /**
   * The formula of a calculated column (as stored, `Sales[[#This Row],[Qty]]*2`; a leading `=` is removed), which a row
   * inserted into the table receives in this column, as Excel fills it. The column's cells hold their own formulas.
   */
  formula?: string
  /**
   * What the column's cell in the totals row contains, as an xlsx file records it: one of Excel's totals functions, a
   * SUBTOTAL over the column that the cell holds as its formula (`custom`: a formula of its own), or a label. The cell
   * calculates whatever it holds.
   */
  totals?: { function: TotalsFunction } | { label: string }
}
```

### `TableInput`

A table (Excel's list object): a rectangle on a sheet with a header row of column names, data rows and a totals row.
Formulas refer to its parts by name with structured references (`Sales[Amount]`, `SUM(Sales[[#Totals],[Qty]])`),
which follow the table as rows and columns are inserted into it and deleted from it. A formula stores a reference to
its own row the way an xlsx file does, `Sales[[#This Row],[Amount]]`; `Sales[@Amount]`, and `[@Amount]` inside the
table, are also accepted.

```ts
interface TableInput {
  /**
   * Excel's rules for names: a letter of any script, `_` or `\` first, then letters, digits, `_`, `.`, `\` and `?`, at
   * most 255 characters; not a cell address (`T2`) or an R1C1 reference (`R`, `C`, `R1C1`), not TRUE or FALSE, nor the
   * workbook locale's words for them, in any case (ИСТИНА and ЛОЖЬ in ru-RU, which Excel for Windows ru-RU rejects).
   * Tables and defined names share one namespace: no other table and no defined name in the workbook may have the same
   * name, compared case-insensitively.
   */
  name: string
  /** The table's rectangle, including its header and totals rows (`B2:E10`), with at least one data row. */
  range: string
  /** The first row is a header row (true by default): its cells hold the column names as text, not formulas. */
  headers?: boolean
  /** The last row is a totals row (false by default). */
  totals?: boolean
  /** The columns from left to right, one for each column of the range. */
  columns: TableColumn[]
  /**
   * How Excel draws the table (xlsx `tableStyleInfo`); without it, Excel's default, `TableStyleMedium2` with banded rows.
   * The engine does not draw it: it stays with the table through operations and is written to the file.
   */
  style?: TableStyle
}
```

### `TableReference`

A structured reference (`Sales[Qty]`, `[@Qty]`): the table name as written (absent inside a table, `[@Qty]`), the
column or the first and last columns of a span, and the item specifiers in their stored form (`#Headers`, `#This Row`
for `@`); a host maps it onto the table's range.

```ts
interface TableReference extends ReferenceSpan {
  kind: 'table'
  table?: string
  columns?: [string] | [string, string]
  items?: ('#All' | '#Data' | '#Headers' | '#Totals' | '#This Row')[]
}
```

### `TableStyle`

A table's style ([`TableInput.style`](#tableinput)): one of Excel's table styles, and the parts of the table it emphasizes, each
false by default, as the checkboxes of Excel's Table Style Options.

```ts
interface TableStyle {
  /** One of Excel's table styles, `TableStyleLight1` to `TableStyleLight21`, `TableStyleMedium1` to `TableStyleMedium28`, `TableStyleDark1` to `TableStyleDark11`; none without it. */
  name?: string
  firstColumn?: boolean
  lastColumn?: boolean
  /** Banded rows. */
  rowStripes?: boolean
  /** Banded columns. */
  columnStripes?: boolean
}
```

### `TextFormat`

The cell's text contains `text`, does not contain it, begins with it or ends with it, in any case. Excel calculates
the rule's formula, which it writes alongside the text (`NOT(ISERROR(SEARCH("ab",A1)))`); without `formula`, the
rule uses that formula.

```ts
interface TextFormat extends ConditionalBase {
  type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith'
  text: string
  formula?: string
  style?: ConditionalStyle
}
```

### `TimePeriodFormat`

The cell contains a date in the period, according to the workbook's clock (`TODAY()`); calculated by `formula`, or by Excel's formula without it.

```ts
interface TimePeriodFormat extends ConditionalBase {
  type: 'timePeriod'
  period: 'today' | 'yesterday' | 'tomorrow' | 'last7Days' | 'thisWeek' | 'lastWeek' | 'nextWeek' | 'thisMonth' | 'lastMonth' | 'nextMonth'
  formula?: string
  style?: ConditionalStyle
}
```

### `TopFormat`

The `rank` highest numbers among the cells (the lowest with `bottom`), or `rank` percent of them with `percent`.

```ts
interface TopFormat extends ConditionalBase {
  type: 'top10'
  /** 1 to 1,000; 1 to 100 with `percent`. */
  rank: number
  percent?: boolean
  bottom?: boolean
  style?: ConditionalStyle
}
```

### `TotalsFunction`

Excel's totals functions of a table column (xlsx `totalsRowFunction`).

```ts
type TotalsFunction = 'sum' | 'average' | 'count' | 'countNums' | 'max' | 'min' | 'stdDev' | 'var' | 'custom'
```

### `TranslateOptions`

[`toLocal`](#xlcalc-tolocal)'s and [`toInvariant`](#xlcalc-toinvariant)'s options: `table`, the table of the formula's cell; `externalBooks`, the book's links to
other workbooks, by which a reference to one is shown as Excel shows it (`'[1]Sheet1'!A1` is `[Book2.xlsx]Sheet1!A1`,
`'../data/[Book3.xlsx]Sheet1'!A1` with the path the link holds) and read back by its file name.

```ts
interface TranslateOptions {
  table?: string
  externalBooks?: ExternalBook[]
}
```

### `Translation`

The result of translating formula text: the text, or a check code with a message in the locale's language and the
position in the given text: `syntax`, or `external_ref` for a reference to a workbook not among the given links.

```ts
type Translation = { ok: true; formula: string } | (Refusal<'syntax' | 'external_ref'> & { at: number })
```

### `UndoCall`

The calls a session can undo: those that change its workbook. `setClock` and a live session's new clock are not steps.

```ts
type UndoCall = 'applyChanges' | 'setCells' | 'replaceCells' | 'clearCells' | 'patchCells' | 'setCellText' | 'setCellTexts' | 'pasteText' | 'fillCellText' | 'replaceInCells' | 'setNames' | 'setExternalBooks' | 'setStyles' | 'setLayout' | 'editStructure'
```

### `UndoOptions`

The steps a session keeps for `undo`: up to `depth` of them, 100 by default as in Excel (0 keeps none), taking at most
`maxMb` megabytes by the engine's estimate, 64 by default, or in a pool with `maxHeapMb`, the session's share of a
thread's heap (`maxHeapMb / (8 × maxSessionsPerThread)`, `docs/async.md`). Beyond either limit, the oldest steps are
dropped; a step larger than `maxMb` on its own is not kept, and the steps before it are dropped too, just as Excel
clears what it cannot undo.

```ts
interface UndoOptions {
  depth?: number
  maxMb?: number
}
```

### `UndoResult`

A session step that was undone or redone ([`Session.undo`](#session), [`Session.redo`](#session)): the edit's result, as the call returned it.

```ts
interface UndoResult extends EditResult {
  /** The call that made the change. */
  call: UndoCall
  /** For the calls that change cells: the cells whose input was restored, so that a host can read their notes and hyperlinks again. */
  cells?: CellRef[]
  /**
   * `editStructure`: where a cell of the workbook before this call is now, or null when no cell of the workbook has it or
   * its sheet is not in that workbook; an interface that moves its displayed values with it and applies `changes` gets
   * the new values, as after `editStructure`.
   */
  mapCell?(ref: CellRef): CellRef | null
}
```

### `ValidateContext`

What [`validateFormula`](#xlcalc-validateformula) checks a formula against: the workbook's sheets, names and tables; whatever is omitted is not checked.

```ts
interface ValidateContext {
  /** Sheets of the workbook; without them, sheets are not checked. */
  sheets?: string[]
  /** Names of the workbook; without them, names are not checked. */
  names?: string[]
  externalFunctions?: string[]
  /**
   * Tables of the workbook with their columns; without them, structured references are not checked. With them, the
   * canonical text spells tables and columns as they are named, and a bare table name refers to a table, not a defined
   * name.
   */
  tables?: { name: string; columns: string[] }[]
  /**
   * The table containing the formula's cell: a structured reference without a table name refers to it (`[@Qty]`), as in
   * a cell of the table, and the canonical text names the table (`Sales[[#This Row],[Qty]]`). Without it, such a
   * reference is rejected when `tables` is given (`unknown_name`), and otherwise stays without a table name.
   */
  table?: string
  /**
   * The workbook's links to other workbooks (`BookInput.externalBooks`): a reference to one of them is checked against
   * its sheets and names. A reference to a workbook not among the links is rejected (`external_ref`), with or without
   * them.
   */
  externalBooks?: ExternalBook[]
}
```

### `ValidateResult`

What [`validateFormula`](#xlcalc-validateformula) returns: the formula's canonical text and what it refers to, or why and where it fails to parse.

```ts
type ValidateResult =
  /** `tables`: the tables named by the formula's structured references (and bare table names, with `ctx.tables`). */
  | { ok: true; canonical: string; functions: string[]; sheets: string[]; names: string[]; externals: string[]; tables: string[] }
  /** `fn`: the function, upper-case, a call to an unknown function (`unknown_function`) or an unsupported call (`unsupported`) names. */
  | (Refusal<CheckCode> & { at: number; fn?: string })
```

### `ValidationAlert`

A value that a cell's data validation rule does not allow ([`Session.setCellText`](#session)): the rule, and its alert's style and message.

```ts
interface ValidationAlert {
  valid: false
  rule: DataValidation
  severity: 'stop' | 'warning' | 'information'
  message: string
}
```

### `ValidationOperator`

How a value is compared with `formula1` (and `formula2`): between and not between use both, inclusive.

```ts
type ValidationOperator = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual'
```

### `ValidationRefusal`

A value rejected because a cell's data validation rule does not allow it ([`Session.validateEntry`](#session)): the rule, and its alert's style and message.

```ts
type ValidationRefusal = Refusal<'validation'> & { rule: DataValidation; severity: ValidationAlert['severity'] }
```

### `ValidationType`

What a data validation rule allows in a cell: any value (`any`, a rule that only shows an input message), a whole
number, a decimal number, a date or a time compared with the rule's formulas, text whose length is compared the same
way, an item of a list, or a value for which a formula (`custom`) returns TRUE.

```ts
type ValidationType = 'any' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'list' | 'custom'
```

### `ValuesOptions`

Options of [`Session.getValues`](#session).

```ts
interface ValuesOptions {
  /** Maximum number of cells in the rectangle, 1,000,000 by default. A larger one is rejected, never truncated. */
  maxCells?: number
}
```

### `Viewport`

What a grid needs to draw a viewport of a sheet ([`Session.getViewport`](#session)).

```ts
interface Viewport {
  /**
   * The rectangle returned (`B2:D9`): the requested one, or, when it holds more than `maxCells` cells, the part up to
   * its first rows that fit (of `ViewportOptions.rows`, when given).
   */
  range: string
  /** The requested viewport held more than `maxCells` cells: `range` is the part returned. */
  truncated: boolean
  /** By address, row by row: every cell with something to display, including an empty cell with its row's or column's style or number format, a conditional format's appearance, a note or a hyperlink. */
  cells: Record<string, ViewportCell>
  /** The styles the viewport's cells have or use (`ViewportCell.style`, `effectiveStyle`), by their index in the workbook's `styles`: copies owned by the result. */
  styles: Record<number, CellStyle>
  /** The columns of `range` from left to right, or those of `ViewportOptions.columns`. */
  columns: ViewportColumn[]
  /** The rows of `range` from top to bottom, or those of `ViewportOptions.rows`. */
  rows: ViewportRow[]
  /** The merged areas that overlap the viewport's cells, in full. */
  merges: string[]
  /** The tables that overlap the viewport's cells, in full, in the sheet's order. */
  tables: ViewportTable[]
  /** The sheet's frozen panes (`SheetInput.frozen`): its first `rows` rows and `columns` columns stay visible, 0 for none. */
  frozen: { rows: number; columns: number }
  /** Where the sheet's cells end, as returned by `Session.getUsedRange`. */
  usedRange: SheetExtent
  /** The session revision that the viewport reflects (`Session.revision`): a host applies the changes of later edits on top of it (`EditResult.revision`). */
  revision: number
}
```

### `ViewportCell`

A cell of a viewport ([`Session.getViewport`](#session)): its value as `getResult()` returns it, with the cell's formatting fully
resolved: its style and number format, including those inherited from rows and columns, the effect of its
conditional formats, its note and its hyperlink.

```ts
interface ViewportCell extends Partial<CellDisplay> {
  value: CellValue
  /** The canonical formula text, as in `CellResult.formula`. */
  formula?: string
  /** With `formulaLocal`: the formula as Excel displays it in the workbook's locale (`Range.FormulaLocal` without `=`), and inside its table the way Excel displays it there (`[@Qty]`). */
  formulaLocal?: string
  spill?: string
  spillFill?: CellValue
  spillOf?: string
  /** On the top-left cell of a legacy array formula, its range (`CellInput.array`). */
  array?: string
  /** As `CellResult.meta`: on the session's thread the workbook's own object, which must not be modified. */
  meta?: JsonValue
  /** The cell's own style, an index into the workbook's `styles` as in `CellView.style`; the style object is in `Viewport.styles`. */
  style?: number
  /**
   * The style the cell uses, an index into the workbook's `styles` (the style object is in `Viewport.styles`): its own,
   * or else its row's, or else its column's. It is always a style in the table: a cell where a styled row crosses a
   * styled column uses the row's, as Excel displays an empty cell from a file there, and a combination of the two would
   * be a style of the cell's own (`Session.patchCells`).
   */
  effectiveStyle?: number
  /** The number format the cell is displayed in, as a format code: its own, or else its row's, or else its column's; a built-in format as its code in the workbook's locale (14 is `dd.mm.yyyy` in ru-RU). */
  numberFormat?: string
  /** Where `numberFormat` comes from. */
  numberFormatSource?: 'cell' | 'row' | 'column'
  /** The effect of the sheet's conditional formats on the cell, given the current values; when they set a number format, `text` and `color` display the value in that format. */
  appearance?: ConditionalAppearance
  note?: CellNote
  link?: CellLink
  /** With `diagnostics`, on a cell with an error or a formula stopped by a circular reference: as returned by `Session.getCellIssue`. */
  issue?: CellIssueDetail
  /**
   * With `editText`: the text the cell opens with in edit mode, as Excel's formula bar displays it: a formula with `=`
   * in the workbook's locale (`formulaLocal`); in every cell of a legacy array formula's range, its formula (without
   * the braces the bar shows around it); a value as `entryText` writes it in the cell's number format; and empty for a
   * cell without a value or formula of its own (an empty cell, or one that a dynamic array spilled into).
   */
  editText?: string
}
```

### `ViewportColumn`

A column of a viewport: its number (A is 1) and its width as in [`ColumnInput.width`](#columninput), or Excel's default where it has none.

```ts
interface ViewportColumn {
  column: number
  width: number
  hidden?: true
}
```

### `ViewportOptions`

Options of [`Session.getViewport`](#session) and [`viewportOf`](#xlcalc-viewportof).

```ts
interface ViewportOptions {
  /** Also return each formula as displayed in the workbook's locale (`ViewportCell.formulaLocal`); false by default. */
  formulaLocal?: boolean
  /**
   * Also return, for each cell with an error and each formula stopped by a circular reference, the reason and the cell
   * where it originated (`ViewportCell.issue`); false by default.
   */
  diagnostics?: boolean
  /** Maximum number of cells a viewport returns, 20,000 by default: for a larger one, the first rows that fit (`Viewport.truncated`). */
  maxCells?: number
  /**
   * Only these rows of `range` (for example, the rows a filter displays): their row numbers on the sheet, from 1, in
   * ascending order, each listed once. The viewport returns them in this order, with only the merged areas and tables
   * that overlap its cells.
   */
  rows?: number[]
  /** Only these columns of `range`, by number (A is 1), in ascending order, each listed once, as with `rows`. */
  columns?: number[]
  /** Also return, for each cell, the text it opens with in edit mode (`ViewportCell.editText`); false by default. */
  editText?: boolean
}
```

### `ViewportRow`

A row of a viewport: its number on the sheet, from 1, and its height in points, or Excel's default (15) where it has none.

```ts
interface ViewportRow {
  row: number
  height: number
  hidden?: true
}
```

### `ViewportTable`

A table that overlaps a viewport ([`Viewport.tables`](#viewport)): its name, its range, and whether it has header and totals rows.

```ts
interface ViewportTable {
  name: string
  /** The entire table, including its header and totals rows (`TableInput.range`). */
  range: string
  /** Its first row is a header row (`TableInput.headers`, true by default). */
  headers: boolean
  /** Its last row is a totals row (`TableInput.totals`, false by default). */
  totals: boolean
}
```

### `WorkbookStatistics`

Excel's Workbook Statistics ([`Session.getWorkbookStatistics`](#session)): the counts of a sheet and of the whole workbook.

```ts
interface WorkbookStatistics {
  /**
   * The sheet's counts and its end: the cell of the last row and the last column that hold data or a note, or that a
   * merged area over such a cell covers (`E9` for data in B9 and E2), not reached by styles, hyperlinks alone or tables
   * as the used range is; null on a sheet without.
   */
  sheet: StatisticsCounts & { end: string | null }
  /** The workbook's counts over all its sheets, hidden ones too, and the number of sheets. */
  workbook: StatisticsCounts & { sheets: number }
}
```

### `WorkerLike`

What the pool needs from a Web Worker ([`CalcPoolOptions.spawn`](#calcpooloptions)).

```ts
interface WorkerLike {
  postMessage(message: unknown): void
  terminate(): void
  addEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { data?: unknown; message?: string }) => void): void
}
```

### `WriteKeptOptions`

Options of [`writeKept`](#xlcalcxlsxkeep-writekept): the language of its issues and of its refusals ([`KeptRefusal`](#xlcalcxlsxkeep-keptrefusal)) and which formula values the
file stores, as [`writeXlsx`](#xlcalcxlsx-writexlsx) takes them, and the limits of reading the original file, as [`readXlsx`](#xlcalcxlsx-readxlsx) takes them.

```ts
interface WriteKeptOptions extends Pick<WriteXlsxOptions, 'locale' | 'values'> {
  /** How large the original file's parts may be once inflated (`ReadXlsxOptions.limits`). */
  limits?: XlsxLimits
}
```

### `WriteXlsxOptions`

Options of [`writeXlsx`](#xlcalcxlsx-writexlsx): the language of its issues, which formula values the file stores, and whether it stores formulas or only values.

```ts
interface WriteXlsxOptions {
  /** The language of the issue messages, as `readXlsx` accepts it (`ReadXlsxOptions.locale`). */
  locale?: string | Locale
  /**
   * The formula values the file stores. `require`: every formula's value and every value of its array, from the result
   * (an `InputError` when one is missing, since then the result is not this workbook's). `omit`: none, and Excel
   * calculates them when it opens the file. `preserve`: those the result has; a formula without one is reported as
   * `noCachedValues`. By default, `require` with a result and `omit` without one, reported as `noCachedValues` when the
   * workbook has formulas.
   */
  values?: 'require' | 'omit' | 'preserve'
  /**
   * `formulas` (the default): formulas, with their values according to `values`. `values`: values only; a formula's
   * cell gets its value from the result (which is then required).
   */
  mode?: 'formulas' | 'values'
}
```

### `XlsxCached`

The values a file stores for formulas and for the cells their arrays spilled into ([`XlsxRead.cached`](#xlsxread)): what Excel
calculated when it saved the file, by sheet in the workbook's order, in the shape of a result.
`cached.sheets[0].cells.B1.value` is read the same way as `compute(book).sheets[0].cells.B1.value`, so the two can be
compared cell by cell, and `{ book, result: cached }` writes the file again with Excel's values ([`writeXlsx`](#xlcalcxlsx-writexlsx)).
`compute(read.book, { cached: read.cached })` (and [`createSession`](#xlcalc-createsession)) opens the workbook as Excel opens the file: the
formulas stopped by a circular reference keep these values ([`CellInput.held`](#cellinput)).

```ts
interface XlsxCached {
  sheets: { name: string; cells: Record<string, XlsxCachedCell> }[]
}
```

### `XlsxCachedCell`

A stored value: a formula's, with `spill`, the range its array fills (for a legacy array formula, its range), or a
cell in an array's range, with `spillOf`, the address of its formula; as in [`CellResult`](#cellresult).

```ts
type XlsxCachedCell = Pick<CellResult, 'value' | 'spill' | 'spillOf'>
```

### `XlsxIssue`

One item lost or changed, and where: the sheet, and the cells (addresses or ranges) or the name it concerns.

```ts
interface XlsxIssue {
  code: XlsxIssueCode
  /** What was lost or changed, and why. */
  message: string
  sheet?: string
  cells?: string[]
  name?: string
  /** The function, upper-case, of a `userFunction` issue: one issue per function, sheet or name, its cells the calls. */
  fn?: string
}
```

### `XlsxIssueCode`

What the workbook cannot hold from a file, or what writing changed. The codes:

- `richText`: text with formatted runs, read as plain text.
- `name`: a name the workbook does not accept (a built-in `_xlnm.` name, one on a sheet that was left out), or a name
  the writer left out.
- `table`: a table the model cannot hold (overlapping another table or a merged area, or with a name already used),
  left out; a header cell that does not hold its column's name as text, given that name; a structured reference
  without a table name.
- `externalLink`: a link that is not to a workbook (DDE, OLE) or whose part is missing, kept as a workbook without
  sheets; a formula that refers to another workbook not among the links (`[3]Sheet!A1` in a file with two), or that
  the engine cannot parse.
- `userFunction`: calls to a function the file marks as user-defined (`_xludf.`), a macro or add-in function, named
  by `fn`.
- `formula`: a formula the engine cannot parse, read as written (its cell calculates to an error), or a formula the
  writer could not write (its value is written instead, or an empty cell when it has none).
- `errorConstant`: an error typed into a cell as its value, read as a formula of that error.
- `errorValue`: an error the model has no code for (`#GETTING_DATA`, `#FIELD!`, `#BLOCKED!` and the like), read
  as `#ERROR!`; or, when writing, `#CYCLE!` written as 0, as Excel stores a circular reference, and `#ERROR!`
  written as `#NULL!`, which Excel has.
- `dataTable`: a what-if data table (`TABLE()`), read as its stored values.
- `calculation`: calculation settings the workbook does not keep: manual mode, automatic except for data tables,
  precision as displayed; an iteration count or maximum change out of range, read as Excel's default.
- `style`: what a style cannot hold (pattern and gradient fills, diagonal borders, rotation and the like).
- `defaultFont`: a default font other than Calibri 11, kept in [`BookInput.font`](#bookinput) and measured as Calibri 11.
- `defaultWidth`: a sheet's default column width other than Excel's standard of 8 digits, which the model assumes
  (Excel for Mac starts its own workbooks at 10).
- `numberFormat`: a built-in East Asian date format the model does not have (50 to 58, read as its en-US code), or
  one this reader does not know (left out).
- `hidden`: hidden sheets, read as visible ones.
- `sheetKind`: a chart, dialog or macro sheet, left out.
- `cellMetadata`: a cell's rich value (`vm`: a data type, a picture in the cell), read as the value the cell stores.
- `autoFilter`: an AutoFilter's criteria and sort, a table's criteria, an advanced filter (the sheet's filter mode is
  read: [`SheetInput.filtered`](#sheetinput)); an AutoFilter over no range of cells or over a table, left out; when writing, the
  filter mode of a sheet without an AutoFilter whose tables have no header row, left out.
- `dataValidation`: a rule the model cannot hold (no range on the sheet, no formula that its type requires, an
  unknown type), left out; cells shared by two rules, kept by the earlier one; an IME mode; a list whose items contain
  commas, read as split at them; when writing, a rule whose formula the file cannot hold, left out.
- `outline`: automatic styles of an outline.
- `conditionalFormat`, `comment`, `hyperlink`, `drawing` (pictures, charts, shapes),
  `pivotTable`, `protection`, `frozenPanes`, `print`
  (headers and footers, page breaks, a manually made page setup), `macros`, `sparkline`, `connection` (data
  connections and queries): parts of the file the model does not have, left
  out. View settings (zoom, gridlines, the selection) and a page setup Excel writes by default are not reported.
- `locale`: a locale the package does not have, left out of the workbook.
- `meta`: when writing, cell `meta` has no place in a file.
- `textLength`: when writing, text longer than Excel's 32,767 characters, truncated, and a rule's texts beyond Excel's
  limits.
- `noCachedValues`: when writing, formulas without their values, which Excel calculates when it opens the file: no
  result was given (`values: 'omit'` by default), or the result has no value for them (`values: 'preserve'`).
- `part`: a part of the file that is missing or cannot be read.

```ts
type XlsxIssueCode =
  | 'richText'
  | 'name'
  | 'table'
  | 'externalLink'
  | 'userFunction'
  | 'formula'
  | 'errorConstant'
  | 'errorValue'
  | 'dataTable'
  | 'calculation'
  | 'style'
  | 'defaultFont'
  | 'defaultWidth'
  | 'numberFormat'
  | 'hidden'
  | 'sheetKind'
  | 'cellMetadata'
  | 'conditionalFormat'
  | 'dataValidation'
  | 'comment'
  | 'hyperlink'
  | 'drawing'
  | 'autoFilter'
  | 'pivotTable'
  | 'protection'
  | 'frozenPanes'
  | 'outline'
  | 'print'
  | 'macros'
  | 'sparkline'
  | 'connection'
  | 'locale'
  | 'meta'
  | 'textLength'
  | 'noCachedValues'
  | 'part'
```

### `XlsxLimits`

The limits [`readXlsx`](#xlcalcxlsx-readxlsx) sets on a file's parts once inflated ([`ReadXlsxOptions.limits`](#readxlsxoptions)): sizes in bytes, as whole numbers of at least 1, and a ratio.

```ts
interface XlsxLimits {
  /**
   * A single part: 536,870,912 (512 MiB) by default. A part is read as one JavaScript string, so a part over about
   * 512 MiB cannot be read whatever the limit.
   */
  entryBytes?: number
  /** All the parts read, together: 1,073,741,824 (1 GiB) by default. */
  totalBytes?: number
  /**
   * How many times its compressed size a part over 16 MiB may inflate to, at least 1: 200 by default. The parts of xlsx
   * files inflate some 5 to 40 times; a zip bomb's, a thousand times.
   */
  ratio?: number
}
```

### `XlsxRead`

A file read into a workbook.

```ts
interface XlsxRead {
  /** The workbook, as the engine accepts it. */
  book: Book
  /** The values the file stores for formulas and the cells their arrays spilled into, for comparison with the engine's result. */
  cached: XlsxCached
  issues: XlsxIssue[]
}
```

### `XlsxSnapshot`

What [`writeXlsx`](#xlcalcxlsx-writexlsx) writes: a workbook and the values of its formulas: a session's `getSnapshot()`, `{ book, result:
compute(book) }`, or a file that was read, with the values it stored, `{ book: read.book, result: read.cached }`.

```ts
interface XlsxSnapshot {
  book: BookInput
  result?: ComputeResult | XlsxCached
}
```

### `XlsxWritten`

What [`writeXlsx`](#xlcalcxlsx-writexlsx) returns: the file, and what the workbook held that the file could not.

```ts
interface XlsxWritten {
  /** The xlsx file. */
  data: Uint8Array
  issues: XlsxIssue[]
}
```
