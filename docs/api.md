# API reference

Every function, value and type the package exports, entry by entry, with the comments of its source. Guides with
examples are the [README](../README.md) and the other files of `docs/`. This file is made from the package's types
(`npx tsx test/api-reference.ts --write`): change the comments in the source, not here.

- [`xlcalc`](#xlcalc)
- [`xlcalc/a1`](#xlcalca1)
- [`xlcalc/text`](#xlcalctext)
- [`xlcalc/xlsx`](#xlcalcxlsx)
- [`xlcalc/worker`](#xlcalcworker)
- [`xlcalc/function-text`](#xlcalcfunction-text)
- [`xlcalc/function-text/ru-RU`](#xlcalcfunction-textru-ru)
- [`xlcalc/locales/ru-RU`](#xlcalclocalesru-ru)
- [`xlcalc/contract`](#xlcalccontract)
- [Types](#types)

## xlcalc

Types: see [Types](#types).

### <a id="xlcalc-builtinformatcode"></a>`builtinFormatCode`

```ts
function builtinFormatCode(id: number, locale?: string | Locale): string | undefined
```

The code a built-in number format (0 to 49, [`NumberFormat`](#numberformat) as a number) shows by in `locale` (a loaded locale's tag
or a locale object; en-US without one), in en-US syntax as a cell's `numberFormat` takes it; undefined for a number that
is not a built-in format's.

### <a id="xlcalc-calcerror"></a>`CalcError`

```ts
class CalcError extends Error
```

Why an asynchronous or limited call failed, by `code`; a `limit` names the limit, its maximum, the size and the formula's cell.

### <a id="xlcalc-callat"></a>`callAt`

```ts
function callAt(text: string, caret?: number, opts?: ScanOptions): CallAt | null
```

The innermost call whose brackets hold `caret` (the end of the text by default) of formula text `text`, and the argument
it is in; null outside any call, and for text that is no formula (not begun by `=`, `+` or `-`). `opts.locale`: the
locale the text is typed in (its list separator and names); without it the text is en-US's.

### <a id="xlcalc-canonvalue"></a>`canonValue`

```ts
function canonValue(value: CellValue): string
```

A value's canonical text, the one its line in the values hash holds: `n 1234.5` (15 significant digits, `-0` as 0),
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

Column letters of column `n` (1 is `A`, 16,384 is `XFD`).

### <a id="xlcalc-colnumber"></a>`colNumber`

```ts
function colNumber(letters: string): number | null
```

The number of a column by its letters in any case (`B` is 2), or null when they are no column of the sheet (`XFE`).

### <a id="xlcalc-columnwidthfrompoints"></a>`columnWidthFromPoints`

```ts
function columnWidthFromPoints(points: number): number
```

Whole points back to a stored width in 256ths; fractional points round down, as the model's geometry does.

### <a id="xlcalc-columnwidthtopoints"></a>`columnWidthToPoints`

```ts
function columnWidthToPoints(width?: number | null): number
```

A stored xlsx width in the library's Mac point model; null or undefined is the default 53 points.

### <a id="xlcalc-compute"></a>`compute`

```ts
function compute(book: BookInput, opts?: ComputeOptions): ComputeResult
```

Computes the whole book. Without `opts.clock` the clock is the current moment and a random seed; the
result records it in `clock`, and passing that clock again reproduces the result. Input of the wrong
shape throws [`InputError`](#xlcalc-inputerror); invalid formulas do not throw, they give error values with reasons in `diagnostics`.

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

The text of a formula copied from cell `from` to cell `to`, as `copyRange` pastes it: relative parts move by the
offset, absolute ones stay, a reference pushed past the sheet's edge is `#REF!`, only the changed references are
rewritten. With the book (its sheets' names and tables), a formula copied out of a table names the table in its
references to it, as Excel pastes it; without it the formula's own sheet names are taken as written.

### <a id="xlcalc-createinlinecalc"></a>`createInlineCalc`

```ts
function createInlineCalc(opts?: InlineCalcOptions): Calc
```

A calculator in the same thread ([`Calc`](#calc)): it computes in portions of `sliceMs` that give way to the event loop, with the
deadlines, cancellation and limits of a pool's calls; its sessions answer the same as a pool's.

### <a id="xlcalc-createsession"></a>`createSession`

```ts
function createSession(book: BookInput, opts?: SessionOptions): Session
```

Opens a session. Without options it keeps the clock of the moment it was opened (with a random seed)
until `setClock`; `{ live: true }` takes a fresh clock on every recomputation.

### <a id="xlcalc-datetoserial"></a>`dateToSerial`

```ts
function dateToSerial(text: string, opts?: DateSystemOptions): number
```

The serial number of a date and time as [`CellInput.date`](#cellinput) takes it (`2026-01-31` is 46053, `2026-01-31T18:00`
46053.75), in `opts.dateSystem` (1900 without one; 1904 counts from 1904-01-01).

### <a id="xlcalc-debuggeneratedcode"></a>`debugGeneratedCode`

```ts
function debugGeneratedCode(book: BookInput, opts?: ComputeOptions): string
```

The `eval` mode's generated code of the book in evaluation order, for reading while debugging; not part of the contract.

### <a id="xlcalc-displayvalue"></a>`displayValue`

```ts
function displayValue(value: CellValue, opts?: DisplayOptions): CellDisplay
```

A value as a cell shows it without a book (a cell being edited, say): by `numberFormat` (General without one), in
`locale` and `dateSystem`, in a column of `width` (no limit without one: a value is not narrowed to a column it was
not given), as [`CellResult.text`](#cellresult) and `color` show it.

### <a id="xlcalc-editstructure"></a>`editStructure`

```ts
function editStructure(book: BookInput, op: StructureOp): StructureResult
```

Restructures a book; the input book is not changed. A copy that pastes values (from a filtered sheet) computes
the book at the current moment.

### <a id="xlcalc-entrytext"></a>`entryText`

```ts
function entryText(value: CellValue, opts?: EntryTextOptions): string
```

The text a cell holding `value` opens with for editing, as Excel's formula bar shows it, in `opts.locale` by the cell's
`opts.numberFormat`: a number with its 15 significant digits and the locale's decimal separator (`1234.5`, `1234,5`), a
percent cell's number as a percent (`12%`), a date or a time cell's by the locale's short date and long time
(`1/31/2025  6:00:00 PM`, `31.01.2025  18:00:00`), a boolean and an error in the locale's words, text with an
apostrophe where typing it would read otherwise (`'007`). The value reads back from the text ([`parseEntry`](#xlcalc-parseentry)).

### <a id="xlcalc-entrywidth"></a>`entryWidth`

```ts
function entryWidth(value: CellValue, opts?: EntryWidthOptions): number | undefined
```

The width Excel gives a column as `value`, shown by `opts.numberFormat`, is typed into one of its cells (`test/excel/fit/`):
a number AutoFit would give more room widens a column whose width is the default or one Excel fitted (`bestFit`), as
AutoFit fits it, though the default width (Excel keeps a fitted width that equals it); undefined when the column keeps
its width (text overflows, General drops decimals, a width set by hand stays, a column never narrows).

### <a id="xlcalc-enus"></a>`enUS`

```ts
const enUS: Locale
```

en-US's formula text, the invariant syntax; the package's other locales are modules (`xlcalc/locales/ru-RU`).

### <a id="xlcalc-errortext"></a>`errorText`

```ts
function errorText(code: ErrorCode, locale?: string | Locale): string
```

An error as a cell shows it in `locale` (a loaded locale's tag or a locale object, en-US without one): `#DIV/0!`, `#ДЕЛ/0!` in ru-RU.

### <a id="xlcalc-escapecolumn"></a>`escapeColumn`

```ts
function escapeColumn(name: string): string
```

A column's name in a structured reference as Excel writes it: `'` before each `[`, `]`, `#` and `'` in it and before
an `@` it starts with (`x#y` is `x'#y`, `@at` is `'@at`), the rest as it is (`Unit Price`).

### <a id="xlcalc-fitcolumns"></a>`fitColumns`

```ts
function fitColumns(book: BookInput, result: ComputeResult, sheet: string, columns: string): Record<string, number | null>
```

The widths AutoFit gives columns `C` or `B:D` of the sheet `sheet` of a computed book, from the values of its result
(`compute(book)`), by column letter: null for the default width; a column with nothing to show is left out.

### <a id="xlcalc-fitrows"></a>`fitRows`

```ts
function fitRows(book: BookInput, result: ComputeResult, sheet: string, rows: string): Record<string, number | null>
```

The heights AutoFit gives rows `5` or `2:9` of the sheet `sheet` of a computed book, from the values of its result
(`compute(book)`), by spans of rows of one height (`2:4`): null for the default height.

### <a id="xlcalc-formatclipboard"></a>`formatClipboard`

```ts
function formatClipboard(rows: string[][], opts?: { trailingNewline?: boolean | undefined; }): string
```

Excel's plain-text copy: tabs, CRLF rows and quoted control characters; Windows ends with CRLF, Mac does not.

### <a id="xlcalc-function_categories"></a>`FUNCTION_CATEGORIES`

```ts
const FUNCTION_CATEGORIES: readonly FunctionCategory[]
```

The categories in the order Excel's window lists them.

### <a id="xlcalc-function_names"></a>`FUNCTION_NAMES`

```ts
const FUNCTION_NAMES: readonly string[]
```

The whitelist: every function a formula can call, by its English name, in alphabetical order.

### <a id="xlcalc-functionsignature"></a>`functionSignature`

```ts
function functionSignature(name: string, locale?: string | Locale): FunctionSignature | null
```

The signature of a function, by its name in English or in the locale (any case), for a tip of its arguments while a
formula is typed; null for a name that is no function of the whitelist.

### <a id="xlcalc-inputerror"></a>`InputError`

```ts
class InputError extends TypeError
```

Input of the wrong shape. `path` points at the offending part, for example `book.sheets[1].cells["B0"]`.

### <a id="xlcalc-internstyle"></a>`internStyle`

```ts
function internStyle(styles: readonly CellStyle[], style: CellStyle): { index: number; styles: CellStyle[]; }
```

A style's number in a style table, as `patchCells` finds it: the first equal style's number and the same array, or a
new array with a copy of the style at the end and its number. Equal styles have the same fields with the same values,
in any order. A table is read once per array: the next call given the answer's table reads only the style added. A
table changed in place (a style added, replaced or removed) is read again; a style changed in place is not seen.

### <a id="xlcalc-lineskey"></a>`linesKey`

```ts
function linesKey(first: number, last: number, axis: "columns" | "rows"): string
```

The key of lines `first` to `last` (`C`, `B:D`; `5`, `5:9`), in order whichever comes first.

### <a id="xlcalc-liveworkbook"></a>`liveWorkbook`

```ts
function liveWorkbook(session: Session): LiveWorkbook
```

The live workbook of a synchronous session ([`LiveWorkbook`](#liveworkbook)): the same object for the same session. An asynchronous
session throws [`InputError`](#xlcalc-inputerror); its writes cannot answer what they computed.

### <a id="xlcalc-locales"></a>`LOCALES`

```ts
const LOCALES: ReadonlyMap<string, Locale>
```

The loaded locales by tag, en-US and those registered, in the order they were loaded. The map is for reading.

### <a id="xlcalc-max_cols"></a>`MAX_COLS`

```ts
const MAX_COLS: 16384
```

A sheet's columns, as Excel's: 16,384, A to XFD.

### <a id="xlcalc-max_formula_depth"></a>`MAX_FORMULA_DEPTH`

```ts
const MAX_FORMULA_DEPTH: 64
```

Excel's limit of calls nested in a formula: 64.

### <a id="xlcalc-max_formula_length"></a>`MAX_FORMULA_LENGTH`

```ts
const MAX_FORMULA_LENGTH: 8192
```

Excel's limit of a formula's text: 8,192 characters.

### <a id="xlcalc-max_rows"></a>`MAX_ROWS`

```ts
const MAX_ROWS: 1048576
```

A sheet's rows, as Excel's: 1,048,576.

### <a id="xlcalc-namekey"></a>`nameKey`

```ts
function nameKey(name: string): string
```

The key a name compares by: sheets, tables and defined names are the same name in any case (`Sales`, `SALES`), by
Unicode's upper case in every locale (`straße` is `STRASSE`).

### <a id="xlcalc-needsapostrophe"></a>`needsApostrophe`

```ts
function needsApostrophe(text: string, opts?: EntryTextOptions): boolean
```

Whether text typed into a cell would read as something else than that text, so the cell needs an apostrophe before it
to hold it: a number, a date, a boolean, an error or a formula in the locale (`007`, `1.5` in ru-RU, `=A1`), a number
in a percent cell, text that begins with an apostrophe, empty text.

### <a id="xlcalc-numberformatcategory"></a>`numberFormatCategory`

```ts
function numberFormatCategory(format: NumberFormat, locale?: string | Locale): NumberFormatCategory
```

The category of a number format in `locale` (a loaded locale's tag or a locale object; en-US without one) as the
ribbon's Number Format box shows it, with its name there: a code as an xlsx file stores it, or a built-in format by
its number.

### <a id="xlcalc-numberformatchoices"></a>`numberFormatChoices`

```ts
function numberFormatChoices(locale?: string | Locale): FormatChoice[]
```

The items of the Number Format list in `locale`, in its order: each one's key, name, the format it applies and that format's code in the locale.

### <a id="xlcalc-outlinegroups"></a>`outlineGroups`

```ts
function outlineGroups(sheet: Pick<SheetInput, "columns" | "rows" | "outline">): SheetOutline
```

A sheet's groups of rows and columns as Excel draws them ([`SheetOutline`](#sheetoutline)), from its lines' outline levels, hidden
state and collapsed marks and its summary settings: a [`SheetInput`](#sheetinput), or a [`SheetLayout`](#sheetlayout) as `getLayout()` gives it. Each
axis's groups by first line, the outer before the inner where two begin together; a group is collapsed when Excel
keeps the mark on its summary line ([`RowInput.collapsed`](#rowinput)), whatever marks the sheet holds.

### <a id="xlcalc-parsecell"></a>`parseCell`

```ts
function parseCell(cell: string): CellPosition | null
```

A cell's row and column (`B12`, `$B$12`, `b12`), or null when the text is no cell of the sheet (`ZZZ1`, `A0`, `B2:C3`).

### <a id="xlcalc-parseclipboard"></a>`parseClipboard`

```ts
function parseClipboard(text: string): string[][]
```

Text from the clipboard as the rectangle Excel pastes ([`Session.pasteText`](#session)): its rows, at line breaks, of its fields, at
tabs, each row as wide as the longest; a field begun by a quote runs to its closing quote, line breaks and tabs inside
it, a doubled quote one; an empty field clears its cell, and a line break after the last row is a row of empty fields.

### <a id="xlcalc-parseentry"></a>`parseEntry`

```ts
function parseEntry(text: string, opts?: EntryOptions): EntryResult
```

Text typed into a cell as Excel takes it in the book's locale: the cell's new input, with the cell's style, meta, note
and link and the format Excel leaves it, or a formula Excel does not take. Empty text clears the value and keeps the rest.

### <a id="xlcalc-parselines"></a>`parseLines`

```ts
function parseLines(key: string, axis: "columns" | "rows"): LineSpan | null
```

The lines a key of [`SheetInput.columns`](#sheetinput) or `rows` names (`C`, `B:D`; `5`, `5:9`), or null for a key a book does not
take: lower case, a span of one line (`C:C`) or the other way round, past the sheet.

### <a id="xlcalc-parserange"></a>`parseRange`

```ts
function parseRange(range: string): RangeRect | null
```

A range's rectangle and kind: `B2:D9` and `B2` are cells, `B:D` whole columns, `2:9` whole rows, with `$` and letters
in any case; corners given the other way round are put in order, as Excel orders a reference typed so. Null for
anything else or a range that leaves the sheet.

### <a id="xlcalc-patchlines"></a>`patchLines`

```ts
function patchLines<T extends object>(map: Record<string, T> | undefined, span: string, patch: LinePatch<T>): Record<string, T>
```

A map of [`SheetInput.columns`](#sheetinput) or `rows` with `patch` applied to lines `span` (`C`, `B:D`; `5`, `5:9`, the axis by the
key's shape), as `setLayout` applies it: a value sets a field, null removes it, a field left out stays; spans the patch
cuts are split, lines left without a field are dropped. A new map; the one given is not changed. The map is checked as
a book's: an object under each span of the axis, no two spans overlapping.

### <a id="xlcalc-quotesheetifneeded"></a>`quoteSheetIfNeeded`

```ts
function quoteSheetIfNeeded(name: string, locale?: Pick<Locale, "booleans"> | undefined): string
```

A sheet's name in a reference as Excel writes it: in quotes, an apostrophe inside doubled, when it is not a plain name
(`'Q1 data'`, `'2024'`, `'O''Brien'`), when it is a cell of the sheet (`'A1'`, `'XFD1048576'`; not `XFE1`), when it
begins with an R1C1 reference (`'R'`, `'R1C1'`, `'R2D2'`) or when it is one of the locale's booleans (`'TRUE'`;
`'ИСТИНА'` with ru-RU). `locale` gives the booleans, en-US's (the stored text, `Range.Formula`) without one. The engine's
stored formulas quote every sheet: this is the text a person reads.

### <a id="xlcalc-rangeaddress"></a>`rangeAddress`

```ts
function rangeAddress(rect: CellRect): string
```

A rectangle's address as Excel writes it: `B2` for one cell, `B2:D9`, `B:D` over every row, `2:9` over every column.

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

What an arrow does at `caret` (the end of the text by default) of `text` typed over a cell: a reference put in at the
caret, the text entered, or nothing ([`PointAction`](#pointaction)). `opts.locale`: the locale the text is typed in, a loaded locale's
tag or a locale object; without it the text is en-US's.

### <a id="xlcalc-registerlocale"></a>`registerLocale`

```ts
function registerLocale(locale: Locale): void
```

Loads a locale for books in it ([`BookInput.locale`](#bookinput)) and for its tag wherever a locale is taken: one of the package's
locale modules (`import { ruRU } from 'xlcalc/locales/ru-RU'`, then `registerLocale(ruRU)`). Loading the same object
again does nothing; another object with a loaded tag is an [`InputError`](#xlcalc-inputerror). A pool's threads load it too.

### <a id="xlcalc-samename"></a>`sameName`

```ts
function sameName(a: string, b: string): boolean
```

Whether two names of sheets, tables or defined names are the same name ([`nameKey`](#xlcalc-namekey)).

### <a id="xlcalc-samevalue"></a>`sameValue`

```ts
function sameValue(a: CellValue | undefined, b: CellValue | undefined): boolean
```

Whether two values are the same value to the values hash: numbers equal to 15 significant digits (`0.1 + 0.2` is
`0.3`, `-0` is 0), errors by their code; an absent value (`undefined`) is an empty cell, as null.

### <a id="xlcalc-scanreferences"></a>`scanReferences`

```ts
function scanReferences(formula: string, opts?: ScanOptions): FormulaReference[]
```

The references to cells, whole columns and rows, and tables in formula text, in order, with their place in it (`at` to
`end`), for colouring them as they are typed. `formula` is a stored formula, or typed text in `opts.locale`; a leading
`=` is skipped. A token that does not read (a stray character, an error literal half typed, a sheet with nothing after
it) is skipped and the text read on past it; an open string, quoted sheet name or bracket takes the rest of the text.
Strings, function names and defined names are no references.

### <a id="xlcalc-serialtodate"></a>`serialToDate`

```ts
function serialToDate(serial: number, opts?: DateSystemOptions): string
```

The date and time of a serial number in `opts.dateSystem`, as [`CellInput.date`](#cellinput) takes it and with no time zone: the
date alone for a whole number (`2026-01-31`), with the time to the millisecond otherwise (`2026-01-31T18:00`,
`2026-01-31T18:00:30.5`). Serial 0 of the 1900 system is `1900-01-00` and 60 is `1900-02-29`, as Excel shows them.

### <a id="xlcalc-toinvariant"></a>`toInvariant`

```ts
function toInvariant(text: string, where: string | Locale, opts?: TranslateOptions): Translation
```

Text typed in a locale (a loaded locale's tag, or a locale object) as a stored formula: `=СУММ(A1:A3;1,5)` is
`SUM(A1:A3,1.5)`. A leading `=` is dropped, as stored formulas have none; English function names, TRUE/FALSE and
error texts are read too, and a sheet's name with quotes or without, kept as typed. Positions of errors are in `text`, their messages in the locale's language. `opts.table`, the
table of the formula's cell, names it before a structured reference without a table name (`[@Qty]` is `Sales[@Qty]`),
as [`toLocal`](#xlcalc-tolocal) drops it, so the formula reads outside the table too.

### <a id="xlcalc-tolocal"></a>`toLocal`

```ts
function toLocal(formula: string, where: string | Locale, opts?: TranslateOptions): Translation
```

A stored formula as a locale shows it (a loaded locale's tag, or a locale object): `SUM(A1:A3,1.5)` is
`СУММ(A1:A3;1,5)` in ru-RU, and a sheet is in quotes only where Excel puts them (`'Data'!B2` is `Data!B2`, in en-US
too; [`quoteSheetIfNeeded`](#xlcalc-quotesheetifneeded) with the locale's booleans). Structured references are shown as Excel shows them, the formula's row as `@`
(`Sales[[#This Row],[Qty]]` is `Sales[@Qty]`) and a table alone by its name (`Sales[]` is `Sales`); `opts.table`, the
table of the formula's cell, drops its name before a column of its own (`[@Qty]`, `SUM([Qty])`), as Excel shows a
formula inside a table.

### <a id="xlcalc-validateformula"></a>`validateFormula`

```ts
function validateFormula(text: string, ctx?: ValidateContext): ValidateResult
```

Checks formula text as a cell stores it (without `=`, en-US syntax): its canonical text and the functions, sheets, names,
external books and tables it names, or why it does not read and where ([`ValidateResult`](#validateresult)). With `ctx`, the sheets, names
and tables it names are checked against the book's.

### <a id="xlcalc-validatelocalformula"></a>`validateLocalFormula`

```ts
function validateLocalFormula(text: string, where: string | Locale, ctx?: ValidateContext): ValidateResult
```

[`validateFormula`](#xlcalc-validateformula) for text typed in a locale (a loaded locale's tag, or a locale object): the same result, with the
canonical stored text, positions of errors in `text` (a leading `=` is allowed) and messages in the locale's language.

### <a id="xlcalc-validatename"></a>`validateName`

```ts
function validateName(name: string, opts?: NameOptions): NameCheck
```

Excel's rules for a name ([`NameInput.name`](#nameinput)) in a book of `opts.locale` (en-US without one; a locale's words for TRUE
and FALSE are refused too, ИСТИНА and ЛОЖЬ in ru-RU).

### <a id="xlcalc-validatenames"></a>`validateNames`

```ts
function validateNames(names: NameInput[], ctx: NamesContext): NameIssue[]
```

The issues of a book's names, by the same check as a book's (`prepareNames`), without the book.

### <a id="xlcalc-validatesheetname"></a>`validateSheetName`

```ts
function validateSheetName(name: string, opts?: NameOptions): SheetNameCheck
```

Whether Excel takes a sheet name, by the rules a book's sheets are checked by: 1 to 31 characters, none of
`: \ / ? * [ ]`, no apostrophe at either end, not `History` in any case. A refusal's message is in `opts.locale`'s
language, English without one; `at` is the place of the character Excel does not take.

### <a id="xlcalc-viewportof"></a>`viewportOf`

```ts
function viewportOf(saved: BookSnapshot, sheet: string, range: string, opts?: SavedViewportOptions): SavedViewport
```

A window over a saved book and its result, as [`Session.getViewport`](#session) gives it over a session of that book, without
computing its formulas: the result's values, text and colour, the book's formats, styles and layout; with
`appearance`, its conditional formats computed over the result's values.

## xlcalc/a1

Described under an entry above: [`cellAddress`](#xlcalc-celladdress), [`colLetters`](#xlcalc-colletters), [`colNumber`](#xlcalc-colnumber), [`linesKey`](#xlcalc-lineskey), [`MAX_COLS`](#xlcalc-max_cols), [`MAX_ROWS`](#xlcalc-max_rows), [`nameKey`](#xlcalc-namekey), [`outlineGroups`](#xlcalc-outlinegroups), [`parseCell`](#xlcalc-parsecell), [`parseLines`](#xlcalc-parselines), [`parseRange`](#xlcalc-parserange), [`patchLines`](#xlcalc-patchlines), [`quoteSheetIfNeeded`](#xlcalc-quotesheetifneeded), [`rangeAddress`](#xlcalc-rangeaddress), [`rectContains`](#xlcalc-rectcontains), [`rectsMeet`](#xlcalc-rectsmeet), [`sameName`](#xlcalc-samename).

## xlcalc/text

Described under an entry above: [`callAt`](#xlcalc-callat), [`columnWidthFromPoints`](#xlcalc-columnwidthfrompoints), [`columnWidthToPoints`](#xlcalc-columnwidthtopoints), [`escapeColumn`](#xlcalc-escapecolumn), [`formatClipboard`](#xlcalc-formatclipboard), [`numberFormatCategory`](#xlcalc-numberformatcategory), [`numberFormatChoices`](#xlcalc-numberformatchoices), [`parseClipboard`](#xlcalc-parseclipboard), [`quoteSheetIfNeeded`](#xlcalc-quotesheetifneeded), [`refAt`](#xlcalc-refat), [`scanReferences`](#xlcalc-scanreferences).

## xlcalc/xlsx

Types: see [Types](#types).

### <a id="xlcalcxlsx-fromlegacy"></a>`fromLegacy`

```ts
function fromLegacy(formula: string, ctx?: LegacyContext): string
```

A formula of the old kind (text without `=`, as a file stores it, after the prefixes of newer functions) as the
dynamic formula Excel 365 shows for it, in canonical text: `A1:A10*2` is `@A1:A10*2`, `INDEX(A1:C3,0,1)` is
`@INDEX(A1:C3,0,1)`, `SUM(A1:A10)` stays. Excel puts `@` before a name by what the name's formula gives, so the
book's names are needed for formulas that use names. A formula the engine does not parse is returned as it is. The
function of the API (`strict-args.ts`), from `xlcalc/xlsx` and the root alike.

### <a id="xlcalcxlsx-readxlsx"></a>`readXlsx`

```ts
function readXlsx(data: ArrayBuffer | Uint8Array<ArrayBufferLike>, opts?: ReadXlsxOptions): Promise<XlsxRead>
```

Reads an xlsx file (bytes of the file) into a book, the values Excel stored and what the book could not hold, its
messages in `opts.locale` (en-US by default).

### <a id="xlcalcxlsx-writexlsx"></a>`writeXlsx`

```ts
function writeXlsx(snapshot: XlsxSnapshot, opts?: WriteXlsxOptions): Promise<XlsxWritten>
```

Writes a book as an xlsx file: its formulas with the values of `snapshot.result` that `opts.values` asks for, or its
values alone (`opts.mode: 'values'`).

## xlcalc/worker

Described under an entry above: [`CalcError`](#xlcalc-calcerror), [`MAX_FORMULA_DEPTH`](#xlcalc-max_formula_depth), [`MAX_FORMULA_LENGTH`](#xlcalc-max_formula_length).

### <a id="xlcalcworker-createcalcpool"></a>`createCalcPool`

```ts
function createCalcPool(opts?: CalcPoolOptions): Calc
```

A pool of threads. Node runs `worker_threads` with `maxHeapMb` as each thread's heap; a browser runs Web Workers,
started by `spawn` or by the thread's module next to this one, and has no memory limit.

## xlcalc/function-text

### <a id="xlcalcfunction-text-functiondescription"></a>`functionDescription`

```ts
function functionDescription(fn: string, locale?: string | Locale): string | null
```

What a function does, by its English name ([`FunctionSignature.fn`](#functionsignature), [`CallAt.fn`](#callat)) in any case: one line in the locale's
words when the locale has them (`{ ...ruRU, ...ruRUFunctionText }`), else in en-US; null for a name outside the whitelist.

## xlcalc/function-text/ru-RU

### <a id="xlcalcfunction-textru-ru-rurufunctiontext"></a>`ruRUFunctionText`

```ts
const ruRUFunctionText: Required<Pick<Locale, "argumentWords" | "functionText">>
```

## xlcalc/locales/ru-RU

### <a id="xlcalclocalesru-ru-ruru"></a>`ruRU`

```ts
const ruRU: Locale
```

ru-RU, for [`registerLocale`](#xlcalc-registerlocale), and for [`toLocal`](#xlcalc-tolocal) and the functions without a book as it is.

## xlcalc/contract

Described under an entry above: [`MAX_FORMULA_DEPTH`](#xlcalc-max_formula_depth), [`MAX_FORMULA_LENGTH`](#xlcalc-max_formula_length).

## Types

### `AbortSignalLike`

What a computation needs of an AbortSignal: the platform's `AbortSignal` is one.

```ts
interface AbortSignalLike {
  readonly aborted: boolean
  readonly reason?: unknown
  addEventListener(type: 'abort', listener: () => void, options?: { once?: boolean }): void
  removeEventListener(type: 'abort', listener: () => void): void
}
```

### `AddressMove`

How a structure operation moved a book's addresses ([`EditResult.moved`](#editresult)). `lines`: the lines from `at` (a row's number,
or a column's, A being 1) on move by `count`: down or right by the lines inserted before `at`, up or left when
`-count` lines from `at` on were deleted. `range`: the cells of `from` moved so that its top-left cell is at `to`,
what was there gone. A sheet renamed, deleted, inserted or moved: `index`, its place among the sheets, from 0.

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

A cell's alignment ([`CellStyle.alignment`](#cellstyle)), as xlsx `<alignment>`: horizontal and vertical, text wrapped, indent.

```ts
interface AlignmentStyle {
  horizontal?: 'general' | 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed'
  vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed'
  wrap?: boolean
  /** Indent level, 0 to 250. */
  indent?: number
}
```

### `AsyncComputeOptions`

Options of [`compute`](#xlcalc-compute) of a pool or of [`createInlineCalc`](#xlcalc-createinlinecalc): those of [`compute`](#xlcalc-compute) and the call's deadline and signal.

```ts
type AsyncComputeOptions = ComputeOptions & CalcOptions
```

### `AsyncEditOptions`

The options of an asynchronous call that changes a session's book.

```ts
type AsyncEditOptions = CalcOptions & EditOptions
```

### `AsyncSession`

A session computed in portions or in a thread: the methods of [`Session`](#session), as promises. Calls are taken in order,
each after the one before has finished. A call that fails (a deadline, an abort, a limit) changes nothing.

```ts
interface AsyncSession {
  /** Mixed cells, names and looks applied atomically, with one final recalculation. */
  applyChanges(operations: SessionOperation[], opts?: ChangesOptions & CalcOptions): Promise<ChangesResult>
  /** Dependent-value changes and errors without changing the session; use its clock when applying. */
  previewChanges(operations: SessionOperation[], opts?: PreviewChangesOptions & CalcOptions): Promise<ChangesPreview>
  /** One revision, multiple ranges, only the fields and number of cells requested. */
  readCells(ranges: RangeRef[], opts?: ReadCellsOptions & CalcOptions): Promise<ReadCellsResult>
  /** Search performed where the book lives, literal or by Excel's Find; output is paged, in Find's order. */
  findCells(ranges: RangeRef[], query: string, opts?: FindCellsOptions & CalcOptions): Promise<ReadCellsResult>
  /** Excel's Replace All, as `Session.replaceInCells`. */
  replaceInCells(ranges: RangeRef[], query: string, replacement: string, opts?: ReplaceOptions & CalcOptions): Promise<ReplaceResult>
  /** Immediate formula dependents; source ranges stay ranges. */
  getDependents(ranges: RangeRef[], opts?: DependentsOptions & CalcOptions): Promise<DependentsResult>
  /** Totals over nonblank stored values and spills, without returning their cells. */
  summarizeRanges(ranges: RangeRef[], opts?: EditOptions & CalcOptions): Promise<RangeSummary>
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
  setClock(clock: Clock, opts?: AsyncEditOptions): Promise<EditResult>
  editStructure(op: StructureOp, opts?: AsyncEditOptions): Promise<AsyncStructureResult>
  setStyles(styles: CellStyle[], opts?: AsyncEditOptions): Promise<EditResult>
  setLayout(edits: LayoutEdit[], opts?: AsyncEditOptions): Promise<EditResult>
  /** The session's book; from a thread, a copy. */
  getBook(): Promise<Book>
  /** As `Session.getLayout`; from a thread, a copy, without the cells. */
  getLayout(): Promise<BookLayout>
  getUsedRange(sheet: string): Promise<SheetExtent>
  getEnd(sheet: string, cell: string, direction: EndDirection): Promise<string>
  explain(sheet: string, cell: string, opts?: ExplainOptions & CalcOptions): Promise<Explanation>
  /** As `Session.getCells`; from a thread only the rectangles come back, in one call. */
  getCells(sheet: string, range: string | string[]): Promise<Record<string, CellView>>
  /** As `Session.copyText`; the text is prepared in one read in the session's thread. */
  copyText(sheet: string, range: string, opts?: CopyTextOptions & CalcOptions): Promise<string>
  getCell(sheet: string, cell: string): Promise<CellView | undefined>
  /** As `Session.getViewport`; from a thread the window comes back in one call. */
  getViewport(sheet: string, range: string, opts?: ViewportOptions): Promise<Viewport>
  getCellIssue(sheet: string, cell: string): Promise<CellIssueDetail | null>
  getValidation(sheet: string, cell: string): Promise<CellValidation | null>
  validateEntry(sheet: string, cell: string, input: CellInput | null, opts?: CalcOptions): Promise<EntryCheck>
  getResult(opts?: CalcOptions): Promise<ComputeResult>
  /** As `Session.getSnapshot`, an independent copy in one call, so no edit comes between the book and its result. */
  getSnapshot(opts?: CalcOptions): Promise<BookSnapshot>
  /** As `Session.getHash`: from a thread only the hash comes back. */
  getHash(opts?: CalcOptions): Promise<string>
  undo(opts?: AsyncEditOptions): Promise<UndoResult | null>
  redo(opts?: AsyncEditOptions): Promise<UndoResult | null>
  /** The session's history as of its last finished call. */
  readonly history: SessionHistory
  /**
   * As `Session.subscribe`, synchronously: a listener runs when a call's answer has come, before the call's promise
   * settles. From a thread of the pool a change comes with the call's answer, so it reaches the listeners of the
   * session that made the call.
   */
  subscribe(listener: (change: SessionChange) => void): () => void
  close(): Promise<void>
}
```

### `AsyncSessionOptions`

`signal` and `timeoutMs` bound the opening (the first computation of the book).

```ts
type AsyncSessionOptions = SessionOptions & CalcOptions
```

### `AsyncStructureResult`

[`AsyncSession.editStructure`](#asyncsession): the same answer as [`Session.editStructure`](#session)'s.

```ts
type AsyncStructureResult = SessionStructureResult
```

### `AverageFormat`

Numbers above the average of the cells' numbers (below it with `below`, or equal too with `equal`), or `stdDev` standard deviations past it.

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

A data bar in a cell: where it is drawn, as fractions of the cell's width from the side its direction starts at.

```ts
interface BarLook {
  /** From 0 to 1; `start` < `end`. */
  start: number
  end: number
  color: Color
  gradient: boolean
  border?: Color
  /** The axis between negative and positive bars, a fraction of the width, for a rule with negative numbers. */
  axis?: number
  axisColor?: Color
  /** The bars run from right to left: the fractions count from the cell's right edge. */
  rightToLeft?: true
}
```

### `BlanksFormat`

The cell is blank (or holds spaces alone), is not, holds an error, or does not. Blanks go by the value, as Excel shows
them (an error is not blank); errors by `formula`, Excel's (`ISERROR(A1)`) without it, which Excel writes beside both.

```ts
interface BlanksFormat extends ConditionalBase {
  type: 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors'
  formula?: string
  style?: ConditionalStyle
}
```

### `Book`

A book as the calls that give one answer it ([`Session.getBook`](#session), `getSnapshot`, `getLayout`, [`editStructure`](#xlcalc-editstructure),
[`readXlsx`](#xlcalcxlsx-readxlsx)): a [`BookInput`](#bookinput), which goes back into every call that takes a book.

```ts
type Book = BookInput
```

### `BookInput`

A workbook as the engine takes it: its sheets with their cells, names, styles, locale and date system; what [`compute`](#xlcalc-compute), [`createSession`](#xlcalc-createsession) and [`writeXlsx`](#xlcalcxlsx-writexlsx) take and [`readXlsx`](#xlcalcxlsx-readxlsx) gives.

```ts
interface BookInput {
  /** The saved active visible sheet, by name; the writer defaults to the first visible sheet. Not part of inputHash. */
  activeSheet?: string
  /** Excel's rules for sheet names: 1 to 31 characters, none of `: \ / ? * [ ]`, no apostrophe at either end; one visible sheet at least. */
  sheets: SheetInput[]
  names?: NameInput[]
  /** The cell styles that cells, rows and columns refer to by index. */
  styles?: CellStyle[]
  /**
   * The locale the book is computed in, the BCP 47 tag of a loaded locale, `LOCALES` (`'en-US'` by default; another
   * of the package's locales is loaded by `registerLocale`). It changes computed values as Excel's regional settings do
   * (text read as a number or a date, numbers and booleans turned into text, `TEXT` format codes, currency), not the
   * formula text: formulas stay invariant, and `toLocal` shows them in a locale. A tag that is not loaded is an
   * `InputError` that names the module to import.
   */
  locale?: string
  /**
   * The date system, as xlsx `workbookPr date1904`: 1900 (the default) counts days from 1900-01-00 with Excel's
   * phantom 1900-02-29, 1904 from 1904-01-01, serial 0. It changes the serial numbers of the dates typed in and
   * computed, TODAY and NOW, date text and how dates show: a 1904 book shows a negative date or time with a minus
   * (`-12:00`). It goes into both hashes only when it is 1904.
   */
  dateSystem?: 1900 | 1904
  /**
   * Iterative calculation of circular references, as Excel's setting and xlsx `calcPr iterate`. Without it the cells of a
   * cycle and the cells that read it are not computed and hold their values (`CellInput.held`). With it the cells of
   * every cycle and the cells that depend on them are computed from 0, again and again, sheet by sheet in the order of
   * the sheets' names (in any case), row by row. That stops after `maxIterations` rounds (100 by default, 1 to 32,767)
   * or when no value changes by `maxChange` or more (0.001 by default). Each computation starts from 0 again, so a
   * session gives what a full computation gives, and `held` has no effect. Excel goes on from the values it holds
   * instead.
   */
  iteration?: { maxIterations?: number; maxChange?: number }
  /**
   * The other workbooks the book's formulas refer to (`[1]Sheet1!A1`, `'[1]My Sheet'!A1:B9`, `[1]!Rate`): element `i`
   * answers `[i+1]`. They are computed from the values stored here, as Excel computes a file it opens without updating
   * its links; the other book itself is never read. It goes into `inputHash` when it is not empty.
   */
  externalBooks?: ExternalBook[]
}
```

### `BookLayout`

A book without the cells of its sheets ([`Session.getLayout`](#session)): its styles, names and settings, and each sheet's layout.

```ts
type BookLayout = Omit<Book, 'sheets'> & { sheets: SheetLayout[] }
```

### `BookSnapshot`

A book and its result: a session's taken at one moment ([`Session.getSnapshot`](#session)), or a book and what [`compute`](#xlcalc-compute) gave for
it, as [`writeXlsx`](#xlcalcxlsx-writexlsx) takes them and [`viewportOf`](#xlcalc-viewportof) reads a version a host saved.

```ts
interface BookSnapshot {
  book: Book
  result: ComputeResult
}
```

### `BorderLine`

One side of a cell's border; the kinds are xlsx's.

```ts
interface BorderLine {
  style: 'thin' | 'medium' | 'thick' | 'dotted' | 'dashed' | 'double' | 'hair' | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot'
  color?: Color
}
```

### `CachedValues`

The values a file stores for its book's formulas, in the shape of a result: [`readXlsx`](#xlcalcxlsx-readxlsx)'s `cached`, or a result of the
book. They are read by sheet name and address, a cell's `value` alone.

```ts
interface CachedValues {
  sheets: { name: string; cells: Record<string, { value: CellValue }> }[]
}
```

### `Calc`

A calculator: [`createInlineCalc`](#xlcalc-createinlinecalc) (the same thread, in portions) or [`createCalcPool`](#xlcalcworker-createcalcpool) (threads).

```ts
interface Calc {
  compute(book: BookInput, opts?: AsyncComputeOptions): Promise<ComputeResult>
  createSession(book: BookInput, opts?: AsyncSessionOptions): Promise<AsyncSession>
  /** Refuses further calls (`closed`), the waiting and running ones too, and closes the sessions. */
  close(): Promise<void>
}
```

### `CalcErrorCode`

Why an asynchronous call failed ([`CalcError.code`](#xlcalc-calcerror)). `soft_timeout`: its `timeoutMs` passed, seen by a check inside
the computation. `aborted`: its `signal` fired. `hard_timeout`: it ran past the pool's `hardTimeoutMs` and its
thread was terminated. `out_of_memory`: its thread reached `maxHeapMb`. `thread_failed`: its thread exited or threw
outside the computation. `restore_failed`: the session could not be restored on a thread and is closed. `limit`:
the book or the edit is over [`CalcLimits`](#calclimits). `queue_full`: `maxQueue` tasks already wait. `too_many_sessions`: the pool holds
`maxSessions` sessions. `closed`: the calculator or the session is closed. `stale_revision`: an edit's `baseRevision`
is not the session's revision ([`EditOptions`](#editoptions)); it changed nothing.

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

Size limits of a computation and of an edit. A book over them is not computed and an edit over them is not made: the
call throws [`CalcError`](#xlcalc-calcerror) with code `limit`, so a result is either exact or absent and the values hash never depends on
the limits.

```ts
interface CalcLimits {
  /** Formulas in the book. */
  formulas?: number
  /** Values of one range read whole or one array that a formula makes, checked before they are made. */
  cells?: number
  /**
   * Values one spill may fill (its stored cells: the cells past a large array's last distinct value are not stored),
   * checked before they are written; `cells` when not given.
   */
  spill?: number
  /**
   * Cells one edit of a session writes with a value or a formula (a look alone goes to the sheet's looks), counted
   * before anything is written: a copy's source cells that hold one times the tiles of its target, a move's source cells
   * that hold one, the cells of `fillCellText`'s range and of `pasteText`'s rectangle, the cells `replaceInCells` enters
   * again, the edits of `setCells`, `replaceCells`, `setCellTexts` and `applyChanges`; their previews too. Operations
   * that write no more than the sheet holds (rows and columns inserted and deleted, `clearCells`) are not bounded by it.
   */
  writes?: number
}
```

### `CalcOptions`

Options of one asynchronous call.

```ts
interface CalcOptions {
  /** Stops the call: a waiting call at once, a running one at its next check. */
  signal?: AbortSignalLike
  /** A soft deadline: milliseconds of computing, counted from when the call starts computing (not while it waits). */
  timeoutMs?: number
}
```

### `CalcPoolOptions`

Options of [`createCalcPool`](#xlcalcworker-createcalcpool).

```ts
interface CalcPoolOptions {
  /** Threads; 2 by default. */
  threads?: number
  /** Node: the heap of each thread (`resourceLimits.maxOldGenerationSizeMb`); not available in a browser. */
  maxHeapMb?: number
  /** A task that runs longer, from when it reaches its thread, terminates the thread (`hard_timeout`). */
  hardTimeoutMs?: number
  /** The soft deadline of every call that gives none. */
  timeoutMs?: number
  /** Tasks waiting for a thread, all threads together; one more is refused (`queue_full`). */
  maxQueue?: number
  /** Open sessions of the pool; one more is refused (`too_many_sessions`). */
  maxSessions?: number
  /** Sessions a thread holds; past it the session called least recently leaves the thread and comes back on its next call. */
  maxSessionsPerThread?: number
  /**
   * The limits of every computation; a computation's own `limits` take their place field by field, the others stay. With `maxHeapMb`, `cells` is `maxHeapMb · 2^20 / 64` where not
   * given, and `spill` the smaller of `cells` and `maxHeapMb · 2^20 / 1024` where not given: one allocation larger than
   * what a thread's heap has left (a large array, the tables of a large spill's cells) stops the whole process, not the
   * thread.
   */
  limits?: CalcLimits
  /** Calls kept in a session's journal before the pool takes a new snapshot of its book; 64 by default. */
  snapshotEvery?: number
  /** Milliseconds of computing between two yields inside a thread, where a cancel is read; 50 by default. */
  sliceMs?: number
  /** Starts a thread in a browser, for a bundler that serves the thread's script elsewhere. */
  spawn?: () => WorkerLike
}
```

### `CallAt`

The call a caret stands in, in formula text as it is typed ([`callAt`](#xlcalc-callat)): the innermost function whose brackets hold the
caret, and the argument it is in, for a tip of the function's arguments ([`functionSignature`](#xlcalc-functionsignature)).

```ts
interface CallAt {
  /** The function's name as typed (`СуммЕсли`). */
  name: string
  /** Its English name in upper case, as `functionSignature` and stored formulas take it: the locale's name translated (`SUMIF`), another name as it is. */
  fn: string
  /** The argument the caret is in, from 0: the list separators before it at the call's own level. */
  index: number
  /** Where the call's `(` is, and its `)`, or null while it is open. */
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

A cell a call changed ([`EditResult.changes`](#editresult)): its address, its new value and, where they changed, its text, formula, look, note or link.

```ts
interface CellChange extends CellRef, Partial<CellDisplay> {
  value: CellValue
  spill?: string
  spillFill?: CellValue
  spillOf?: string
  /** The cell's formula, when it changed: canonical text, as `CellResult.formula`; null when the cell has none now. */
  formula?: string | null
  /** The cell's own style, when it changed, as `CellView.style`; null when it has none now (it shows its row's or column's). */
  style?: number | null
  /** The cell's own number format, when it changed; null when it has none now. */
  numberFormat?: NumberFormat | null
  /** A changed `numberFormat` other than null as a code, as `CellView.numberFormatCode`. */
  numberFormatCode?: string
  /** The cell's note, when its text or author changed; null when it was removed. */
  note?: CellNote | null
  /** The cell's hyperlink, when its url, location or tooltip changed; null when it was removed. */
  link?: CellLink | null
}
```

### `CellDisplay`

A value as a cell shows it: its text and the colour of its format's section.

```ts
interface CellDisplay {
  /**
   * What Excel shows in the cell (VBA's `Range.Text`): the value by its number format in the book's locale, General
   * narrowed to the column's width, `#` across the column for a number that does not fit, fills `*x` repeated
   * across it; text as it is (it overflows into the next cells). Derived from the value, the format, the locale and
   * the width: not part of the values hash.
   */
  text: string
  /** The colour of the format's section (`[Red]`, `[Color10]`), `#RRGGBB`; absent without one. */
  color?: Color
}
```

### `CellEdit`

A cell edit ([`Session.setCells`](#session)): the fields given set the cell's, null removes one, and the fields left out stay, as
Excel keeps a cell's look when a value is typed over it. The content is one: a new `value` removes the formula, its
`array` and its `held`, a new `formula` the value and the `held` (a formula entered where a circular reference stops it
holds 0). `input: null` clears the content and the hyperlink and keeps the rest, as Excel's
Delete key.

```ts
interface CellEdit extends CellRef {
  input: CellFields | null
}
```

### `CellFields`

Fields of a cell to set ([`Session.setCells`](#session)): a value sets the field, null removes it, a field left out stays.

```ts
type CellFields = { [K in keyof CellInput]?: CellInput[K] | null }
```

### `CellImpact`

A cell a previewed change would change ([`ChangesPreview`](#changespreview)): its value before and after.

```ts
interface CellImpact extends CellRef {
  before: CellValue
  after: CellValue
  /** The error after the operation, or the circular reference that stops the cell then, when there is one. */
  issue?: CellIssueDetail
}
```

### `CellInput`

An input cell. A date is passed explicitly (`{ date: '2025-01-31' }`) and becomes an Excel serial number.

```ts
interface CellInput {
  value?: Scalar | { date: string }
  /**
   * Formula text without `=`; a single leading `=` is dropped as the book is taken (`=SUM(A1:A2)` is `SUM(A1:A2)`, the
   * same input and `inputHash`; `==1` stays an error). When present, `value` is ignored.
   */
  formula?: string
  /**
   * Number format, an Excel format code as an xlsx file stores it (en-US syntax: `#,##0.00`, `dd.mm.yyyy`, `[Red]`),
   * shown in the book's locale (`CellResult.text`), or a built-in format by its number (`NumberFormat`). Without it
   * a cell shows its row's format, else its column's.
   */
  numberFormat?: NumberFormat
  /** A style of the book's `styles`, by index. Without it a cell shows its row's style, else its column's. */
  style?: number
  /** Arbitrary cell data (for example ids of source rows), passed through untouched. */
  meta?: JsonValue
  /**
   * An array formula of the old kind (entered with Ctrl+Shift+Enter; in xlsx `<f t="array" ref="B2:C5">` without
   * dynamic array metadata), on the top-left cell of its area: the area (`B2:C5`, or the cell alone, `B2`). The
   * formula is computed as an array, with no implicit intersection, and fills the area, which it never leaves:
   * a smaller array is repeated when it has one row or one column, else the cells past it are #N/A, a larger one is
   * cut, one value fills every cell; nothing spills and nothing blocks it. `ROW()` and `COLUMN()` give the rows and
   * columns of the area. The other cells of the area hold no input, and a session refuses to change some of them
   * without the rest (`Session.setCells`). Only with `formula`.
   */
  array?: string
  /** A note on the cell (Excel's comment of the old kind). Like a style, it goes with the cell and takes no part in computation. */
  note?: CellNote
  /** A hyperlink on the cell; the cell shows its value. Like a style, it goes with the cell and takes no part in computation. */
  link?: CellLink
  /**
   * The value a formula holds while a circular reference stops it. In a book without `iteration` a formula in a cycle,
   * or one that reads a cycle (through other cells too), is not computed, as Excel computes none of them: it shows
   * `held`, or 0 without it (a formula entered there shows 0 in Excel). Elsewhere the field has no effect, and the book
   * as the engine takes it leaves it out, so it counts in `inputHash` in such a cell alone. A session keeps it as Excel
   * keeps such cells: a cell that another edit brings into a cycle's way holds what it showed, a formula entered there
   * holds 0, a cell out of the way is computed again and holds nothing; `getBook()` gives it (`docs/cycles.md`). The
   * values a file stores come in as `held` through `cached` (`ComputeOptions.cached`).
   */
  held?: Exclude<CellValue, null>
}
```

### `CellIsFormat`

The cell's value compared with one formula's value, or two for `between` and `notBetween` (both ends included, either
way round), as `=`, `<` and the like compare; a blank cell a formula refers to is 0 there.

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

A cell with an error and its reason: an invalid formula (with the check code and position) or a computation error; or
a formula a circular reference stops (`error: 'CYCLE'`), whose value is the one it holds ([`CellInput.held`](#cellinput)).

```ts
interface CellIssue extends CellRef {
  error: ErrorCode
  reason: string
  check?: CheckCode
  /** Position in the formula text for `check`. */
  at?: number
}
```

### `CellIssueDetail`

Why a cell holds an error, or a circular reference stops it ([`Session.getCellIssue`](#session)): its code, reason, check code and
position as the result's `diagnostics` gives them, and `origin`, the cell where the error began, following the cells
its reason names (`from 'S'!A2`) to one whose error is its own; in a cycle, the first cell of the cycle the way
reaches, and so for a formula a cycle stops, through the cells it reads that the cycle stops too.

```ts
interface CellIssueDetail {
  error: ErrorCode
  reason: string
  check?: CheckCode
  at?: number
  origin: CellRef
}
```

### `CellLink`

A hyperlink, at least one of `url` and `location`. `url`: an address outside the book, as Excel's Address
(`https://example.com/`, `mailto:a@example.com`, a file). `location`: a place in the book, as Excel's SubAddress
(`'Sheet 2'!B12`, `Data!A1:C3`, a defined name), or with `url` a place in what it points to (the part after `#`).

```ts
interface CellLink {
  url?: string
  location?: string
  /** The text Excel shows when the pointer rests on the cell (ScreenTip). */
  tooltip?: string
}
```

### `CellNote`

A note: its author and its text, lines separated by `\n`.

```ts
interface CellNote {
  author?: string
  text: string
}
```

### `CellPatch`

A change of the look of a cell (`cell`) or a range (`range`: `B2:D9`, whole columns `B:D`, whole rows `2:5`)
([`Session.patchCells`](#session)): each field given sets the cells' own, null removes it, a field left out stays; values and
formulas stay as they are. A style by its fields ([`StylePatch`](#stylepatch)) goes over the style each cell shows, its own or else
its row's or column's, as Excel formats a cell; a style by its number of the book's `styles` replaces it.

```ts
type CellPatch = { sheet: string; set: PatchFields } & ({ cell: string; range?: never } | { range: string; cell?: never })
```

### `CellPosition`

A cell's row and column, numbered from 1 as A1 numbers them: `B12` is row 12, column 2.

```ts
interface CellPosition {
  row: number
  col: number
}
```

### `CellReadField`

A field `readCells` and `findCells` give of a cell: `editText` is the text it opens with for editing, as a window's ([`ViewportCell.editText`](#viewportcell)).

```ts
type CellReadField = 'value' | 'text' | 'formula' | 'editText' | 'style' | 'numberFormat' | 'note' | 'link' | 'issue'
```

### `CellRect`

A rectangle of a sheet: rows and columns numbered from 1, bounds inclusive; a whole column runs from row 1 to [`MAX_ROWS`](#xlcalc-max_rows).

```ts
interface CellRect {
  top: number
  left: number
  bottom: number
  right: number
}
```

### `CellRef`

A cell of a book: its sheet's name and its A1 address (`B2`).

```ts
interface CellRef {
  sheet: string
  cell: string
}
```

### `CellReplace`

A cell replaced whole ([`Session.replaceCells`](#session)): its new input, or null to remove the cell.

```ts
interface CellReplace extends CellRef {
  input: CellInput | null
}
```

### `CellResult`

A cell's value, and by default (`display`) what the cell shows: `text` and `color`, as [`CellDisplay`](#celldisplay).

```ts
interface CellResult extends Partial<CellDisplay> {
  /** The value: numbers stay numbers whatever the format shows. */
  value: CellValue
  /**
   * Canonical formula text (sheet always quoted, function names upper-case), with sheets, tables and defined names
   * spelled as the book names them, as Excel shows them (`ставка*2` with the name `Ставка` is `Ставка*2`).
   */
  formula?: string
  /** The cell's own number format, as its input has it. */
  numberFormat?: NumberFormat
  meta?: JsonValue
  /**
   * A formula whose array spilled: the range it fills on its sheet (`B2:D9`, this cell at the top-left); for an array
   * formula of the old kind (`CellInput.array`), its area.
   */
  spill?: string
  /**
   * The value of the cells of `spill` that the result does not list: only for an array of more than 65,536
   * values whose cells past some row and column all hold one value (`A:A*2`: 0 below the data), which are
   * then left out of `cells` and of the hash's lines but for one line of the range.
   */
  spillFill?: CellValue
  /**
   * A cell an array spilled into, or a cell of an old-kind array formula's area: the address of the array's formula
   * on the same sheet (an empty element is 0).
   */
  spillOf?: string
}
```

### `CellStyle`

A cell style of the book's table ([`BookInput.styles`](#bookinput)), as an xlsx cell format without its number format
(that is the cell's `numberFormat`). Styles do not take part in computation or in the values hash.

```ts
interface CellStyle {
  font?: FontStyle
  /** A solid fill. */
  fill?: { color: Color }
  border?: { top?: BorderLine; right?: BorderLine; bottom?: BorderLine; left?: BorderLine }
  alignment?: AlignmentStyle
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

[`Session.setCellText`](#session)'s and [`Session.setCellTexts`](#session)' options.

```ts
interface CellTextOptions extends EditOptions {
  /** False: the width a typed number widens its column to is given (`width`), not set; true by default, as Excel widens it. */
  widen?: boolean
  /**
   * True: a value a cell's rule of data validation does not take, of any style, is refused (`CellValidationRefusal`)
   * and nothing changes, so that a host shows Excel's alert before the value is written and calls again without the
   * option where the person keeps it (a warning's Yes, an information's OK). False by default: the value is written,
   * with its alert.
   */
  refuseInvalid?: boolean
  /**
   * True: the text was edited in the cell's own text (F2, a double click, the formula bar), not typed over the cell.
   * Excel's percent editor then adds `%` only to a number begun by a sign: in a cell of `0%`, `6%` is 0.06, `6` is 6
   * (shown 600%) and `-5` is -0.05. False by default.
   */
  inPlace?: boolean
}
```

### `CellTextPreview`

What [`Session.setCellText`](#session) would do ([`Session.previewCellText`](#session)): the cell's input, its value, the width and the alert, with nothing changed.

```ts
type CellTextPreview = { ok: true; input: CellInput | null; value: CellValue; width?: number; validation?: ValidationAlert } | CellTextRefused | CellValidationRefusal
```

### `CellTextRefusal`

Why [`Session.setCellText`](#session) did not take the text: a formula Excel does not take ([`CheckCode`](#checkcode)), part of an array formula, a formula in a table's header cell.

```ts
type CellTextRefusal = 'syntax' | 'empty' | 'too_long' | 'too_deep' | 'array' | 'header'
```

### `CellTextRefused`

Text [`Session.setCellText`](#session) did not take, and the formula Excel offers to correct it to, as [`parseEntry`](#xlcalc-parseentry)'s `correction`.

```ts
type CellTextRefused = Refusal<CellTextRefusal> & { correction?: string }
```

### `CellTextResult`

Text typed into a cell ([`Session.setCellText`](#session)): the edit's result, with the width the column took (`width`) and the
rule of data validation the value breaks (`validation`); or why the text was not taken, with its position in the text,
or, with `refuseInvalid`, the rule the value breaks.

```ts
type CellTextResult = (EditResult & { ok: true; width?: number; validation?: ValidationAlert }) | CellTextRefused | CellValidationRefusal
```

### `CellTextsPreview`

What [`Session.setCellTexts`](#session) would do ([`Session.previewCellTexts`](#session)): each cell's input and value, the widths and the alerts, with nothing changed.

```ts
type CellTextsPreview = { ok: true; cells: PreviewCell[]; widths?: ColumnWidth[]; validations?: CellValidationAlert[] } | CellTextsRefusal
```

### `CellTextsRefusal`

Why [`Session.setCellTexts`](#session) took none of the texts: the first one refused, as `setCellText` refuses it, with its cell;
with `refuseInvalid`, the cell whose value breaks the strictest rule of data validation once all are written (stop, then
warning, then information; the first cell of that style).

```ts
type CellTextsRefusal = (Refusal<CellTextRefusal> & CellRef) | CellValidationRefusal
```

### `CellTextsResult`

Texts typed into cells at once ([`Session.setCellTexts`](#session)): the edit's result, with the widths the columns took
(`widths`) and the cells whose values their rules of data validation do not take (`validations`); or the first text
refused, with its cell, and nothing written.

```ts
type CellTextsResult = (EditResult & { ok: true; widths?: ColumnWidth[]; validations?: CellValidationAlert[] }) | CellTextsRefusal
```

### `CellValidation`

A cell's data validation ([`Session.getValidation`](#session)): the rule over it, the rule's formulas as at the cell (relative
references moved from the rule's first cell), and for a list the items its dropdown offers.

```ts
interface CellValidation {
  rule: DataValidation
  formula1?: string
  formula2?: string
  /**
   * A list's items in order, computed in the book as it is now: a literal list's items as written (`text`, and `value` the
   * same text), or the cells of the reference its formula gives, each with its value and the text it shows by its number
   * format (blank cells past the last value left out). An error where the formula gives no reference or list. `input` is
   * what Excel writes into the cell when the item is picked from its dropdown, for `setCells`, which keeps the cell's
   * format: a cell's value as it is (`001` stays text, a number shown `50%` is 0.5), a literal item read as typed in
   * en-US, the locale a file holds the list in (`1/2` is the 2nd of January, `d-mmm`), null for a blank cell. Excel
   * checks no rule on a picked item.
   */
  list?: { value: CellValue; text: string; input: CellInput | null }[] | { error: ErrorCode }
}
```

### `CellValidationAlert`

A cell whose value its rule of data validation does not take ([`Session.setCellTexts`](#session)), with the rule's alert.

```ts
type CellValidationAlert = CellRef & ValidationAlert
```

### `CellValidationRefusal`

Typed text refused for the value a cell's rule of data validation does not take ([`CellTextOptions.refuseInvalid`](#celltextoptions)): the cell, the rule, its alert's style and message.

```ts
type CellValidationRefusal = ValidationRefusal & CellRef
```

### `CellValue`

A computed value. Dates come out as serial numbers, numbers are rounded to 15 significant digits.

```ts
type CellValue = Scalar | { error: ErrorCode }
```

### `CellView`

A cell of [`Session.getCells`](#session): as [`CellResult`](#cellresult), and the cell's own `style` (an index of the book's `styles`), `note` and
`link`, and the `appearance` its sheet's conditional formats give it (with a number format of theirs, `text` and
`color` show the value by it).

```ts
interface CellView extends CellResult {
  style?: number
  note?: CellNote
  link?: CellLink
  appearance?: ConditionalAppearance
  /** On the top-left cell of an array formula of the old kind, its area (`CellInput.array`); `spill` is that area too. */
  array?: string
  /** `numberFormat` as a code, as `ViewportCell.numberFormat` gives one: a built-in format by the code of the book's locale (14 is `dd.mm.yyyy` in ru-RU). */
  numberFormatCode?: string
}
```

### `ChangesOptions`

An explicit clock makes a preview and its application reproducible, including in a live session.

```ts
interface ChangesOptions extends EditOptions { clock?: Clock }
```

### `ChangesPreview`

A complete computation with bounded output; the session, its revision and history are unchanged.

```ts
interface ChangesPreview extends ChangesResult {
  impacts: CellImpact[]
  totalChanges: number
  truncated: boolean
}
```

### `ChangesResult`

What `applyChanges` answers: the edit's result and the clock it computed with.

```ts
interface ChangesResult extends EditResult { clock: Clock }
```

### `ChangingMethod`

The calls of a session that change its book, by name ([`SessionChange.method`](#sessionchange)).

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
  | 'setClock'
  | 'editStructure'
  | 'setStyles'
  | 'setLayout'
  | 'undo'
  | 'redo'
```

### `CheckCode`

Formula check codes. `unsupported`: an external function called with arguments that are not literals.
`unknown_name` covers tables too: an unknown table or column, a structured reference without a table outside one.

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

What [`Session.clearCells`](#session) clears: the content and hyperlinks (Excel's ClearContents), the look (ClearFormats), or both and the rest (Clear).

```ts
type ClearKind = 'contents' | 'formats' | 'all'
```

### `Clock`

The moment (ISO) and the random seed that TODAY, NOW, RAND and RANDBETWEEN are computed from.

```ts
interface Clock {
  now: string
  seed: number
}
```

### `Color`

A colour, `#RRGGBB`.

```ts
type Color = string
```

### `ColorScaleFormat`

Each number gets a colour between those of the points around it; other cells get none.

```ts
interface ColorScaleFormat extends ConditionalBase {
  type: 'colorScale'
  /** Two or three points, from the lowest. */
  points: { value: ConditionalValue; color: Color }[]
}
```

### `ColumnInput`

Columns of a sheet, as xlsx `<col>`: the width, and the style and format of their cells that have none.

```ts
interface ColumnInput {
  /**
   * In characters of the default font's digit with the cell's padding, as xlsx stores it, 0 to 255. Without it a
   * column has Excel's default width (8 digits and the padding; Excel for Windows shows it as 8.43).
   */
  width?: number
  /**
   * The width is one Excel fitted, by AutoFit or as a number typed into the column widened it, not one set by hand (xlsx
   * `bestFit`): a number typed into the column widens it again (`entryWidth`), as into a column of the default width.
   */
  bestFit?: boolean
  style?: number
  numberFormat?: NumberFormat
  /** Hidden (xlsx `hidden`); a column of width 0 is hidden too. SUBTOTAL and AGGREGATE count the cells of hidden columns all the same, as Excel does. */
  hidden?: boolean
  /** The outline level (Excel's Group, xlsx `outlineLevel`), 1 to 7; a column without one is not grouped. */
  outlineLevel?: number
  /** The mark of a collapsed group on its summary column (xlsx `collapsed`), as `RowInput.collapsed`. */
  collapsed?: boolean
}
```

### `ColumnsFit`

Excel's AutoFit of columns `C` or `B:D` (a double click on a column's border): each column with something to show gets
the width of its widest cell as Excel for Mac fits it ([`fitColumns`](#xlcalc-fitcolumns)), marked `bestFit` and shown if it was hidden; a
column with nothing to show keeps its width. The widths are computed from the values and layout the edits before it
leave (`docs/formats.md`, "Fitting columns").

```ts
interface ColumnsFit {
  sheet: string
  columns: string
  fit: true
}
```

### `ColumnWidth`

The width a column took ([`Session.setCellTexts`](#session)): its sheet, its letters (`B`) and the width, as [`ColumnInput.width`](#columninput).

```ts
interface ColumnWidth {
  sheet: string
  column: string
  width: number
}
```

### `ComputeOptions`

Options of [`compute`](#xlcalc-compute) and of a session: the clock, the way formulas run, limits and the values a file stored.

```ts
interface ComputeOptions {
  /**
   * The clock for TODAY, NOW, RAND and RANDBETWEEN. By default the current moment and a random seed;
   * pass the `clock` of a result to reproduce that result exactly.
   */
  clock?: Clock
  externals?: Externals
  /** `closures` by default. */
  execution?: Execution
  /** Each cell's shown text and colour (`CellResult.text`, `color`); true by default, false computes values alone. */
  display?: boolean
  /** None by default. */
  limits?: CalcLimits
  /**
   * The values the book's file stores (`readXlsx`'s `cached`), as Excel opens a file with a circular reference: a formula
   * the cycle stops (`CellInput.held`) without a `held` of its own holds the value stored for it, empty text where the
   * file stores none. Nothing else is read from them.
   */
  cached?: CachedValues
}
```

### `ComputeResult`

What [`compute`](#xlcalc-compute) gives: each sheet's cells with their values and shown text, the issues, the values' `hash`, `inputHash` and the clock it computed with.

```ts
interface ComputeResult {
  sheets: { name: string; cells: Record<string, CellResult> }[]
  /**
   * Every cell with an error and every formula a circular reference stops (`CellIssue`), sheet by sheet in the book's
   * order and row by row, then every name with a problem, in the order of the book's names.
   */
  diagnostics: Diagnostic[]
  /** Cells that depend on volatile functions (directly or through other cells). */
  volatile: CellRef[]
  /** The clock the book was computed with (for a live session, the latest one). */
  clock: Clock
  /** sha256 of the canonical text of the values (spilled cells too) together with the engine configuration version. */
  hash: string
  /** sha256 of the input (sheets, cells, names, and styles, columns, rows, merged areas, frozen panes, conditional formats, data validation and outline settings when there are any; order does not matter). */
  inputHash: string
  configVersion: string
  ms: number
}
```

### `ConditionalAppearance`

What a sheet's conditional formats give a cell ([`CellView.appearance`](#cellview)), by the values the cells hold now. A field the
cell shows otherwise than by its own style and value is set.

```ts
interface ConditionalAppearance {
  /** The styles of the rules that hold, merged in their order (the first rule's field wins); a colour scale's colour is `fill`. */
  style?: ConditionalStyle
  bar?: BarLook
  icon?: Icon
  /** A data bar or an icon set hides the value: the cell shows its bar or icon alone. */
  hideValue?: true
}
```

### `ConditionalBase`

What every kind of conditional format shares ([`ConditionalFormat`](#conditionalformat)): the cells it applies to and Excel's Stop If True.

```ts
interface ConditionalBase {
  /**
   * The cells it applies to: rectangles of the sheet (`B2:B20`, `D2`, whole columns `F:F`, whole rows `3:5`), as Excel's
   * "Applies to" and xlsx `sqref`. Its formulas are written for the top-left cell of the rectangle around them: a
   * reference without `$` moves from there to each cell, as a copied formula does.
   */
  ranges: string[]
  /** Excel's Stop If True: the rules after this one do not apply to a cell this one holds for. */
  stopIfTrue?: boolean
}
```

### `ConditionalFormat`

A conditional format: the cells it applies to, a condition and what a cell that meets it shows. The rules that hold
for a cell apply in the sheet's order: a field of a style an earlier rule sets wins, and `stopIfTrue` leaves out
the rules after one that holds. Its formulas are stored text, as a cell's: a leading `=` is dropped.

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

What a rule gives the cells it holds for, as an xlsx differential format (`dxf`): each field it sets replaces the cell's own.

```ts
interface ConditionalStyle {
  /** `bold: false` takes away a bold the cell has, and so on. */
  font?: Pick<FontStyle, 'bold' | 'italic' | 'underline' | 'strike' | 'color'>
  fill?: { color: Color }
  border?: CellStyle['border']
  /** The number format the cell shows its value by. */
  numberFormat?: NumberFormat
}
```

### `ConditionalValue`

A threshold of a colour scale, a data bar or an icon set, over the numbers of the cells the rule applies to. `min`
and `max`: the lowest and the highest number; `num`: the number `value` gives; `percent`: `value` percent of the
way from the lowest to the highest; `percentile`: the `value`-th percentile of the numbers; `formula`: what `value`
computes. A data bar's `autoMin` and `autoMax` are Excel 2010's automatic ends: 0, or the lowest number below 0
(the highest above it).

```ts
interface ConditionalValue {
  type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula' | 'autoMin' | 'autoMax'
  /** A formula as a cell stores it (`10`, `$B$1*2`), for `num`, `percent`, `percentile` and `formula`. */
  value?: string
}
```

### `CopyContext`

What [`copyFormula`](#xlcalc-copyformula) reads of a book: its sheets' names and their tables; a [`BookInput`](#bookinput) is one.

```ts
interface CopyContext {
  sheets: { name: string; tables?: TableInput[] }[]
}
```

### `CopyTextOptions`

Options of [`Session.copyText`](#session), the selection's plain-text copy.

```ts
interface CopyTextOptions {
  /** A terminal CRLF, as on Windows; false for Mac. True by default. */
  trailingNewline?: boolean
  /** Maximum visible cells to copy, 20,000 by default. A larger copy is refused, never truncated. */
  maxCells?: number
  /** Maximum characters of the encoded text, 16 Mi characters by default, checked before assembling its rows. */
  maxChars?: number
}
```

### `DataBarFormat`

Each number gets a bar across its cell, as long as it is far between `min` and `max`. The fields from `gradient` on
are Excel 2010's (xlsx `x14:dataBar`).

```ts
interface DataBarFormat extends ConditionalBase {
  type: 'dataBar'
  min: ConditionalValue
  max: ConditionalValue
  color: Color
  /** False: the cell shows the bar without its value. */
  showValue?: boolean
  /** The shortest and the longest bar in percent of the cell's width: 0 and 100 by default, 10 and 90 for an Excel 2007 bar. */
  minLength?: number
  maxLength?: number
  /** A gradient fill (true by default) or a solid one. */
  gradient?: boolean
  /** The bar's border; none without it. */
  border?: Color
  /** The bars' direction; by default the sheet's (left to right). */
  direction?: 'leftToRight' | 'rightToLeft'
  /** The fill of bars of negative numbers; the bar's colour without it. */
  negativeColor?: Color
  /** Their border; `border` without it. */
  negativeBorder?: Color
  /** Where the axis between negative and positive bars stands: by the numbers (by default), at the middle of the cell, or none (negative bars go the same way). */
  axis?: 'middle' | 'none'
  /** The axis's colour; black without it. */
  axisColor?: Color
  /** A bar of Excel 2007 (a file without Excel 2010's part of it): a gradient without a border, negative numbers as shorter bars, no axis; the fields from `gradient` on do not apply. */
  excel2007?: boolean
}
```

### `DataValidation`

A rule of data validation, as xlsx `<dataValidation>`. Its formulas are stored formula text without `=` (en-US, as a
cell's; a leading `=` is dropped), relative to the rule's first cell, the top-left cell of the first range of `range`:
a rule over `B2:B9` with `formula1: 'A2>0'` checks `A5>0` for B5. A literal list is a text constant of items separated by commas
(`'"Yes,No,Maybe"'`), as a file stores it in every locale; a list from cells is a reference (`'$D$2:$D$9'`,
`"'Lists'!$A$1:$A$20"`), a defined name or a formula giving one.

```ts
interface DataValidation {
  /** The cells: ranges separated by spaces (`B2:B20 D2`, `C:C`), as xlsx `sqref`; no other rule of the sheet meets them. */
  range: string
  type: ValidationType
  /** For whole, decimal, date, time and textLength: how the value compares; `between` by default. */
  operator?: ValidationOperator
  /** The value, the low bound, the list or the formula of a custom rule; every type but `any` has one. */
  formula1?: string
  /** The high bound of `between` and `notBetween`. */
  formula2?: string
  /**
   * A blank cell passes, and so does any value while a bound or a list refers to a blank cell (Excel's Ignore blank,
   * xlsx `allowBlank`); true by default, as Excel makes a rule.
   */
  allowBlank?: boolean
  /** A list shows its dropdown in the cell; true by default (false is xlsx `showDropDown="1"`). */
  dropdown?: boolean
  /** The input message Excel shows while the cell is selected: its title (Excel takes 32 characters) and text (255). */
  promptTitle?: string
  prompt?: string
  /** The input message is shown; true by default, as Excel makes a rule. */
  showPrompt?: boolean
  /** The alert for a value the rule does not take: stop (the default), warning or information, and its title (32) and text (225). */
  errorStyle?: 'stop' | 'warning' | 'information'
  errorTitle?: string
  error?: string
  /** The alert is shown; true by default, as Excel makes a rule. Without it Excel takes any value typed. */
  showError?: boolean
}
```

### `DateSystemOptions`

The book's date system, as [`BookInput.dateSystem`](#bookinput): 1900 without one.

```ts
interface DateSystemOptions {
  dateSystem?: 1900 | 1904
}
```

### `DependentCell`

A cell whose formula reads the ranges `getDependents` was asked about, with the references that read them.

```ts
interface DependentCell extends CellRef {
  /** Matching source references remain rectangles, never expanded into individual cells. */
  references: RangeRef[]
  referencesTruncated?: boolean
}
```

### `DependentsOptions`

Options of `getDependents`: those of a page and how many references to give for each cell.

```ts
interface DependentsOptions extends PageOptions {
  /** Maximum matching source references per dependent, 64 by default, at most 1000. */
  maxReferences?: number
}
```

### `DependentsResult`

What `getDependents` answers: a page of the cells whose formulas read the ranges.

```ts
interface DependentsResult extends QueryPage { dependents: DependentCell[] }
```

### `Diagnostic`

What a result reports about its book, by `kind`: a cell with an error and its reason (`cell`, a [`CellIssue`](#cellissue)), or a
name the book left out or keeps with an error (`name`, a [`NameIssue`](#nameissue)).

```ts
type Diagnostic = ({ kind: 'cell' } & CellIssue) | ({ kind: 'name' } & NameIssue)
```

### `DisplayOptions`

[`displayValue`](#xlcalc-displayvalue)'s options: where the value is shown.

```ts
interface DisplayOptions {
  /** The number format; General without it. */
  numberFormat?: NumberFormat
  /** A loaded locale's tag (`LOCALES`), or a locale object, loaded or not; en-US without one. */
  locale?: string | Locale
  /**
   * The column's width as `ColumnInput.width` (several columns' widths added for a merged area). Without it the width
   * is not limited: General shows all the digits it would in a wide column and no number turns into `#`.
   */
  width?: number
  /** The date system a date shows by, as `BookInput.dateSystem`; 1900 without it. */
  dateSystem?: 1900 | 1904
}
```

### `DuplicatesFormat`

Values that occur more than once among the cells (text in any case), or once.

```ts
interface DuplicatesFormat extends ConditionalBase {
  type: 'duplicateValues' | 'uniqueValues'
  style?: ConditionalStyle
}
```

### `EditOptions`

Options of a call that changes a session's book.

```ts
interface EditOptions {
  /**
   * The revision the caller's view of the book is of (`EditResult.revision`, `Session.revision`): when the session's is
   * another, the call changes nothing and throws `CalcError` with code `stale_revision`, so that a host does not apply
   * an edit over one it has not seen.
   */
  baseRevision?: number
}
```

### `EditResult`

What a call that changes a session's book answers: the session's revision after it, the cells it changed, and what
else a host showing the book reads again.

```ts
interface EditResult {
  /** The session's revision after the call: every call that changes the book adds one, `undo` and `redo` too. */
  revision: number
  /**
   * The cells whose value, shown text, formula, look, note or hyperlink changed, by `CellChange`; a cell another cell's value changed
   * is listed with its value alone. After a structure operation, the cells whose value or formula differ from those of
   * the cell that was at their old place (`moved`).
   */
  changes: CellChange[]
  /** The sheets whose layout changed (rows, columns, merged areas, tables, conditional formats, validation…): a host reads `getLayout()` again. */
  layout?: string[]
  /** The book's style table, when the call changed it. */
  styles?: CellStyle[]
  /** How addresses moved, in order, after a structure operation or its undo: a host moves its window by them. */
  moved?: AddressMove[]
  /**
   * Rectangles where cells may look otherwise without being in `changes`: the cells of the book's conditional formats,
   * after a call that changed a value or the rules (a rule's look reads values anywhere its formulas point).
   */
  repaint?: RangeRef[]
}
```

### `EndDirection`

The way Control and an arrow key go from a cell ([`Session.getEnd`](#session)).

```ts
type EndDirection = 'up' | 'down' | 'left' | 'right'
```

### `EntryCheck`

Whether a value meets a cell's rule ([`Session.validateEntry`](#session)); when it does not, the rule, whose alert Excel shows,
its style (`errorStyle`, stop by default) and its message: the rule's `error`, else the text of Excel's alert in the
book's locale.

```ts
type EntryCheck = { ok: true } | ValidationRefusal
```

### `EntryOptions`

[`parseEntry`](#xlcalc-parseentry)'s options: the book text is typed into and the cell it goes to.

```ts
interface EntryOptions {
  /**
   * The book's locale, a loaded locale's tag (`LOCALES`) or a locale object, loaded or not; en-US without one: numbers,
   * dates, booleans, errors and formulas are read as Excel reads them there.
   */
  locale?: string | Locale
  /** The book's date system, as `BookInput.dateSystem`; 1900 without it. */
  dateSystem?: 1900 | 1904
  /**
   * The cell's entry context: its effective `numberFormat`, `style`, `meta`, `note` and `link`; other fields are not read.
   * A `CellInput` or `CellView` of `getCell` or `getCells` carries only the cell's own format. To account for a row's or
   * column's format, pass the resolved cell of `getViewport`, or set `numberFormat` to the effective format yourself.
   * The new input keeps its style, meta, note and link, and its format where Excel keeps it.
   * The returned input stores that format on the cell; `setCellText` keeps an inherited format on its row or column.
   */
  cell?: { numberFormat?: NumberFormat; style?: number; meta?: JsonValue; note?: CellNote; link?: CellLink } | null
  /** The moment whose year completes a date typed without one (`15.3` in ru-RU, `3/15` in en-US), ISO as `Clock.now`; the current moment without it. */
  now?: string
  /**
   * The book's links to other workbooks (`BookInput.externalBooks`): a formula's reference to one, typed as Excel shows
   * it (`=[Book2.xlsx]Sheet1!A1`, `='C:\data\[Book2.xlsx]Sheet1'!A1`, `=Book2.xlsx!Rate`) or by number, is found by the
   * file name and numbered (`'[1]Sheet1'!A1`); one to a book they do not hold is refused (`external_ref`). Without them
   * such a formula is kept as typed.
   */
  externalBooks?: ExternalBook[]
}
```

### `EntryResult`

Text typed into a cell as Excel takes it ([`parseEntry`](#xlcalc-parseentry)): the cell's new input, `null` when nothing is left of the
cell; or a formula Excel does not take, with a check code (`syntax`, `too_long`, `too_deep`), a message in the
language of `opts.locale` (English for en-US) and the position in the typed text. A formula typed with brackets,
quotes or array braces left open at its end Excel closes on Return when it begins with a function, has no `)` and
leaves only functions' brackets open (`=SUM(1,2` is `=SUM(1,2)`); any other such formula it refuses and offers to
correct, when the formula closed reads with every function known: `correction` is that text, as typed with what was
left open closed (`=(1+2` offers `=(1+2)`). Pasted text is not closed.

```ts
type EntryResult = { ok: true; input: CellInput | null } | (Refusal<CheckCode> & { at: number; correction?: string })
```

### `EntryTextOptions`

[`entryText`](#xlcalc-entrytext)'s and [`needsApostrophe`](#xlcalc-needsapostrophe)'s options: the cell's number format (its row's or column's without its own), the
book's locale and date system, and for text the moment whose year completes a date typed without one, as [`parseEntry`](#xlcalc-parseentry)
reads it.

```ts
interface EntryTextOptions {
  numberFormat?: NumberFormat
  /** A loaded locale's tag (`LOCALES`) or a locale object; en-US without one. */
  locale?: string | Locale
  dateSystem?: 1900 | 1904
  /** ISO, as `Clock.now`; the current moment without it. */
  now?: string
}
```

### `EntryWidthOptions`

[`entryWidth`](#xlcalc-entrywidth)'s options: the format the value shows by (the cell's after the entry), the book's locale and date system,
and the column typed into, its `width` (the default without one) and `bestFit`.

```ts
interface EntryWidthOptions extends DisplayOptions {
  /** The column's width is one Excel fitted (`ColumnInput.bestFit`), which typing widens as it widens the default width. */
  bestFit?: boolean
}
```

### `ErrorCode`

`CALC`: Excel's #CALC!, an empty array (`FILTER` that keeps nothing) or a text over 32,767 characters (`CONCAT`).
`SPILL`: an array that cannot spill from its formula's cell (its range is not empty, meets a merged area or
leaves the sheet). `NULL`: Excel's #NULL!, ranges that do not meet under the intersection operator (`A1:A3 C1:C3`).
`ERROR`: a formula that does not compute (it does not parse, say).

```ts
type ErrorCode = 'DIV0' | 'NA' | 'NAME' | 'NUM' | 'REF' | 'VALUE' | 'NULL' | 'CYCLE' | 'SPILL' | 'CALC' | 'ERROR'
```

### `Execution`

How formulas run; values, issues and hashes are the same either way. `closures` (the default) builds each
formula as a tree of JavaScript closures and generates no code, so it works under a Content Security
Policy without 'unsafe-eval'. `eval` generates JavaScript text and compiles it with `new Function`, which
such a policy blocks (computing then throws an `EvalError`).

```ts
type Execution = 'closures' | 'eval'
```

### `ExplainNode`

A node of an explanation ([`Explanation.root`](#explanation)): a cell or a range, its formula and value, and the nodes it reads.

```ts
interface ExplainNode {
  /** The local node's sheet and A1 cell or range for navigation; absent for another workbook. */
  location?: RangeRef
  /** `'Sheet'!B3` or `'Sheet'!D:D`. */
  ref: string
  value?: CellValue
  formula?: string
  /** A cell an array spilled into: its formula's cell (`'Sheet'!B2`), which is the node's child. */
  spillOf?: string
  /** For a range: the number of non-empty cells. */
  cells?: number
  /**
   * A cell with an error, or a formula a circular reference stops: the cell the error began in (`'Sheet'!A2`, the node's
   * own `ref` when it began there), as `CellIssueDetail.origin`. The way there is in the tree past `depth` too, and
   * through a range by the cell it came from.
   */
  origin?: string
  children: ExplainNode[]
}
```

### `ExplainOptions`

Options of `explain`: how deep the tree goes and how many nodes it holds.

```ts
interface ExplainOptions {
  /** Depth over precedents, 3 by default. */
  depth?: number
  /** Node limit, 200 by default. */
  maxNodes?: number
  /** Expand ranges into their non-empty cells (by default a range is a single node). */
  expand?: boolean
  /** How many cells past `depth` the way from an error to the cell it began in goes on, 64 by default. */
  errorDepth?: number
}
```

### `Explanation`

What `explain` answers: why a cell holds its value, as a tree of what it reads and as text.

```ts
interface Explanation {
  root: ExplainNode
  nodes: number
  truncated: boolean
  /**
   * The tree as indented text (for logs, an LLM agent or a person) in the book's locale: in ru-RU its words, numbers,
   * formulas, booleans and errors are Russian (`План!B2 = 480  ← ОКРУГЛ(B1*1,2;0)`), and a sheet is in quotes only where
   * Excel puts them; the nodes keep the stored text. An error is written as the book's cells show it (`#DIV/0!`), with the
   * cell it began in (`(error from S!A2)`).
   */
  text: string
}
```

### `ExpressionFormat`

A formula that holds, as IF reads its condition: TRUE, a number other than 0, the text TRUE.

```ts
interface ExpressionFormat extends ConditionalBase {
  type: 'expression'
  formula: string
  style?: ConditionalStyle
}
```

### `ExternalBook`

Another workbook as a book keeps it beside the formulas that read it (xlsx `externalLinkN.xml`): where it is, and
the values Excel stored of it when it last read it.

```ts
interface ExternalBook {
  /** Where the link points, as the file holds it: a file name, a relative or an absolute path, a `file:` or web address. */
  target: string
  /** Its sheets in the order the file lists them. */
  sheets: ExternalSheet[]
  /** Its defined names the book's formulas use: of the book, or of one of its sheets. */
  names?: ExternalName[]
}
```

### `ExternalName`

A defined name of another workbook.

```ts
interface ExternalName {
  name: string
  /** What it refers to, as the file holds it without `=`: a reference into its book (`Sheet1!$B$1`), `#REF!`. */
  ref: string
  /** The index in `sheets` of the sheet whose own name it is; a name of the book has none. */
  sheet?: number
}
```

### `Externals`

External functions (host functions evaluated outside the book): their values are resolved in
advance, in one request, and passed to the computation. The key of a call is its canonical text
(`BALANCE("2025-01")`); arguments are literals only.

```ts
interface Externals {
  functions: string[]
  values?: Record<string, CellValue>
}
```

### `ExternalSheet`

A sheet of another workbook and the values stored of it.

```ts
interface ExternalSheet {
  name: string
  /**
   * The stored values by A1 address, null for a cell stored empty. A cell they do not hold is blank, or `#REF!` with
   * `refreshError`. Without `cells` nothing of the sheet is stored, and every cell of it is `#REF!`.
   */
  cells?: Record<string, ExternalValue | null>
  /**
   * Excel's mark that the last update of the sheet's values failed (`refreshError`): a cell they do not hold is `#REF!`
   * instead of blank. Excel sets it on every sheet of a closed book when a formula referring to that book is entered,
   * and so does a session.
   */
  refreshError?: true
}
```

### `ExternalValue`

A value stored of another workbook's cell.

```ts
type ExternalValue = number | string | boolean | { error: ErrorCode }
```

### `FillTextOptions`

[`Session.fillCellText`](#session)'s options.

```ts
interface FillTextOptions extends CellTextOptions {
  /** The selection's active cell, which the text is typed into: a cell of the range (`B2`); its top-left cell by default. */
  active?: string
}
```

### `FillTextPreview`

What [`Session.fillCellText`](#session) would do ([`Session.previewFillCellText`](#session)): each cell's input and value, the widths and the alert, with nothing changed.

```ts
type FillTextPreview = { ok: true; cells: PreviewCell[]; widths?: ColumnWidth[]; validation?: ValidationAlert } | CellTextsRefusal
```

### `FillTextResult`

A text entered into every cell of a range at once ([`Session.fillCellText`](#session), Excel's Control+Return): the edit's result,
with the widths the columns took (`widths`) and the alert of the active cell's rule of data validation
(`validation`); or why it was not taken, with the cell, and nothing written.

```ts
type FillTextResult = (EditResult & { ok: true; widths?: ColumnWidth[]; validation?: ValidationAlert }) | CellTextsRefusal
```

### `FillType`

What a fill takes from its source ([`StructureOp`](#structureop) `fillRange`), as Excel's `XlAutoFillType`. `default` is the drag of
the fill handle (`xlFillDefault`). A line's source is cut into runs of cells of one kind, empty cells skipped: numbers
go on by their step or trend; dates by days, by their step, or by months when they share a day of the month or are
months' ends; times by hours or their step; text with a number at the start (and a space) or at the end by the number,
quarters and en-US ordinals too; names of days and months through their list; dates, times and text without one step
are copied, and so are text, booleans and formulas. A single number goes on by 1 unless the source holds plain numbers
alone. `series` (`xlFillSeries`, the drag with Ctrl over a number) makes a single number go on by 1 always. `copy`
(`xlFillCopy`) copies every cell.

```ts
type FillType = 'default' | 'series' | 'copy'
```

### `FindCellsOptions`

Options of `findCells`: what is searched and how, beside those of `readCells`.

```ts
interface FindCellsOptions extends ReadCellsOptions {
  /**
   * What is searched: the text a cell shows (`text`, the default; Excel's Look in Values), the text it opens with for
   * editing (`edit`: a formula in the book's locale, a number by its format, a text as it is, without an apostrophe;
   * Excel's Look in Formulas), its stored formula (`formula`) or its value (`value`). No regular expressions.
   */
  in?: 'text' | 'edit' | 'formula' | 'value'
  /** Excel's Match case; any case by default (`straße` is `STRAẞE`, `i` is not `İ`). */
  caseSensitive?: boolean
  /** The query as Excel's Find reads it: `*` any text, `?` one character (two for one outside the BMP), `~` before a character takes it as it is. Literal without it. */
  wildcards?: boolean
  /** Excel's Match entire cell contents: the whole text matches the query, not a part of it. */
  wholeCell?: boolean
  /** Excel's Search By Columns: each range's cells column by column; row by row without it. */
  byColumns?: boolean
  /**
   * Cells of hidden rows and columns are skipped as Excel's Find skips them: always where it reads what cells show
   * (`text`, Look in Values), and in every kind of search on a sheet with a filter, where the cells come by the blocks
   * the hidden lines leave, each block in turn (`docs/live-books.md`). Without it every cell of the ranges is searched.
   */
  skipHidden?: boolean
  /**
   * The cell Find begins after, as `Range.Find`'s `After` (a person's Find Next from the active cell): the cells come in
   * Find's order from the first after it, round to the ranges' first and on to it, itself last; a cell of the ranges.
   */
  after?: CellRef
}
```

### `FontStyle`

A font; what a style leaves out is the book's default font.

```ts
interface FontStyle {
  name?: string
  /** Points, 1 to 409. */
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: 'single' | 'double'
  strike?: boolean
  color?: Color
}
```

### `FormatCategory`

A category of number formats as the Number Format box of Excel's Home tab shows it for the active cell
([`numberFormatCategory`](#xlcalc-numberformatcategory)): the category of Excel's Format Cells window that makes the format, or `custom` for a format
none of them makes.

```ts
type FormatCategory = 'general' | 'number' | 'currency' | 'accounting' | 'date' | 'time' | 'percentage' | 'fraction' | 'scientific' | 'text' | 'special' | 'custom'
```

### `FormatChoice`

An item of the Number Format list in a locale: its key, its name in the list (`Short Date`), the format it applies (a
built-in format by number, or a code as Excel writes it to a file) and that format's code in the locale, in en-US
syntax as a cell's `numberFormat` takes it.

```ts
interface FormatChoice {
  key: FormatChoiceKey
  name: string
  format: NumberFormat
  code: string
}
```

### `FormatChoiceKey`

An item of the Number Format box's list on the Home tab, in its order ([`numberFormatChoices`](#xlcalc-numberformatchoices)).

```ts
type FormatChoiceKey = 'general' | 'number' | 'currency' | 'accounting' | 'shortDate' | 'longDate' | 'time' | 'percentage' | 'fraction' | 'scientific' | 'text'
```

### `FormulaReference`

A reference found in formula text: to cells or lines of a sheet, or to a table.

```ts
type FormulaReference = GridReference | TableReference
```

### `FrozenPanes`

A sheet's frozen panes ([`SheetInput.frozen`](#sheetinput), Excel's Freeze Panes): the rows above the split and the columns left of it.

```ts
interface FrozenPanes {
  /** 1 to 1,048,575. */
  rows?: number
  /** 1 to 16,383. */
  columns?: number
}
```

### `FunctionArgument`

An argument of a function's signature: its name in the locale and whether a call can leave it out.

```ts
interface FunctionArgument {
  name: string
  optional: boolean
}
```

### `FunctionCategory`

A category of functions, as Excel's Insert Function window lists them; `compatibility` holds the old names Excel keeps
beside their newer versions (`NORMDIST` beside `NORM.DIST`).

```ts
type FunctionCategory = 'financial' | 'date' | 'math' | 'statistical' | 'lookup' | 'database' | 'text' | 'logical' | 'information' | 'engineering' | 'compatibility'
```

### `FunctionSignature`

A function's signature for a tip of its arguments while a formula is typed ([`functionSignature`](#xlcalc-functionsignature)), as Excel's tip lists
them: `SUM(number1, [number2], ...)`, in ru-RU `СУММ(число1;[число2];...)`.

```ts
interface FunctionSignature {
  /** The function's name in the locale (`СУММ`). */
  name: string
  /** Its English name, as stored formulas, `CallAt.fn` and `functionDescription` take it (`SUM`). */
  fn: string
  /** The arguments in the tip's order, those that repeat listed twice (`number1`, `number2`), the second time optional. */
  args: FunctionArgument[]
  /**
   * How many of the last arguments repeat past the list (the tip's `...`): 1 for SUM, 2 for SUMIFS, 0 when none do. The
   * argument at index `i` past the list is named as the one `repeat` before it, with its number counted on.
   */
  repeat: number
  /** The tip's text, with the locale's list separator. */
  text: string
  /** The function's category, as Excel's Insert Function window lists it (`FUNCTION_CATEGORIES`). */
  category: FunctionCategory
  /** The category's name in the locale (`Math & Trig`, `Математические`). */
  categoryName: string
}
```

### `GridReference`

A reference to cells, whole columns or whole rows, with its rectangle ([`CellRect`](#cellrect)) and its sheet as written (unquoted):
`sheet` alone for one sheet, `sheet` and `sheetLast` for a reference across sheets (`Jan:Mar!A1`), neither for the
formula's own sheet. `spill`: a spill reference (`E2#`), whose rectangle is the formula's cell.

```ts
interface GridReference extends ReferenceSpan, CellRect {
  kind: 'cells' | 'columns' | 'rows'
  sheet?: string
  sheetLast?: string
  spill?: true
}
```

### `Icon`

An icon of a set: the set and its number in it, from 0, in Excel's order (xlsx `iconId`: for `3Arrows` 0 is the red arrow down, 2 the green arrow up).

```ts
interface Icon {
  set: IconSetName
  index: number
}
```

### `IconSetFormat`

Each number gets an icon of the set by the thresholds it reaches.

```ts
interface IconSetFormat extends ConditionalBase {
  type: 'iconSet'
  icons: IconSetName
  /**
   * One threshold per icon of the set, from the lowest; a number gets the icon of the last threshold it reaches (`gte`:
   * at or above it, true by default; false: above it), the first threshold always. `icon` puts another icon there,
   * `null` none (Excel 2010's custom icons).
   */
  points: { value: ConditionalValue; gte?: boolean; icon?: Icon | null }[]
  /** The icons in the reverse order: the last for the lowest numbers. */
  reverse?: boolean
  /** False: the cell shows the icon without its value. */
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
  /** Milliseconds of computing between two yields to the event loop; 10 by default. */
  sliceMs?: number
  /** The soft deadline of every call that gives none. */
  timeoutMs?: number
  /** The limits of every computation; a computation's own `limits` take their place field by field, the others stay. */
  limits?: CalcLimits
}
```

### `JsonValue`

Data a host keeps on a cell ([`CellInput.meta`](#cellinput)), as JSON.

```ts
type JsonValue = Scalar | JsonValue[] | { [key: string]: JsonValue }
```

### `LayoutEdit`

A change of a sheet's layout.

```ts
type LayoutEdit =
  /** Columns `C` or `B:D`; the patch applies to each. Setting or removing width clears bestFit unless explicitly given. */
  | { sheet: string; columns: string; set: LinePatch<ColumnInput> }
  | ColumnsFit
  | RowsFit
  /** Rows `5` or `5:9`; the patch applies to each. Setting or removing height clears bestFit unless explicitly given. */
  | { sheet: string; rows: string; set: LinePatch<RowInput> }
  /** All the merged areas of the sheet. */
  | { sheet: string; merges: string[] }
  /** The sheet's filter mode (`SheetInput.filtered`). */
  | { sheet: string; filtered: boolean }
  /** The sheet's frozen panes (`SheetInput.frozen`); null unfreezes them. */
  | { sheet: string; frozen: FrozenPanes | null }
  /** All the conditional formats of the sheet (`SheetInput.conditionalFormats`), in priority order. */
  | { sheet: string; conditionalFormats: ConditionalFormat[] }
  /** All the rules of data validation of the sheet (`SheetInput.validations`). */
  | { sheet: string; validations: DataValidation[] }
  /** Where the sheet's summary rows and columns are (`SheetInput.outline`); null for Excel's default, below and right. */
  | { sheet: string; outline: OutlineSettings | null }
```

### `LegacyContext`

What [`fromLegacy`](#xlcalcxlsx-fromlegacy) knows of the book: its names, and the sheet of the formula (whose local names it sees).

```ts
interface LegacyContext {
  names?: NameInput[]
  sheet?: string
}
```

### `LinePatch`

Fields to change on columns or rows: a value sets the field, null removes it, a field left out stays.

```ts
type LinePatch<T> = { [K in keyof T]?: T[K] | null }
```

### `LineSpan`

The lines a key of [`SheetInput.columns`](#sheetinput) or `rows` names (`C`, `B:D`, `5`, `5:9`), numbered from 1.

```ts
interface LineSpan {
  first: number
  last: number
}
```

### `LiveCell`

A cell of a [`LiveSheet`](#livesheet), by its address: after rows are inserted above it, it is the cell now at that address. Each
property reads the session as it is now; an assignment edits it at once, or as part of a `batch`, and throws what the
session's call throws, the book unchanged.

```ts
interface LiveCell {
  readonly sheet: LiveSheet
  /** The cell's address, `B12`. */
  readonly address: string
  /** The computed value, null for an empty cell. */
  get value(): CellValue
  /** As `setCells` with a `value`: the content is replaced, the formula goes, the look stays. Null clears the content and the hyperlink, as `clearCells` with `contents`. */
  set value(value: Scalar | { date: string })
  /** The text the cell shows (`CellResult.text`); undefined for an empty cell and in a session without `display`. */
  get text(): string | undefined
  /** The formula without `=`, as `CellResult.formula`; undefined for a cell without one. */
  get formula(): string | undefined
  /** As `setCells` with a `formula`; null removes it. */
  set formula(formula: string | null)
  /** The cell's own number format, as `CellView.numberFormat`. */
  get numberFormat(): NumberFormat | undefined
  /** As `patchCells`: the value and formula stay; null removes the cell's own format. */
  set numberFormat(format: NumberFormat | null)
  /** The cell's own style, an index of the book's `styles`, as `CellView.style`. */
  get style(): number | undefined
  /** As `patchCells`: a style by its number, or by its fields over the style the cell shows; null removes the cell's own. */
  set style(style: number | StylePatch | null)
  /** Text typed into the cell, as `Session.setCellText` takes it and answers. */
  type(text: string, opts?: CellTextOptions): CellTextResult
  /**
   * Calls `listener` after each change that touched the cell: its value, shown text, formula or look changed
   * (`EditResult.changes`, the cells a recomputation changed too), or a structure operation moved other cells to its
   * address (`moved`). Answers the function that unsubscribes; after the sheet is deleted it is called no more.
   */
  watch(listener: (change: SessionChange) => void): () => void
}
```

### `LiveSheet`

A sheet of a [`LiveWorkbook`](#liveworkbook). After the sheet is deleted, every use of the object throws [`InputError`](#xlcalc-inputerror).

```ts
interface LiveSheet {
  /** The sheet's name now: a rename changes it. */
  readonly name: string
  /**
   * The sheet's cells by address in any case, without `$` (`cells.B12`, `cells.b12`), each made when it is read and the
   * same object while it is held; any other key (a range, `then`, `toJSON`, a symbol) is undefined.
   */
  readonly cells: { readonly [address: string]: LiveCell }
  /** The cell at `address`, as `cells` gives it; `InputError` for anything but the address of a cell. */
  cell(address: string): LiveCell
  /**
   * Calls `listener` after each change that touched the sheet: its cells, its layout, a structure operation on it, its
   * rename or move, the book's style table. Answers the function that unsubscribes; after the sheet is deleted it is
   * called no more.
   */
  watch(listener: (change: SessionChange) => void): () => void
}
```

### `LiveWorkbook`

Objects of a synchronous session's sheets and cells ([`liveWorkbook`](#xlcalc-liveworkbook)): a cell reads what the session computed, an
assignment to it is the session's edit (recomputed, one step of the undo), and a listener learns of the changes of a
cell, a sheet or the book, whoever makes them. One per session: [`liveWorkbook`](#xlcalc-liveworkbook) gives the same object for it again.

```ts
interface LiveWorkbook {
  /** The session the objects read and edit. */
  readonly session: Session
  /**
   * The sheet by its name in any case: the same object while the sheet lives, renamed or moved. `InputError` for a name
   * the book has no sheet by; a sheet deleted and brought back by `undo` is a new object.
   */
  sheet(name: string): LiveSheet
  /**
   * Runs `fn` and applies the assignments it makes to cells as one `applyChanges`: one recomputation, one revision, one
   * step of the undo, one change told. Until `fn` returns, the cells read as they were before it. A batch inside another
   * is part of the outer one; when `fn` throws, none of its assignments is applied and the error goes on. Text typed
   * (`LiveCell.type`) cannot be part of a batch, nor an `fn` that answers a promise.
   */
  batch(fn: () => void): void
  /** Calls `listener` after each change of the book, as `Session.subscribe` does; answers the function that unsubscribes. */
  watch(listener: (change: SessionChange) => void): () => void
}
```

### `Locale`

A locale ([`ruRU`](#xlcalclocalesru-ru-ruru) of `xlcalc/locales/ru-RU`): its formula text and, for a book that computes in it, its rules of computation and messages.

```ts
interface Locale {
  /** BCP 47 tag: `en-US`, `ru-RU`. A locale without messages of its own takes those of the loaded locale of its tag (English for another tag). */
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
   * The specifiers of structured references by their invariant text (`'#Data': '#Данные'`); without them, or for one
   * not listed, the English one. Between a reference's items the list separator stands (`[#Данные];[Qty]` in ru-RU).
   */
  readonly tableItems?: Readonly<Partial<Record<TableItem, string>>>
  /**
   * The words of the arguments' names in the tips of functions (`functionSignature`) by their English word (`number:
   * 'число'`, so `number1` is `число1`); a word not listed stays English. The package's locales leave them, and
   * `functionText`, to modules of their own (`xlcalc/function-text/ru-RU`), which a host joins to the locale.
   */
  readonly argumentWords?: Readonly<Record<string, string>>
  /**
   * The words of Excel's Insert Function window: the categories' names (`functionSignature`) and what each function does
   * (`functionDescription`), by its English name; what is not listed stays English.
   */
  readonly functionText?: { readonly categories: Readonly<Partial<Record<FunctionCategory, string>>>; readonly descriptions: Readonly<Record<string, string>> }
  /**
   * The rules of computation as data: number and date text, names, `TEXT`'s codes, the text order, the code page. A
   * locale a book computes in has them, but en-US, whose rules are the engine's own code; a locale without them reads
   * and shows values as en-US does.
   */
  readonly regional?: RegionalData
  /**
   * Messages for people in the locale's language (`messages.ts`): why a formula does not read, why a structure operation
   * is refused, the words of `explain`, what an xlsx file lost. A message the locale lacks is English. They are functions, the one part of a
   * locale a structured clone does not carry: a pool's thread imports the package's module of the tag itself.
   */
  readonly messages?: Partial<Messages>
}
```

### `LookRange`

A rectangle of a sheet's looks ([`SheetInput.looks`](#sheetinput)): a cell (`B2`), a range (`B2:D9`), whole columns (`B:D`) or whole
rows (`2:9`), and the style and the number format its cells without an entry show; neither field, and its cells show
none of their own. The books the engine gives back hold the looks in one form for one set of them: rows in blocks, the
runs of a group of rows as rectangles across the group, then the rows of the group that differ, its rows without looks
as rows of none.

```ts
interface LookRange {
  range: string
  /** A style of the book's `styles`, by index. */
  style?: number
  numberFormat?: NumberFormat
}
```

### `NameCheck`

What [`validateName`](#xlcalc-validatename) answers: the name is good, or why not.

```ts
type NameCheck = { ok: true } | Refusal<NameRefusal>
```

### `NameInput`

A defined name of a book ([`BookInput.names`](#bookinput)): the name, what it refers to and its scope.

```ts
interface NameInput {
  /**
   * Excel's rules for names, a table's too: a letter of any script, `_` or `\` first, then letters, digits, `_`, `.`,
   * `\` and `?`, at most 255 characters; not a cell address (`Q1`) or R1C1 (`R`, `C`, `R2C3`), not TRUE or FALSE, and
   * not the book's locale's words for them either, in any case (ИСТИНА and ЛОЖЬ in ru-RU, which Excel ru-RU refuses;
   * `Ставка`, `НДС_2025`, `АБ12` and `Став?ка` are names). Compared in any case: `ставка` is the name `Ставка`.
   */
  name: string
  /**
   * What the name refers to: a formula without `=` (a leading one is dropped), as Excel's "Refers to" and xlsx
   * `<definedName>` hold it. A reference with a sheet (`'Data'!$D$2:$D$3401`, `'Data'!$D:$D`), a constant (`0.2`,
   * `"text"`, `{1,2,3}`), a formula (`'Data'!$A$1*2`, `SUM('Data'!$A:$A)`, `INDEX('Data'!$A$1:$C$9,2,0)`, which is a
   * reference), a union of references on one sheet (`'Data'!$A$1:$A$3,'Data'!$C$1:$C$3`). Every reference has its sheet.
   * A reference without `$` is relative to A1, as a file stores it: used in a cell it moves by that cell's offset from
   * A1, wrapping round the sheet's edges (`'Data'!B1` used in C5 is `'Data'!D5`, `'Data'!A1048576` is the row above).
   * Names in it are the global ones (a sheet's own as `'Sheet'!name`), as a file stores a name's formula. The formula is
   * computed where the name is used, as part of that formula.
   */
  ref: string
  /** The sheet where the name is visible; a local name shadows a global one. Without it the name is global. */
  scope?: string
}
```

### `NameIssue`

Why a defined name gives no value: the name, its scope (none for a global name) and the reason.

```ts
interface NameIssue {
  name: string
  scope?: string
  reason: string
  /**
   * The book has no such name: one Excel would not take, a duplicate, an unknown scope sheet, a formula that does not
   * parse or a reference without its sheet; a formula using the word finds no name. False: the name is in the book and
   * gives an error where it is used (an unknown sheet, table or function in its formula).
   */
  dropped: boolean
}
```

### `NameOptions`

The book whose rules [`validateName`](#xlcalc-validatename) checks a name by.

```ts
interface NameOptions {
  /** The book's locale (`BookInput.locale`): a loaded locale's tag, or a locale object; en-US without one. */
  locale?: string | Locale
}
```

### `NameRefusal`

Why Excel would not take a name ([`validateName`](#xlcalc-validatename)). `chars`: it does not start with a letter of any script, `_` or `\`,
holds a character other than letters, digits, `_`, `.`, `\` and `?`, or is empty or longer than 255 characters.
`reference`: it reads as a cell address or an R1C1 reference (`Q1`, `R`, `R2C3`). `boolean`: TRUE or FALSE, or the
book's locale's word for one, in any case (`ИСТИНА` in ru-RU).

```ts
type NameRefusal = 'chars' | 'reference' | 'boolean'
```

### `NamesContext`

The parts of a book [`validateNames`](#xlcalc-validatenames) checks names against.

```ts
interface NamesContext {
  /** Sheets of the book: a name's scope must be one of them, and its formula's references to others give `#REF!`. */
  sheets: string[]
  /** Tables of the book with their columns, for the structured references of names' formulas. */
  tables?: { name: string; columns: string[] }[]
  /** The book's links to other workbooks, for the names' references to them (`ValidateContext.externalBooks`). */
  externalBooks?: ExternalBook[]
  externalFunctions?: string[]
  /** The book's locale (`BookInput.locale`): a loaded locale's tag, or a locale object; en-US without one. */
  locale?: string | Locale
}
```

### `NumberFormat`

A number format: a format code as an xlsx file stores it, or the number of one of Excel's built-in formats, 0 to
49, as a file stores a cell's format with no code of its own (`numFmtId`). Excel shows some built-in formats by the
locale: 14 is the locale's short date (`m/d/yyyy` in en-US, `dd.mm.yyyy` in ru-RU), 5 to 8 and 37 to 44 its money
and number formats, 15 to 22 its other dates and times. A code is shown as written whatever the locale.

```ts
type NumberFormat = string | number
```

### `NumberFormatCategory`

What the Number Format box shows for a format: its category and the category's name in the locale (`Number`, `Числовой`).

```ts
interface NumberFormatCategory {
  category: FormatCategory
  name: string
}
```

### `OutlineGroup`

A group of rows or columns as Excel draws it ([`outlineGroups`](#xlcalc-outlinegroups)): the longest run of lines `first` to `last` (columns by
number) at outline level `level` or deeper, so a group of level 2 lies inside one of level 1.

```ts
interface OutlineGroup extends LineSpan {
  level: number
  /**
   * The summary line, which holds the group's button: the line past it, below or right (before it with summaries above or
   * left, `OutlineSettings`); null where that is past the sheet's edge, where Excel draws no button. Groups that end
   * together share one.
   */
  summary: number | null
  /** Collapsed, the button showing `+`: its summary line marked as Excel keeps the mark (`RowInput.collapsed`); without a summary line, all its lines hidden. */
  collapsed: boolean
}
```

### `OutlineSettings`

The summary of a group of rows is the row below it (`summaryBelow`, true by default) or above; of columns, right of it or left.

```ts
interface OutlineSettings {
  summaryBelow?: boolean
  summaryRight?: boolean
}
```

### `PageOptions`

Options of a query that answers in pages (`readCells`, `findCells`, `getDependents`): how many results a page holds and where the last page stopped.

```ts
interface PageOptions extends EditOptions {
  /** Maximum results returned, 1000 by default, at most 10000. */
  maxCells?: number
  /** Opaque continuation of the same query and revision; an edited session refuses it. */
  cursor?: string
  /** Count all matching cells, even after the page is full; false by default. */
  count?: boolean
}
```

### `PasteTextPreview`

What [`Session.pasteText`](#session) would do ([`Session.previewPasteText`](#session)): the range and each of its cells' input and value, with nothing changed.

```ts
type PasteTextPreview = { ok: true; range: string; cells: PreviewCell[] } | PasteTextRefusal
```

### `PasteTextRefusal`

Why [`Session.pasteText`](#session) pasted nothing, with the cell: part of an array formula, a formula into a table's header cell
(as `setCellText` refuses them), or a rectangle that runs past the sheet's last row or column (`off_sheet`).

```ts
type PasteTextRefusal = Refusal<'array' | 'header' | 'off_sheet'> & CellRef
```

### `PasteTextResult`

Text pasted from the clipboard ([`Session.pasteText`](#session)): the edit's result with the range the clipboard's rectangle filled (`B1:D2`); or why nothing was pasted.

```ts
type PasteTextResult = (EditResult & { ok: true; range: string }) | PasteTextRefusal
```

### `PatchFields`

The fields [`CellPatch`](#cellpatch) sets; on whole rows and columns, `meta`, `note` and `link` can only be removed.

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

What an arrow does at a place of formula text typed over a cell (Excel's Enter mode), as [`refAt`](#xlcalc-refat) says it: `insert`, a
reference to the cell the arrow points at goes in at `at` (Point mode), after an `=` of its own with `equals` (at the
start of the text, where Excel begins a formula anew: `=A1` becomes `=D3=A1`); `commit`, the text is entered into the cell and
the arrow moves the selection, as after a value; `none`, nothing happens (the text cannot take a reference there and is
not a formula Excel enters: inside a string, after a character the locale's formulas do not have).

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

Options of `previewChanges`: those of `applyChanges` and how many changed cells to give.

```ts
interface PreviewChangesOptions extends ChangesOptions {
  /** Maximum changed cells returned, 1000 by default, at most 10000. The computation itself remains complete. */
  maxChanges?: number
}
```

### `QueryPage`

A page of a query's answer: the revision and clock it read, whether more is left and where to go on.

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

A range's rectangle and what its text names: cells (`B2:D9`, `B2`), whole columns (`B:D`) or whole rows (`2:9`).

```ts
interface RangeRect extends CellRect {
  kind: 'cells' | 'columns' | 'rows'
}
```

### `RangeRef`

A rectangle on a sheet: `B2`, `B2:D9`, whole columns `B:D` or whole rows `2:9` (no `$`, no sheet).

```ts
interface RangeRef {
  sheet: string
  range: string
}
```

### `RangeSummary`

What `summarizeRanges` answers, as Excel's status bar shows a selection: how many cells hold something, how many numbers and errors, their sum, the least and the greatest number.

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

A cell as `readCells` and `findCells` give it: its address and the fields asked for.

```ts
interface ReadCell extends CellRef, Partial<CellView> {
  issue?: CellIssueDetail | null
  /** With the field `editText`: the text the cell opens with for editing, as `ViewportCell.editText`; what `findCells` reads in `edit`, but for the apostrophe before a text that reads as something else, which it leaves out. */
  editText?: string
}
```

### `ReadCellsOptions`

Options of `readCells`: the fields to read and which rows and columns.

```ts
interface ReadCellsOptions extends PageOptions {
  /** value, text and formula by default. Address and sheet are always returned. */
  fields?: CellReadField[]
  /**
   * Only the cells of these rows (the rows a host's filter shows, say), as `ViewportOptions.rows`: their numbers on the
   * sheet, from 1, in ascending order, each once; the cells of each range in them, `total` counting those alone. A row
   * outside a range leaves that range nothing of it.
   */
  rows?: number[]
}
```

### `ReadCellsResult`

What `readCells` and `findCells` answer: a page of cells.

```ts
interface ReadCellsResult extends QueryPage { cells: ReadCell[] }
```

### `ReadXlsxOptions`

The options of [`readXlsx`](#xlcalcxlsx-readxlsx).

```ts
interface ReadXlsxOptions {
  /**
   * The language of the issues' messages: a loaded locale's tag (`'ru-RU'` once `registerLocale(ruRU)` ran) or a locale
   * object (`ruRU`), en-US by default. Codes and the other fields of an issue are the same in every language. A file
   * does not say what language its reader reads, so the host passes the locale the person chose.
   */
  locale?: string | Locale
}
```

### `Refusal`

A check or an edit that did not pass: a code a program reads, a message for people (in the book's locale where the
call has one, English otherwise) and, for typed text, the position in it. Every check answers `{ ok: true, … }` or
a [`Refusal`](#refusal) with its own codes.

```ts
interface Refusal<Code extends string = string> {
  ok: false
  code: Code
  message: string
  at?: number
}
```

### `ReplaceOptions`

[`Session.replaceInCells`](#session)' options: Find's, as `findCells` takes them, but what it reads, which is the text a cell opens with (Look in Formulas), and the order.

```ts
interface ReplaceOptions extends EditOptions, Pick<FindCellsOptions, 'caseSensitive' | 'wildcards' | 'wholeCell' | 'skipHidden' | 'rows'> {}
```

### `ReplaceResult`

What [`Session.replaceInCells`](#session) did: the edit's result, and how many cells' texts it wrote (`replaced`, an array of the
old kind as one); the cells whose new text Excel would not take, left as they were, with why (`refused`).

```ts
interface ReplaceResult extends EditResult {
  replaced: number
  refused?: (Refusal<CellTextRefusal> & CellRef)[]
}
```

### `RowInput`

Rows of a sheet, as xlsx `<row>`: the height, and the style and format of their cells that have none.

```ts
interface RowInput {
  /** Points, 0 to 409.5. Without it a row has Excel's default height, 15 points (a line of Calibri 11). */
  height?: number
  /**
   * The height is one Excel fitted, by AutoFit or as text that wraps was typed into the row, not one set by hand (xlsx
   * writes it without `customHeight`): text typed into the row fits it again (`fitRows`), as a row of the default height.
   */
  bestFit?: boolean
  style?: number
  numberFormat?: NumberFormat
  /**
   * Hidden (xlsx `hidden`), by hand or by a filter: SUBTOTAL with codes 101–111 and AGGREGATE with options 1, 3, 5
   * and 7 skip the row's cells, and on a `filtered` sheet SUBTOTAL with codes 1–11 too, as Excel does. A row of
   * `height` 0 is hidden too, as Excel takes it. A collapsed group's rows are hidden rows.
   */
  hidden?: boolean
  /** The outline level (Excel's Group, xlsx `outlineLevel`), 1 to 7; a row without one is not grouped. */
  outlineLevel?: number
  /**
   * The mark of a collapsed group on its summary row (xlsx `collapsed`). Excel keeps it on every summary row whose group
   * has all its rows hidden, however they were hidden, and drops it when one shows; a summary row inside a group whose rows
   * are all hidden keeps the mark it had, so an inner group stays collapsed when the outer one is shown again. A book may
   * hold other marks (`setLayout` sets the field as given): `outlineGroups` and the commands read the marks as Excel
   * keeps them, and `writeXlsx` writes them so.
   */
  collapsed?: boolean
}
```

### `RowsFit`

Excel's AutoFit of rows `5` or `2:9` (a double click on the border below a row's number): each row gets the height of
its highest cell as Excel for Mac fits it ([`fitRows`](#xlcalc-fitrows)), marked `bestFit` and shown if it was hidden; a row as high as
the default gets no height. The heights are computed from the values and layout the edits before it leave
(`docs/formats.md`, "Fitting rows").

```ts
interface RowsFit {
  sheet: string
  rows: string
  fit: true
}
```

### `SavedViewport`

A window over a saved result ([`viewportOf`](#xlcalc-viewportof)): as [`Viewport`](#viewport), without a session's revision, its cells [`SavedViewportCell`](#savedviewportcell).

```ts
interface SavedViewport extends Omit<Viewport, 'revision' | 'cells'> {
  cells: Record<string, SavedViewportCell>
}
```

### `SavedViewportCell`

A cell of a window over a saved result ([`viewportOf`](#xlcalc-viewportof)): as [`ViewportCell`](#viewportcell), its value, `text` and `color` as the result
has them; without `text` and `color` when the result was computed without the display. The `appearance` of
conditional formats only with [`SavedViewportOptions.appearance`](#savedviewportoptions): without it the cell's `text` shows the value by the
cell's own format, not a rule's.

```ts
type SavedViewportCell = ViewportCell
```

### `SavedViewportOptions`

[`viewportOf`](#xlcalc-viewportof)'s options: `getViewport`'s, and `appearance`.

```ts
interface SavedViewportOptions extends ViewportOptions {
  /**
   * Each cell also with the `appearance` its sheet's conditional formats give it, and its `text` and `color` by a number
   * format of theirs, as `getViewport` gives them in a session of the book on the result's clock; false by default. A
   * result holds no appearance: the rules are computed over the result's values, its formulas are not (TODAY in a rule
   * is the result's day). The first window with it over a book and a result reads all their values, the later ones
   * cost what the window and its rules read.
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

[`scanReferences`](#xlcalc-scanreferences)'s options.

```ts
interface ScanOptions {
  /** The locale the text is typed in, a loaded locale's tag (`LOCALES`) or a locale object; without it the text is a stored formula. */
  locale?: string | Locale
}
```

### `Session`

A session over a book: edits recompute only the cells that depend on them. The current clock of the
session (`clock`) is what a saved result should record: a full computation with it gives the same hash.

```ts
interface Session {
  /** Mixed cells, names and looks applied atomically, with one final recalculation. */
  applyChanges(operations: SessionOperation[], opts?: ChangesOptions): ChangesResult
  /** Dependent-value changes and errors without changing the session; use its clock when applying. */
  previewChanges(operations: SessionOperation[], opts?: PreviewChangesOptions): ChangesPreview
  /** One revision, multiple ranges, only the fields and number of cells requested. */
  readCells(ranges: RangeRef[], opts?: ReadCellsOptions): ReadCellsResult
  /** Search performed where the book lives, literal or by Excel's Find; output is paged, in Find's order. */
  findCells(ranges: RangeRef[], query: string, opts?: FindCellsOptions): ReadCellsResult
  /**
   * Excel's Replace All (`Range.Replace`): the cells `findCells` finds in the text they open with (Look in Formulas),
   * each one's text with the query's matches put as `replacement` is, entered back as its own text edited: a formula in
   * the book's language, a number by its format without the percent editor's `%`, a text that reads as something else
   * kept a text, an array of the old kind whole, from its first cell, wherever a cell of it was found. A `*` takes as
   * little as it can, the query's last one the rest. A text Excel would not take (a formula that no longer reads) leaves
   * its cell as it was (`refused`). No column widens and no rule of data validation is checked. One recomputation, one
   * revision and one step of the undo; none when nothing is found (`docs/live-books.md`).
   */
  replaceInCells(ranges: RangeRef[], query: string, replacement: string, opts?: ReplaceOptions): ReplaceResult
  /** Immediate formula dependents; source ranges stay ranges. */
  getDependents(ranges: RangeRef[], opts?: DependentsOptions): DependentsResult
  /** Totals over nonblank stored values and spills, without returning their cells. */
  summarizeRanges(ranges: RangeRef[], opts?: EditOptions): RangeSummary
  readonly clock: Clock
  /** The book's revision: 0 when opened, one more after each call that changed the book (`EditResult.revision`). */
  readonly revision: number
  /**
   * Edits of cells, applied atomically with one recomputation (`CellEdit`): each sets the fields it gives and keeps the
   * cell's others, as Excel keeps a cell's look when a value is typed over it; null clears the content, as Delete. An
   * edit that keeps the cell's value and formula (a new `style`, `numberFormat`, `meta`, `note` or `link`) recomputes nothing.
   * A value or formula put into a spill range stops the array spilling (#SPILL!); clearing it lets the array spill
   * again. Edits that change some cells of an old-kind array formula's area (`CellInput.array`) but not all of them
   * throw `InputError` and change nothing, as Excel refuses to change part of an array; an edit of its top-left cell that
   * keeps an array formula over the same area changes all of it. An edit of a table's header cell renames its column, as
   * Excel does: the structured references to it follow, a number or a boolean becomes its text, a cleared cell names the
   * column `Column1` (the first such name free, in the word of the book's locale: `Столбец1` in ru-RU), and a name
   * another column has already gives the column further right a number (`Price2`). A formula in a header cell throws
   * `InputError`.
   */
  setCells(edits: CellEdit[], opts?: EditOptions): EditResult
  /** Cells replaced whole, as `setCells` applies edits otherwise: what an input leaves out the cell loses, null removes the cell. */
  replaceCells(edits: CellReplace[], opts?: EditOptions): EditResult
  /**
   * The cells of rectangles cleared, as Excel's ClearContents (`contents`: values, formulas and hyperlinks), ClearFormats
   * (`formats`: styles and number formats, the conditional formats over them) and Clear (`all`: both, notes, meta, the
   * rules of data validation over them); whole rows and columns lose their own style and format too (`docs/edits.md`).
   */
  clearCells(ranges: RangeRef[], what: ClearKind, opts?: EditOptions): EditResult
  /**
   * The look of cells and ranges (`CellPatch`), applied in order and atomically, without the cells' values and
   * formulas, which stay: an empty cell of a rectangle gets a cell with the look, a cell left with nothing is removed, a
   * cell of an old-kind array formula's area or of a spill takes it and the array stays. A style by its fields goes over
   * each cell's (`docs/edits.md`); equal styles share one entry of the book's `styles`. Recomputes nothing.
   */
  patchCells(patches: CellPatch[], opts?: EditOptions): EditResult
  /**
   * Text typed into a cell, as Excel takes it in the book's locale, date system and clock, by the format the cell shows
   * (`parseEntry`): the cell keeps its style, meta, note and link, and a format it shows from its row or column is not
   * written into it. A number that does not fit a column of the default width widens it in the same step (`width`), as
   * Excel does, unless `widen` is false. Text with a line break wraps the cell (its style gets `alignment.wrap`, a style
   * of the book's table found or added), and the row of a cell that wraps fits it (`fitRows`) when nobody set its height
   * by hand, in the same step; a value the cell's rule of data validation does not take is written, with the
   * rule (`validation`), or with `refuseInvalid` refused, with the rule, and nothing written. A formula Excel does not
   * take, or a change of part of an array, is refused, with a message in the book's locale.
   */
  setCellText(sheet: string, cell: string, text: string, opts?: CellTextOptions): CellTextResult
  /** What `setCellText` would do, with nothing changed. */
  previewCellText(sheet: string, cell: string, text: string, opts?: CellTextOptions): CellTextPreview
  /**
   * Texts typed into cells at once (a paste of text, an entry into a selection), each read as `setCellText` reads it in
   * its cell, written with one recomputation, one revision and one step of the undo; a cell given twice takes the later
   * text, read by what the earlier one wrote. A column widens once, to the largest width a text typed alone into one of
   * its cells gives it (`widths`), unless `widen` is false; cells wrap and rows fit as `setCellText` makes them; the values the cells' rules of data validation do not take
   * after all the texts are written, with the cells (`validations`). A text `setCellText` refuses refuses the whole
   * call, with its cell, and nothing is written; with `refuseInvalid`, so does the cell whose value breaks the strictest
   * rule after all the texts are written, the first of that style.
   */
  setCellTexts(entries: CellTextEntry[], opts?: CellTextOptions): CellTextsResult
  /** What `setCellTexts` would do, with nothing changed. */
  previewCellTexts(entries: CellTextEntry[], opts?: CellTextOptions): CellTextsPreview
  /**
   * Text pasted from the clipboard at a cell, as Excel pastes it: the rows at its line breaks and the fields at its tabs
   * (`parseClipboard`, a field begun by a quote running to its closing quote) fill a rectangle from the cell, a short
   * row's missing fields and a line break after the last row clearing their cells; a field with a line break wraps its
   * cell, as `setCellText` wraps it, and fits no row. Each field is read as `setCellText` reads it, but an apostrophe stays part of the text, a
   * formula Excel does not take is pasted as text, a percent cell has no percent editor (`5` is 500%) and `€5` takes the
   * euro format; no column widens and no rule of data validation is checked. One recomputation, one revision and one
   * step of the undo. Refused, with the cell, where Excel does not paste: part of an array formula, a formula into a
   * table's header cell; and a rectangle past the sheet's edge.
   */
  pasteText(sheet: string, cell: string, text: string, opts?: EditOptions): PasteTextResult
  /** What `pasteText` would do, with nothing changed. */
  previewPasteText(sheet: string, cell: string, text: string, opts?: EditOptions): PasteTextPreview
  /**
   * A text entered into every cell of a range at once, as Excel's Control+Return enters it from the selection's active
   * cell (`active`, the range's top-left cell by default): read once, as `setCellText` reads it in the active cell, it
   * gives every cell what the active cell holds after it, a formula with its relative references moved as a copy from
   * the active cell (`copyFormula`), and the active cell's number format, where it has one. The columns widen as typing
   * widens them (`widths`), unless `widen` is false, and a text with a line break wraps every cell and fits their rows, as
   * `setCellText` does; a value the active cell's rule of data validation does not take is
   * written, with that rule's alert (`validation`), the only rule checked, or with `refuseInvalid` refused, with the
   * active cell. One recomputation, one revision and one step of the undo. A formula Excel does not take, part of an
   * array formula and a formula into a table's header cell are refused, with the cell, and nothing is written.
   */
  fillCellText(sheet: string, range: string, text: string, opts?: FillTextOptions): FillTextResult
  /** What `fillCellText` would do, with nothing changed. */
  previewFillCellText(sheet: string, range: string, text: string, opts?: FillTextOptions): FillTextPreview
  setNames(names: NameInput[], opts?: EditOptions): EditResult
  /** Non-live sessions only. */
  setClock(clock: Clock, opts?: EditOptions): EditResult
  /** A structure operation on the session's book, recomputed; refused as `editStructure` refuses it. A sheet hidden or shown recomputes nothing. */
  editStructure(op: StructureOp, opts?: EditOptions): SessionStructureResult
  /** Replaces the book's style table; every style that cells, rows and columns use must stay in it. Recomputes nothing; the shown text does not depend on styles. */
  setStyles(styles: CellStyle[], opts?: EditOptions): EditResult
  /**
   * Column widths, row heights, row and column styles and formats, hidden rows and columns, outline levels, merged
   * areas, the filter mode, frozen panes, conditional formats, rules of data validation and outline settings; applied
   * atomically. Merged areas stop arrays spilling over them: the formulas whose spill range an added or removed area
   * meets are recomputed. Hidden rows and the filter mode change what SUBTOTAL and AGGREGATE count: their formulas over
   * the rows concerned are recomputed. Nothing else is recomputed. With `display`, the cells whose shown text or colour
   * a width, a format or a merged area changed are listed too.
   * A row height set or removed without bestFit clears its previous fitted mark; bestFit explicitly given is honored.
   */
  setLayout(edits: LayoutEdit[], opts?: EditOptions): EditResult
  /**
   * The session's book with every change so far: what a host saves as a version. Its objects are the
   * session's own and later calls change them, so do not modify it; serialize or copy what must stay.
   */
  getBook(): Book
  /**
   * The session's book without its cells: the styles, names and settings, and every field of each sheet but `cells` and `looks`
   * (columns, rows, merged areas, tables, frozen panes, conditional formats, data validation and the rest). It costs
   * what the layout holds, not the cells; its objects are the session's own, as `getBook()`'s are.
   */
  getLayout(): BookLayout
  /**
   * Where a sheet's cells end: the last row and column that a cell of its input (a value, a formula, or a style, a
   * format, meta, a note or a link alone), a spilled value (the whole spill range), a merged area or a table reaches.
   * Settings of whole rows and columns, conditional formats and data validation do not count. It costs the sheet's
   * columns, spills and cells without a value, not a pass over all its cells.
   */
  getUsedRange(sheet: string): SheetExtent
  /**
   * The cell Control and an arrow key come to from `cell`, its address, as Excel's `Range.End` gives it
   * (`docs/live-books.md`): from a cell that holds something with such a cell next to it, the last of that block; else
   * the next cell that holds something, or the sheet's edge (row 1 or 1,048,576, column A or XFD), where a cell at the
   * edge stays. A value or a formula holds something, whatever it shows (an empty text, zero, an error), and so does
   * each cell of a spill and of an old-kind array; a look, a note or a link alone does not. Hidden rows and columns are
   * not there: by hand, by a filter or by a size of 0, `End` never stops on them and a block goes on across them, and
   * with every line hidden past a cell the edge is the last line shown. A merged area is its cells, its value in the
   * top-left one. It costs the cells of the line it goes along, up to where the line's values end, not a pass over the sheet.
   */
  getEnd(sheet: string, cell: string, direction: EndDirection): string
  explain(sheet: string, cell: string, opts?: ExplainOptions): Explanation
  /**
   * The cells of a rectangle of a sheet (`B2:D9`, `B2`, whole columns `B:D`, whole rows `2:9`), or of several (a cell in
   * two of them is listed once), as `getResult()` gives them, by address: the same cells (spilled ones too, empty ones
   * left out) with the same fields, and each cell's own `style`, `note` and `link`, an old-kind array formula's `array`,
   * and the `appearance` the sheet's conditional formats give it by the current values (an empty cell with one is listed
   * with the value null). For a window over a sheet: it costs what the rectangles hold, not the book; the statistics of
   * a rule's cells (lowest, highest, percentiles) are kept until an edit changes them. Edits do not list cells whose
   * conditional appearance changed: they give the rules' cells as `repaint`, for a host to read again where its window meets them.
   */
  getCells(sheet: string, range: string | string[]): Record<string, CellView>
  /**
   * Plain-text clipboard data for a range: computed values by their effective formats and the book's locale,
   * without column clipping, skipping hidden rows and columns. Explicit rectangles keep selected blanks;
   * whole rows/columns stop at the used extent, and an empty sheet gives one empty field. Changes nothing.
   */
  copyText(sheet: string, range: string, opts?: CopyTextOptions): string
  /** One cell as `getCells(sheet, cell)` gives it; undefined for a cell `getCells` leaves out. */
  getCell(sheet: string, cell: string): CellView | undefined
  /**
   * A window over a sheet for a grid (`B2:D9`, `B:D`, `2:9`): its cells with their styles and number formats resolved,
   * empty ones with a style or a number format of their row or column too, the widths and heights of its lines, the
   * merged areas and tables that meet it and the sheet's frozen panes (`Viewport`); only some of its rows and columns
   * with `opts.rows` and `opts.columns`. It costs what the window holds, as `getCells`, not the book.
   */
  getViewport(sheet: string, range: string, opts?: ViewportOptions): Viewport
  /**
   * Why a cell holds an error, or a circular reference stops it, and the cell it began in (`CellIssueDetail`), without the
   * book's result; null for a cell without one.
   */
  getCellIssue(sheet: string, cell: string): CellIssueDetail | null
  /** The data validation of a cell (`CellValidation`), its list computed now; null when no rule is over the cell. */
  getValidation(sheet: string, cell: string): CellValidation | null
  /**
   * Whether the cell holding `input` (null: cleared) would meet its rule, as Excel checks a cell's value
   * (`Range.Validation.Value`): the value the input computes to, in the book with it. Changes nothing; a cell without a
   * rule takes anything. Text typed into the cell is `parseEntry`'s input first.
   */
  validateEntry(sheet: string, cell: string, input: CellInput | null): EntryCheck
  getResult(): ComputeResult
  /** An independent copy of the book and its result at one moment; later edits and changes to the copy cannot affect each other. */
  getSnapshot(): BookSnapshot
  /** The values hash `getResult()` gives, without building the result. */
  getHash(): string
  /**
   * Takes back the last change of the book, as Excel's Undo: the whole of one call that changed it (`UndoCall`),
   * recomputed; null when there is nothing to undo. A change after an undo forgets what `redo` could do again.
   */
  undo(opts?: EditOptions): UndoResult | null
  /** Does again the last change `undo` took back; null when there is none. */
  redo(opts?: EditOptions): UndoResult | null
  /** The steps `undo` and `redo` can take now. */
  readonly history: SessionHistory
  /**
   * Calls `listener` after each call that changes the book, whoever makes it: an edit, a structure operation, `undo`,
   * `redo` (`SessionChange`: what the call answered of `EditResult`, and its name). Not after a preview, a read, a
   * refusal or a call that changed nothing. A listener runs when the call is done, before it returns, so the session it
   * reads is the one after the change; listeners run in the order they subscribed, and changes reach them in the order
   * of their revisions: a change a listener makes reaches every listener after the one it answers. A listener that
   * throws stops neither the call nor the other listeners; its error is thrown again in a microtask. Answers the
   * function that unsubscribes; `close()` unsubscribes every listener.
   */
  subscribe(listener: (change: SessionChange) => void): () => void
  close(): void
}
```

### `SessionChange`

A change of a session's book as its subscribers get it ([`Session.subscribe`](#session)): the [`EditResult`](#editresult) of the call that made
it (the revision after it, the cells it changed, the sheets whose layout changed, the style table, how addresses
moved, where to repaint) and the call's name.

```ts
interface SessionChange extends EditResult {
  method: ChangingMethod
}
```

### `SessionHistory`

What a session can undo and redo now: the steps kept each way.

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

Session options. A session keeps one clock, `clock` or by default the moment of opening and a random
seed, until `setClock`. A live session (`live: true`) takes a fresh clock and a new seed on every
recomputation instead. Edits are compiled in the `execution` the session was opened with (`closures` by
default). With `display` (the default) results and changes carry what the cells show. `limits` bound every call;
a call over them throws [`CalcError`](#xlcalc-calcerror) and leaves the session as it was. `undo` bounds the steps [`Session.undo`](#session) keeps.
`cached`, as [`compute`](#xlcalc-compute) takes it, gives the formulas a circular reference stops the values the book's file stores; the
session's book holds them as `held` from then on.

```ts
type SessionOptions = ({ clock?: Clock; live?: false } | { live: true }) & { externals?: Externals; execution?: Execution; display?: boolean; limits?: CalcLimits; undo?: UndoOptions; cached?: CachedValues }
```

### `SessionStructureResult`

[`Session.editStructure`](#session): the edit's result and where the cells went, without the book (`getBook()` gives it), in
every kind of session. `changes` are the cells whose value or formula differs from those of the cell that was at
their old place (`mapCell`), by new address, a cell no old cell came to counting as empty before; with `display`, the
cells whose shown text or colour differs too. An interface that moves its shown cells by `moved` (or `mapCell`) and
applies these has the new ones.

```ts
type SessionStructureResult =
  | (EditResult & {
      ok: true
      /** Where a cell of the book before the operation is now; null if it was deleted or overwritten, or its sheet is none of that book's. */
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

Where a sheet's cells end ([`Session.getUsedRange`](#session)): everything lies in its first `rows` rows and `columns` columns; 0 and 0 for an empty sheet.

```ts
interface SheetExtent {
  rows: number
  columns: number
}
```

### `SheetInput`

A sheet of a book ([`BookInput.sheets`](#bookinput)): its name, its cells by address, and what else a sheet keeps (rows, columns, merged areas, tables, looks, rules, frozen panes).

```ts
interface SheetInput {
  name: string
  /** Cells by A1 address (`B12`). */
  cells: Record<string, CellInput>
  /**
   * The look of the cells that have no entry in `cells`, by rectangles, each over those before it: a sheet formatted
   * across many empty cells (whole rows to XFD) is held without an object per cell. A cell without an entry shows the
   * style and the number format of the last rectangle that covers it, both from it (a rectangle without a field gives
   * the cell none of its own); a cell with an entry shows its own fields, whatever the rectangles say. A field a cell
   * has none of comes from its row, else its column, as for an entry. `[{ range: 'A20:XFD22', style: 0 },
   * { range: 'R20:R22', style: 1, numberFormat: '0.0%' }]` is the same sheet as an entry `{ style: 0 }` in each cell of
   * A20:XFD22 but R20:R22, which hold `{ style: 1, numberFormat: '0.0%' }`. A book is taken with every entry that has
   * only a style or a number format moved here, and the calls that give a book back give this list in a form of their
   * own (`LookRange`).
   */
  looks?: LookRange[]
  /** Columns by letters, one (`C`) or a span (`B:D`) as xlsx `<col min max>`; spans do not overlap. */
  columns?: Record<string, ColumnInput>
  /** Rows by number, one (`5`) or a span (`5:9`); spans do not overlap. */
  rows?: Record<string, RowInput>
  /** Merged areas (`B2:D3`), as xlsx `<mergeCell>`: two cells or more each, none overlapping another. */
  merges?: string[]
  /**
   * A filter of the sheet has criteria, its autofilter's or a table's (Excel's filter mode): every hidden row of the
   * sheet counts as one a filter hid, which SUBTOTAL skips with codes 1–11 too and a copy leaves out, as Excel does
   * whatever hid the row.
   */
  filtered?: boolean
  /** The sheet's tables (Excel's list objects); they do not overlap, and no merged area or array area of several cells meets one. */
  tables?: TableInput[]
  /**
   * Frozen panes (Excel's Freeze Panes): the first `rows` rows and the first `columns` columns stay in view as the sheet
   * scrolls, one of them at least (`{ rows: 1 }` freezes the top row). Not part of computation.
   */
  frozen?: FrozenPanes
  /** Saved window state from xlsx; does not affect computation or inputHash. */
  view?: SheetView
  /**
   * The sheet's conditional formats (Excel's conditional formatting rules) in priority order, the first one first. They
   * do not change values: a session gives what they show in `CellView.appearance` (`docs/conditional-formats.md`).
   */
  conditionalFormats?: ConditionalFormat[]
  /**
   * Data validation (Excel's Data Validation): the rules of the sheet's cells, at most one rule over a cell. Not part of
   * computation: a session gives a cell's rule and its list (`Session.getValidation`) and checks a value against it
   * (`Session.validateEntry`).
   */
  validations?: DataValidation[]
  /** Where the summary rows and columns of the sheet's groups are (`RowInput.outlineLevel`); below and right by default. */
  outline?: OutlineSettings
  /**
   * The sheet is hidden (Excel's Hide, xlsx `state="hidden"`), or very hidden (`'very'`, xlsx `state="veryHidden"`, a
   * sheet Excel's window does not offer to show again, only its object model). Not part of computation: formulas read a
   * hidden sheet as any other. A book keeps one visible sheet at least.
   */
  hidden?: boolean | 'very'
}
```

### `SheetLayout`

A sheet without its cells and their looks ([`Session.getLayout`](#session)).

```ts
type SheetLayout = Omit<Sheet, 'cells' | 'looks'>
```

### `SheetNameCheck`

[`validateSheetName`](#xlcalc-validatesheetname)'s answer; `at` is the position of the character Excel does not take.

```ts
type SheetNameCheck = { ok: true } | Refusal<SheetNameCode>
```

### `SheetNameCode`

Why Excel does not take a sheet name: empty, over 31 characters, one of `: \ / ? * [ ]`, an apostrophe at either end, `History`.

```ts
type SheetNameCode = 'empty' | 'too_long' | 'character' | 'apostrophe' | 'reserved'
```

### `SheetOutline`

A sheet's groups ([`outlineGroups`](#xlcalc-outlinegroups)), each axis by first line, the outer group before the inner where two begin together.

```ts
interface SheetOutline {
  rows: OutlineGroup[]
  columns: OutlineGroup[]
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
}
```

### `StructureOp`

One structure operation ([`editStructure`](#xlcalc-editstructure)). Rows are 1-based numbers and columns are letters, as in A1
addresses; `count` is 1 by default; sheets and tables are matched case-insensitively. Tables move with their cells
and grow and shrink with the rows and columns inserted and deleted inside them; an inserted row gets the formulas of
the calculated columns, an inserted column a name (`Column1`), and references to a deleted column become #REF!.

```ts
type StructureOp =
  /** Inserts `count` rows before `row`: rows from `row` on move down; the new rows take the format of the row above. */
  | { op: 'insertRows'; sheet: string; row: number; count?: number }
  /** Deletes rows `row` … `row + count - 1`: the rows below move up. */
  | { op: 'deleteRows'; sheet: string; row: number; count?: number }
  /** Inserts `count` columns before `column`: columns from `column` on move right; the new ones take the format of the column on the left. */
  | { op: 'insertColumns'; sheet: string; column: string; count?: number }
  /** Deletes `count` columns from `column` on: the columns to the right move left. */
  | { op: 'deleteColumns'; sheet: string; column: string; count?: number }
  /** Cut and paste: the range moves so that its top-left cell lands on `to`, replacing what was there; cells keep their `meta`. */
  | { op: 'moveRange'; from: RangeRef; to: CellRef }
  /**
   * Copy and paste: the range is pasted at `to`; a target that is a whole multiple of the source's size is
   * tiled with copies (fill down and fill right are copies of the first row or column). Copies have no
   * `meta` unless `withMeta`.
   */
  | { op: 'copyRange'; from: RangeRef; to: RangeRef; withMeta?: boolean }
  /** A new name by Excel's rules (see `BookInput.sheets`) that no other sheet has. */
  | { op: 'renameSheet'; sheet: string; name: string }
  | { op: 'deleteSheet'; sheet: string }
  /**
   * A new empty sheet named `name` (by Excel's rules, not taken), before the sheet `before`, else after the last one.
   * A reference across sheets takes in a sheet inserted between its ends.
   */
  | { op: 'insertSheet'; name: string; before?: string }
  /**
   * The sheet moved before the sheet `before`, else after the last one; before a hidden sheet it goes before the next
   * visible one, as Excel moves it. A reference across sheets covers the sheets between its ends in the new order; an end
   * that moves stays its end until it passes the other end, then the next sheet inward takes its place.
   */
  | { op: 'moveSheet'; sheet: string; before?: string }
  /** The sheet hidden (Excel's Hide Sheet), very hidden with `very` (see `SheetInput.hidden`); a book keeps one visible sheet. */
  | { op: 'hideSheet'; sheet: string; very?: boolean }
  /** The sheet shown again (Excel's Unhide Sheet). */
  | { op: 'unhideSheet'; sheet: string }
  /** A table over cells of `sheet` (Excel's Format as Table): with a header row its header cells take the columns' names. */
  | { op: 'addTable'; sheet: string; table: TableInput }
  /** A new name for the table, by the rules of `TableInput.name`; every structured reference to it follows. */
  | { op: 'renameTable'; table: string; name: string }
  /** The table and its cells deleted (Excel's Delete Table): structured references to it become #REF!. */
  | { op: 'deleteTable'; table: string }
  /**
   * The table turned into plain cells (Excel's Convert to Range): structured references to it become references to its
   * cells with `$` and the sheet (`'S'!$B$2:$B$9`; a reference to the formula's own row keeps its row relative, `'S'!$B2`).
   */
  | { op: 'convertTable'; table: string }
  /**
   * Excel's Group and Ungroup of rows `5:7` or columns `B:D` (one line: `5`, `C`). Group puts each line one outline level
   * deeper, those at the deepest level (7) staying there, and is refused (`outline_deep`) when all of them are; lines next
   * to each other at a level make one group, so a group added right after another of its level joins it
   * (`outlineGroups`). Ungroup takes a level off each line that has one, refused (`not_grouped`) when none has. Hidden
   * lines stay hidden.
   */
  | { op: 'group' | 'ungroup'; sheet: string; rows: string }
  | { op: 'group' | 'ungroup'; sheet: string; columns: string }
  /**
   * Excel's Show Detail and Hide Detail, a group's `+` and `−` button: at a summary line (`OutlineGroup.summary`) on the
   * lines it sums up, at a line of a group on the group of the line's level that holds it, through its summary line. Hiding
   * hides the lines and marks the summary line `collapsed`; showing shows them but the lines of the collapsed groups inside,
   * and drops the mark. Refused at a line in no group or in a group with no summary line, past the sheet's edge
   * (`not_grouped`), and where the detail is shown already, or hidden (`detail`).
   */
  | { op: 'showDetail' | 'hideDetail'; sheet: string; row: number }
  | { op: 'showDetail' | 'hideDetail'; sheet: string; column: string }
  /**
   * Excel's level buttons (`Outline.ShowLevels`): the lines of outline level `rows` (or `columns`) and deeper hidden, the
   * other grouped lines shown, the collapsed marks as Excel keeps them; 8 shows every group. Lines outside groups stay as
   * they are.
   */
  | { op: 'showLevels'; sheet: string; rows?: number; columns?: number }
  /**
   * Excel's Merge Cells over `range` (`B2:D3`, whole columns and rows too): the range, grown by the merged areas it
   * meets, becomes one merged area (`SheetInput.merges`) that keeps its top-left cell; when that cell is empty, the
   * area's first cell holding a value or a formula, row by row, moves into it as cut and paste moves it. The other cells
   * lose their values and formulas and take the top-left cell's style and number format. With `across`, each row of the
   * range is merged on its own (Merge Across); with `center`, the area is centred as well (Merge & Center). A range that
   * meets a table is refused (`table`), and one that meets an array formula of the old kind (`array`).
   */
  | { op: 'mergeCells'; sheet: string; range: string; across?: boolean; center?: boolean }
  /** Excel's Unmerge Cells: the merged areas that meet `range` are split back into their cells. */
  | { op: 'unmergeCells'; sheet: string; range: string }
  /**
   * Excel's fill handle (`Range.AutoFill`): the cells of `range` go on over `to`, which holds `range` and reaches past it
   * along one axis, down, up, right or left (`A1:B2` over `A1:B9`, or over `A1:F2`). The cells added take the source's
   * cells in turn, as copies do (formulas, styles, number formats), the part of the source nearest to it at the far end
   * when `to` is not a whole number of sources long; refusals are a copy's. Each column of a fill down or up, and each row
   * of one right or left, goes on by itself, and what goes on as a series takes its next values instead (`FillType`,
   * docs/structure.md "The fill handle"). A `to` of another shape is an `InputError`.
   */
  | { op: 'fillRange'; sheet: string; range: string; to: string; type?: FillType }
```

### `StructureRefusal`

Why Excel would refuse an operation.

```ts
type StructureRefusal =
  /** Cells would leave the sheet (past XFD or row 1,048,576), or a paste does not fit. */
  | 'off_sheet'
  /** A new sheet name that breaks Excel's rules or is taken (sheet names compare case-insensitively). */
  | 'sheet_name'
  /** A book keeps one visible sheet at least: its last visible sheet cannot be deleted or hidden. */
  | 'last_sheet'
  /** A very hidden sheet (`SheetInput.hidden` 'very') cannot be deleted or moved, as Excel refuses: it is shown again first. */
  | 'very_hidden'
  /** A move or copy would cut through a merged area, at its source or where it lands. */
  | 'merged'
  /** A copy that leaves out rows a filter hid, pasted over hidden rows of a filtered sheet (several areas onto several areas). */
  | 'filtered'
  /** The operation would change part of the area of an array formula of the old kind (`CellInput.array`). */
  | 'array'
  /**
   * The operation would shift or cut part of a table as Excel refuses it: rows or columns inserted into or deleted from
   * two tables at once, a table's header row deleted without the rest, a move that cuts a table, a table over another.
   */
  | 'table'
  /** A table name that breaks the rules of `TableInput.name` or that a table or a defined name of the book has. */
  | 'table_name'
  /** Group of lines that are all at the deepest outline level, 7. */
  | 'outline_deep'
  /** Ungroup of lines none of which has an outline level; Show or Hide Detail at a line in no group, or in a group with no summary line. */
  | 'not_grouped'
  /** Show Detail of a group shown already, Hide Detail of one hidden already (its summary line marked `collapsed`). */
  | 'detail'
```

### `StructureResult`

[`editStructure`](#xlcalc-editstructure)'s answer.

```ts
type StructureResult =
  | {
      ok: true
      /** The new book; the input book is not changed, unchanged sheets and cells are shared. */
      book: Book
      /** Where a cell of the old book is now; null if it was deleted or overwritten, or its sheet is none of the old book's. */
      mapCell(ref: CellRef): CellRef | null
    }
  /** A refusal: its code, and a message in the language of the book's locale (English for en-US). */
  | Refusal<StructureRefusal>
```

### `StylePatch`

Fields of a style to change ([`CellPatch.set.style`](#cellpatch)): a value sets a field, null removes it, a field left out stays.

```ts
interface StylePatch {
  font?: { [K in keyof FontStyle]?: FontStyle[K] | null } | null
  fill?: { color: Color } | null
  border?: { [K in keyof NonNullable<CellStyle['border']>]?: BorderLine | null } | null
  alignment?: { [K in keyof AlignmentStyle]?: AlignmentStyle[K] | null } | null
}
```

### `TableColumn`

A column of a table ([`TableInput.columns`](#tableinput)): its header's name and what its cells and totals row compute.

```ts
interface TableColumn {
  /** 1 to 255 characters, unique in the table in any case; the header cell's text. */
  name: string
  /**
   * The formula of a calculated column (stored text, `Sales[[#This Row],[Qty]]*2`; a leading `=` is dropped), which a row
   * inserted into the table gets in this column, as Excel fills it. The column's cells hold their own formulas.
   */
  formula?: string
  /**
   * What the column's cell of the totals row is, as an xlsx file records it: one of Excel's totals functions, a
   * SUBTOTAL over the column that the cell holds as its formula (`custom`: a formula of its own), or a label. The cell
   * computes what it holds.
   */
  totals?: { function: TotalsFunction } | { label: string }
}
```

### `TableInput`

A table (Excel's list object): a rectangle of a sheet with a header row of column names, data rows and a totals row.
Formulas refer to its parts by name with structured references (`Sales[Amount]`, `SUM(Sales[[#Totals],[Qty]])`),
which follow the table as rows and columns are inserted and deleted in it. A formula stores a reference to its own
row as an xlsx file does, `Sales[[#This Row],[Amount]]`; `Sales[@Amount]`, and `[@Amount]` inside the table, are read too.

```ts
interface TableInput {
  /**
   * Excel's rules for names: a letter of any script, `_` or `\` first, then letters, digits, `_`, `.`, `\` and `?`, at
   * most 255 characters; not a cell address (`T2`) or R1C1 (`R`, `C`, `R1C1`), not TRUE or FALSE, and not the book's
   * locale's words for them either, in any case (ИСТИНА and ЛОЖЬ in ru-RU, as Excel for Windows ru-RU refuses them).
   * Tables and defined names share one namespace: no other table and no defined name of the book has the name, in any case.
   */
  name: string
  /** The table's rectangle, its header and totals rows included (`B2:E10`), with one data row at least. */
  range: string
  /** The first row is a header row (true by default): its cells hold the columns' names as text, and no formulas. */
  headers?: boolean
  /** The last row is a totals row (false by default). */
  totals?: boolean
  /** The columns from left to right, one for each column of the range. */
  columns: TableColumn[]
}
```

### `TableReference`

A structured reference (`Sales[Qty]`, `[@Qty]`): the table as written (absent inside a table, `[@Qty]`), the column or
the first and last of a span, and the specifiers in stored words (`#Headers`, `#This Row` for `@`); a host lays it on
the table's range.

```ts
interface TableReference extends ReferenceSpan {
  kind: 'table'
  table?: string
  columns?: [string] | [string, string]
  items?: ('#All' | '#Data' | '#Headers' | '#Totals' | '#This Row')[]
}
```

### `TextFormat`

The cell's text holds `text`, does not, begins or ends with it, in any case. Excel computes the rule's formula, which
it writes beside the text (`NOT(ISERROR(SEARCH("ab",A1)))`); without `formula` the rule has that one.

```ts
interface TextFormat extends ConditionalBase {
  type: 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith'
  text: string
  formula?: string
  style?: ConditionalStyle
}
```

### `TimePeriodFormat`

The cell holds a date in the period, by the book's clock (`TODAY()`); computed by `formula`, Excel's without it.

```ts
interface TimePeriodFormat extends ConditionalBase {
  type: 'timePeriod'
  period: 'today' | 'yesterday' | 'tomorrow' | 'last7Days' | 'thisWeek' | 'lastWeek' | 'nextWeek' | 'thisMonth' | 'lastMonth' | 'nextMonth'
  formula?: string
  style?: ConditionalStyle
}
```

### `TopFormat`

The `rank` highest numbers of the cells (the lowest with `bottom`), or the `rank` percent of them with `percent`.

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

### `Translation`

The result of translating formula text: the text, or a check code with a message in the locale's language and the
position in the given text: `syntax`, or `external_ref` for a reference to a workbook the given links do not hold.

```ts
type Translation = { ok: true; formula: string } | (Refusal<'syntax' | 'external_ref'> & { at: number })
```

### `UndoCall`

The calls a session can undo: those that change its book. `setClock` and a live session's new clock are not steps.

```ts
type UndoCall = 'applyChanges' | 'setCells' | 'replaceCells' | 'clearCells' | 'patchCells' | 'setCellText' | 'setCellTexts' | 'pasteText' | 'fillCellText' | 'replaceInCells' | 'setNames' | 'setStyles' | 'setLayout' | 'editStructure'
```

### `UndoOptions`

The steps a session keeps for `undo`: `depth` of them, 100 by default as in Excel (0 keeps none), holding `maxMb`
megabytes at most by the engine's estimate, 64 by default, in a pool with `maxHeapMb` its share of a thread's heap
(`maxHeapMb / (8 × maxSessionsPerThread)`, `docs/async.md`). Past either the oldest steps go; a step larger than
`maxMb` alone is not kept, and the steps before it go too, as Excel forgets what it cannot undo.

```ts
interface UndoOptions {
  depth?: number
  maxMb?: number
}
```

### `UndoResult`

A step of a session undone or done again ([`Session.undo`](#session), [`Session.redo`](#session)): the edit's result, as the call gave it.

```ts
interface UndoResult extends EditResult {
  /** The call whose change it was. */
  call: UndoCall
  /** The calls that change cells: the cells whose input changed back, so a host can read their note and link again. */
  cells?: CellRef[]
  /**
   * `editStructure`: where a cell of the book before this call is now, null when no cell of the book has it or its sheet
   * is none of that book's; an interface that moves its shown values by it and applies `changes` has the new values, as
   * after `editStructure`.
   */
  mapCell?(ref: CellRef): CellRef | null
}
```

### `ValidateContext`

What [`validateFormula`](#xlcalc-validateformula) checks a formula against: the book's sheets, names and tables; without them they are not checked.

```ts
interface ValidateContext {
  /** Sheets of the book; without them sheets are not checked. */
  sheets?: string[]
  /** Names of the book; without them names are not checked. */
  names?: string[]
  externalFunctions?: string[]
  /**
   * Tables of the book with their columns; without them structured references are not checked. With them the canonical
   * text spells tables and columns as they are named, and a bare table name is a table, not a name.
   */
  tables?: { name: string; columns: string[] }[]
  /**
   * The table of the formula's cell: a structured reference without a table name is to it (`[@Qty]`), as in a cell of
   * the table, and the canonical text names it (`Sales[[#This Row],[Qty]]`). Without it such a reference is refused
   * with `tables` (`unknown_name`) and stays without a name otherwise.
   */
  table?: string
  /**
   * The book's links to other workbooks (`BookInput.externalBooks`): a reference to one is checked against its sheets
   * and names. A reference to a book the links do not hold is refused (`external_ref`), with or without them.
   */
  externalBooks?: ExternalBook[]
}
```

### `ValidateResult`

What [`validateFormula`](#xlcalc-validateformula) answers: the formula's canonical text and what it names, or why it does not read and where.

```ts
type ValidateResult =
  /** `tables`: the tables the formula's structured references name (and bare table names, with `ctx.tables`). */
  | { ok: true; canonical: string; functions: string[]; sheets: string[]; names: string[]; externals: string[]; tables: string[] }
  | (Refusal<CheckCode> & { at: number })
```

### `ValidationAlert`

A value a cell's rule of data validation does not take ([`Session.setCellText`](#session)): the rule, its alert's style and message.

```ts
interface ValidationAlert {
  valid: false
  rule: DataValidation
  severity: 'stop' | 'warning' | 'information'
  message: string
}
```

### `ValidationOperator`

How a value compares with `formula1` (and `formula2`): between and not between them take both, inclusive.

```ts
type ValidationOperator = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual'
```

### `ValidationRefusal`

A value a cell's rule of data validation does not take, refused ([`Session.validateEntry`](#session)): the rule, its alert's style and message.

```ts
type ValidationRefusal = Refusal<'validation'> & { rule: DataValidation; severity: ValidationAlert['severity'] }
```

### `ValidationType`

What a rule of data validation lets a cell hold: any value (`any`, a rule for its input message alone), a whole number,
a number, a date or a time compared with the rule's formulas, a text whose length compares so, one of a list, or a value
for which a formula (`custom`) is TRUE.

```ts
type ValidationType = 'any' | 'whole' | 'decimal' | 'date' | 'time' | 'textLength' | 'list' | 'custom'
```

### `Viewport`

What a grid needs to draw a window of a sheet ([`Session.getViewport`](#session)).

```ts
interface Viewport {
  /**
   * The rectangle given (`B2:D9`): the one asked for, or as far as its first rows (of `ViewportOptions.rows`, with them)
   * that fit when it holds more than `maxCells` cells.
   */
  range: string
  /** The window asked for held more than `maxCells` cells: `range` is the part given. */
  truncated: boolean
  /** By address, row by row: every cell with something to show, an empty one with its row's or column's style or number format, an appearance of a conditional format, a note or a link too. */
  cells: Record<string, ViewportCell>
  /** The styles the window's cells have or show (`ViewportCell.style`, `effectiveStyle`), by their index in the book's `styles`: copies, the answer's own. */
  styles: Record<number, CellStyle>
  /** The columns of `range` left to right, or those of `ViewportOptions.columns`. */
  columns: ViewportColumn[]
  /** The rows of `range` top to bottom, or those of `ViewportOptions.rows`. */
  rows: ViewportRow[]
  /** The merged areas that meet the window's cells, whole. */
  merges: string[]
  /** The tables that meet the window's cells, whole, in the sheet's order. */
  tables: ViewportTable[]
  /** The sheet's frozen panes (`SheetInput.frozen`): its first `rows` rows and `columns` columns stay in view, 0 for none. */
  frozen: { rows: number; columns: number }
  /** Where the sheet's cells end, as `Session.getUsedRange`. */
  usedRange: SheetExtent
  /** The session's revision the window is of (`Session.revision`): a host takes the changes of later edits over it (`EditResult.revision`). */
  revision: number
}
```

### `ViewportCell`

A cell of a window ([`Session.getViewport`](#session)): its value as `getResult()` gives it, with the cell's whole look resolved, its
style and number format as rows and columns give them, what its conditional formats show, its note and link.

```ts
interface ViewportCell extends Partial<CellDisplay> {
  value: CellValue
  /** The canonical formula text, as `CellResult.formula`. */
  formula?: string
  /** With `formulaLocal`: the formula as Excel shows it in the book's locale (`Range.FormulaLocal` without `=`), inside its table as Excel shows it there (`[@Qty]`). */
  formulaLocal?: string
  spill?: string
  spillFill?: CellValue
  spillOf?: string
  /** On the top-left cell of an array formula of the old kind, its area (`CellInput.array`). */
  array?: string
  meta?: JsonValue
  /** The cell's own style, an index of the book's `styles` as `CellView.style`; its object is in `Viewport.styles`. */
  style?: number
  /**
   * The style the cell shows, an index of the book's `styles` (its object in `Viewport.styles`): its own, else its row's,
   * else its column's. It is always one of the table's: a cell where a styled row crosses a styled column shows the row's,
   * as Excel shows a file's empty cell there, and the two merged are a style of the cell's own (`Session.patchCells`).
   */
  effectiveStyle?: number
  /** The number format the cell shows by, as a code: its own, else its row's, else its column's, a built-in format by the code of the book's locale (14 is `dd.mm.yyyy` in ru-RU). */
  numberFormat?: string
  /** Where `numberFormat` comes from. */
  numberFormatSource?: 'cell' | 'row' | 'column'
  /** What the sheet's conditional formats give the cell by the current values; with a number format of theirs, `text` and `color` show the value by it. */
  appearance?: ConditionalAppearance
  note?: CellNote
  link?: CellLink
  /** With `diagnostics`, on a cell with an error or a formula a circular reference stops: as `Session.getCellIssue`. */
  issue?: CellIssueDetail
  /**
   * With `editText`: the text the cell opens with for editing, as Excel's formula bar shows it: a formula with `=` in the
   * book's locale (`formulaLocal`), in every cell of an old-kind array formula's area its formula (without the braces
   * the bar shows around it), a value as `entryText` writes it by the cell's number format, and empty for a cell
   * without a value or formula of its own (an empty one, one a dynamic array spilled into).
   */
  editText?: string
}
```

### `ViewportColumn`

A column of a window: its number (A is 1), its width as [`ColumnInput.width`](#columninput), Excel's default where it has none.

```ts
interface ViewportColumn {
  column: number
  width: number
  hidden?: true
}
```

### `ViewportOptions`

[`Session.getViewport`](#session)'s and [`viewportOf`](#xlcalc-viewportof)'s options.

```ts
interface ViewportOptions {
  /** Each formula also as the book's locale shows it (`ViewportCell.formulaLocal`); false by default. */
  formulaLocal?: boolean
  /**
   * Each cell with an error, and each formula a circular reference stops, also with its reason and the cell it began in
   * (`ViewportCell.issue`); false by default.
   */
  diagnostics?: boolean
  /** The most cells a window gives, 20,000 by default: of a larger one, its first rows that fit (`Viewport.truncated`). */
  maxCells?: number
  /**
   * Only these rows of `range` (the rows a filter shows, say): their numbers on the sheet, from 1, in ascending order,
   * each once. The window gives them in this order, and only the merged areas and tables that meet its cells.
   */
  rows?: number[]
  /** Only these columns of `range`, by number (A is 1), in ascending order, each once, as `rows`. */
  columns?: number[]
  /** Each cell also with the text it opens with for editing (`ViewportCell.editText`); false by default. */
  editText?: boolean
}
```

### `ViewportRow`

A row of a window: its number on the sheet, from 1, its height in points, Excel's default (15) where it has none.

```ts
interface ViewportRow {
  row: number
  height: number
  hidden?: true
}
```

### `ViewportTable`

A table that meets a window ([`Viewport.tables`](#viewport)): its name, its range and which of its rows are a header and totals.

```ts
interface ViewportTable {
  name: string
  /** The whole table, its header and totals rows included (`TableInput.range`). */
  range: string
  /** Its first row is a header row (`TableInput.headers`, true by default). */
  headers: boolean
  /** Its last row is a totals row (`TableInput.totals`, false by default). */
  totals: boolean
}
```

### `WorkerLike`

What the pool needs of a Web Worker ([`CalcPoolOptions.spawn`](#calcpooloptions)).

```ts
interface WorkerLike {
  postMessage(message: unknown): void
  terminate(): void
  addEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { data?: unknown; message?: string }) => void): void
}
```

### `WriteXlsxOptions`

Options of [`writeXlsx`](#xlcalcxlsx-writexlsx): the language of its issues, which values of formulas the file stores, formulas or values alone.

```ts
interface WriteXlsxOptions {
  /** The language of the issues' messages, as `readXlsx` takes it (`ReadXlsxOptions.locale`). */
  locale?: string | Locale
  /**
   * The values of formulas the file stores. `require`: each formula's and each value of its array, from the result (an
   * `InputError` when one is missing: the result is not the book's). `omit`: none, and Excel computes them when it opens
   * the file. `preserve`: those the result has; a formula without one is reported as `noCachedValues`. By default
   * `require` with a result and `omit` without one, reported as `noCachedValues` when the book has formulas.
   */
  values?: 'require' | 'omit' | 'preserve'
  /**
   * `formulas` (the default): formulas, with their values by `values`. `values`: values only, a formula's cell gets its
   * value from the result (required then).
   */
  mode?: 'formulas' | 'values'
}
```

### `XlsxCached`

The values a file stores for formulas and for the cells their arrays spilled into ([`XlsxRead.cached`](#xlsxread)): what Excel
computed when it saved the file, by sheet in the book's order, in the shape of a result. `cached.sheets[0].cells.B1.value`
reads as `compute(book).sheets[0].cells.B1.value` does, so the two compare cell by cell, and `{ book, result: cached }`
writes the file again with Excel's values ([`writeXlsx`](#xlcalcxlsx-writexlsx)). `compute(read.book, { cached: read.cached })` (and
[`createSession`](#xlcalc-createsession)) open the book as Excel opens the file: the formulas a circular reference stops hold these values
([`CellInput.held`](#cellinput)).

```ts
interface XlsxCached {
  sheets: { name: string; cells: Record<string, XlsxCachedCell> }[]
}
```

### `XlsxCachedCell`

A stored value: a formula's, with `spill`, the range its array fills (an array formula of the old kind: its area), or
a cell of an array's range, with `spillOf`, the address of its formula; as in [`CellResult`](#cellresult).

```ts
type XlsxCachedCell = Pick<CellResult, 'value' | 'spill' | 'spillOf'>
```

### `XlsxIssue`

One thing lost or changed, and where: the sheet, and the cells (addresses or ranges) or the name it is about.

```ts
interface XlsxIssue {
  code: XlsxIssueCode
  /** What was lost or changed, and why. */
  message: string
  sheet?: string
  cells?: string[]
  name?: string
}
```

### `XlsxIssueCode`

What the book cannot hold of a file, or what writing changed. The codes:

- `richText`: text with formatted runs, read as its plain text.
- `name`: a name the book does not take (a built-in `_xlnm.` name, one on a sheet left out), a hidden name read as
  a visible one, or a name the writer left out.
- `table`: a table the model cannot hold (over another table or a merged area, its name taken), left out; a header
  cell that does not hold its column's name as text, given it; a structured reference without a table name.
- `externalLink`: a link that is no workbook's (DDE, OLE) or whose part is missing, kept as a book without sheets; a
  formula that refers to another workbook the links do not hold (`[3]Sheet!A1` in a file of two), or that the engine
  does not parse.
- `userFunction`: a call of a function the file marks as user-defined (`_xludf.`), a macro or add-in function.
- `formula`: a formula the engine does not parse, read as written (its cell computes to an error), or a formula
  the writer could not write (its value is written instead).
- `errorConstant`: an error typed into a cell as its value, read as a formula of that error.
- `errorValue`: an error the model has no code for (`#GETTING_DATA`, `#FIELD!`, `#BLOCKED!` and the like), read
  as `#ERROR!`; or written: `#CYCLE!` as 0, as Excel stores a circular reference.
- `dataTable`: a what-if data table (`TABLE()`), read as its stored values.
- `calculation`: calculation settings the book does not keep: the manual mode, automatic but for data tables,
  precision as displayed; an iteration count or change out of range, read as Excel's default.
- `style`: what a style cannot hold (pattern and gradient fills, diagonal borders, rotation and the like).
- `defaultFont`: a default font other than Calibri 11, which the model assumes.
- `defaultWidth`: a sheet's default column width other than Excel's standard of 8 digits, which the model assumes
  (Excel for Mac starts its own books at 10).
- `numberFormat`: a built-in East Asian date format the model does not have (50 to 58, read as its en-US code), or
  one this reader does not know (left out).
- `hidden`: hidden sheets, read as visible ones.
- `sheetKind`: a chart, dialog or macro sheet, left out.
- `cellMetadata`: a cell's rich value (`vm`: a data type, a picture in the cell), read as the value the cell stores.
- `autoFilter`: an autofilter's range, criteria and sort, a table's criteria, an advanced filter (the sheet's filter
  mode is read: [`SheetInput.filtered`](#sheetinput)); written: the filter mode of a sheet whose tables have no header row, left out.
- `dataValidation`: a rule the model cannot hold (no range of the sheet, no formula its type needs, an unknown type),
  left out; cells two rules share, kept by the earlier; an IME mode; a list whose items hold commas, read split at
  them; written: a rule whose formula the file cannot hold, left out.
- `outline`: automatic styles of an outline.
- `conditionalFormat`, `comment`, `hyperlink`, `drawing` (pictures, charts, shapes),
  `pivotTable`, `protection`, `frozenPanes`, `print`
  (headers and footers, page breaks, a page setup made by hand), `macros`, `sparkline`, `connection` (data
  connections and queries): parts of the file the model does not have, left
  out. View settings (zoom, gridlines, the selection) and a page setup Excel writes by default are not reported.
- `locale`: a locale the package does not have, left out of the book.
- `meta`: written: cell `meta` has no place in a file.
- `textLength`: written: text over Excel's 32,767 characters, cut, and a rule's texts over Excel's limits.
- `noCachedValues`: written: formulas without their values, which Excel computes when it opens the file: no result was
  given (`values: 'omit'` by default), or the result has no value for them (`values: 'preserve'`).
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

### `XlsxRead`

A file read into a book.

```ts
interface XlsxRead {
  /** The book, as the engine takes it. */
  book: Book
  /** The values the file stores for formulas and the cells their arrays spilled into, to compare with the engine's result. */
  cached: XlsxCached
  issues: XlsxIssue[]
}
```

### `XlsxSnapshot`

What [`writeXlsx`](#xlcalcxlsx-writexlsx) writes: a book and the values of its formulas, a session's `getSnapshot()`, `{ book, result:
compute(book) }`, or a file read with the values it stored, `{ book: read.book, result: read.cached }`.

```ts
interface XlsxSnapshot {
  book: BookInput
  result?: ComputeResult | XlsxCached
}
```

### `XlsxWritten`

What [`writeXlsx`](#xlcalcxlsx-writexlsx) answers: the file and what the book held that the file could not.

```ts
interface XlsxWritten {
  /** The xlsx file. */
  data: Uint8Array
  issues: XlsxIssue[]
}
```
