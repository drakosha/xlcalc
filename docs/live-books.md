# Batch changes, previews and paged queries in a live session

A grid or an agent editing a live session applies a batch of mixed changes as one revision and one undo step, previews changes with all their effects before applying them, reads and finds cells a page at a time (including Excel's Find and Replace, and the cell that Ctrl and an arrow key go to) and asks which formulas read a range. Each call works in synchronous sessions, `createInlineCalc` and worker pools, with closures and eval.

## Mixed changes

`session.applyChanges(operations, options?)` accepts plain data, with operations applied in order and calculation performed over their final state. The call is atomic and produces one revision and one undo step. The result is an `EditResult` with the clock used.

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'Plan', cells: { A2: { value: 10 }, C2: { formula: 'B2+1' } } }] })
const operations = [
  { op: 'setCells', edits: [{ sheet: 'Plan', cell: 'B2', input: { formula: 'A2*Rate' } }] },
  { op: 'setNames', names: [{ name: 'Rate', ref: '1.2' }] },
  { op: 'patchCells', patches: [{ sheet: 'Plan', range: 'B2', set: { numberFormat: '0.00' } }] },
]
const preview = await session.previewChanges(operations, { baseRevision: session.revision })
console.log(preview.impacts.map((i) => [i.cell, i.before, i.after]), session.revision)
// → [ [ 'B2', null, 12 ], [ 'C2', 1, 13 ] ] 0
const applied = await session.applyChanges(operations, { baseRevision: preview.revision, clock: preview.clock })
console.log(applied.revision, session.getCell('Plan', 'B2').text, session.history)
// → 1 12.00 { undo: 1, redo: 0 }
```

The operations are `setCells`, `replaceCells`, `clearCells`, `patchCells`, `setNames`, `setStyles` and `setLayout`. Their fields and replacement rules are those of the existing methods. In particular, `setNames` replaces the whole name list. Structure operations and AutoFit are separate calls. AutoFit reads calculated values; run it after the batch. For localized typed text and data validation alerts, use the existing text-entry methods. These changes add no new data validation policy.

Cell and formatting batches use sparse overlays while their input is checked, and keep only the inputs they touched for rollback and history. Changed names, table headers, entered formulas that link to other workbooks, merges and calculation over hidden rows require a state built on the side, as the existing operations do. No complete `BookSnapshot` is required. A formatting-only change applied without an explicit clock keeps the existing clock.

## Preview

`previewChanges` calculates the complete effects, including dependent values and formula errors, then restores the workbook, clock, revision and history. It returns `changes`, `impacts` (`before`, `after`, and an error's `issue`), `totalChanges`, `truncated` and the clock used. `maxChanges` defaults to 1000 and is bounded by 10000; it limits output, not calculation.

Use both `preview.revision` and `preview.clock` when applying the operations. Revision checks prevent overwriting unseen edits. The clock also fixes NOW, TODAY, RAND and related functions in a live session. A later ordinary live edit resumes the session's normal clock behavior. The library does not hold a lock while the host reviews a preview.

A cell preview recalculates forward and back using a sparse value journal. A preview that needs rebuilding drops its temporary state afterward. Formula changes, layouts and clock-dependent values therefore have their full effects without making a durable fork. `meta` remains host data and does not appear in `CellChange`, as in existing edits.

## Reading and finding

```js
import { createSession } from 'xlcalc'

// A1:A300 hold numbers, every 7th row the text 'overdue'
const cells = Object.fromEntries(Array.from({ length: 300 }, (_, i) => [`A${i + 1}`, { value: i % 7 ? i : 'overdue' }]))
const session = createSession({ sheets: [{ name: 'Plan', cells }] }, { clock: { now: '2026-10-16T12:00:00Z', seed: 1 } })
const ranges = [{ sheet: 'Plan', range: 'A1:D1000' }]
const page = await session.readCells(ranges, { fields: ['value', 'formula'], maxCells: 100, baseRevision: session.revision })
console.log(page.cells.length, page.cells[0], page.truncated, typeof page.cursor)
// → 100 { sheet: 'Plan', cell: 'A1', value: 'overdue' } true string
const hits = await session.findCells(ranges, 'overdue', { in: 'text', fields: ['text', 'value'], maxCells: 10, count: true })
console.log(hits.cells.length, hits.total, hits.cells[1])
// → 10 43 { sheet: 'Plan', cell: 'A8', text: 'overdue', value: 'overdue' }
const { cells: n, numbers, sum, min, max } = await session.summarizeRanges(ranges)
console.log({ cells: n, numbers, sum, min, max })
// → { cells: 300, numbers: 257, sum: 38529, min: 1, max: 299 }
```

`readCells` reads several ranges at one revision, deduplicating overlaps: a cell is returned in the first range that contains it, and a sheet's cells are read once however many ranges cover it. The cells come range by range, each range by rows (`A1`, `B1`, `A2`), whatever order the workbook's cells were given in. They are found through an index of each sheet's cells by position, the same one Find uses, built when a sheet is first read and kept up to date by every cell write, so a page reads only its own cells. A viewport of 1,600 cells on a sheet of 506,594 takes about 0.9 ms with the default fields and 0.2 ms with values only, the whole sheet in pages of 10,000 about 160 ms, and the first read of the sheet about 100 ms.

Fields are `value`, `text`, `formula`, `editText`, `style`, `numberFormat`, `note`, `link` and `issue`; value, text and formula are the default. A read without `text`, `editText` and `issue` takes only each cell's value and input, and is the fastest. `editText` is the text the cell opens with in edit mode, the same one `getViewport` returns with its option `editText`. Sheet and address are always returned. Input cells and spill cells are included, including unstored large spills, with output limited before they are materialized. Own styles and number formats are returned; use `getViewport` for the resolved grid appearance.

`rows` accepts the rows a host displays, as `getViewport` does: numbers from 1, ascending, each once. `readCells` and `findCells` then return only the cells of the ranges in those rows, and `total` counts only those; a range with none of those rows contributes nothing. A grid with its own row filter searches what it displays in one call.

`findCells` searches the text a cell displays (`in: 'text'`, the default), the text it opens with in edit mode (`edit`), its stored formula or its value, by default for a literal substring, or like Excel's Find with the options of [Finding as Excel finds](#finding-as-excel-finds). Matching is case-insensitive unless `caseSensitive` is true. No regular expression or user code is executed. Searching happens where the workbook lives; the host does not need a second copy of every displayed value. Sorting remains an application policy; a host's row filter goes in `rows`. Found cells come range by range, in the order the ranges are given, each range's cells by rows (by columns with `byColumns`), each cell once.

Pages contain `revision`, `clock`, `truncated` and, when another page exists, `cursor`. `maxCells` defaults to 1000 and is bounded by 10000. Use a cursor only with the same query; an edited session rejects it with `stale_revision`. A cursor holds the position of the page's last cell in the query's order, and the next page continues after it without reading again what came before. `count: true` counts all matches and returns `total`, even when the first page is full; the cursor of a page that counted keeps the count, so a later page with `count` also stops when it is full. Without `count`, a page stops when it has enough results to determine that another page exists. The limit counts cells, not serialized bytes; a host may also apply its own tool-output budget.

`summarizeRanges` returns cells, numbers, errors, a compensated sum, min and max, with the revision and clock. Null and empty text do not count as cells; text and errors do not add to the numeric sum. An overflowing sum is `{ error: 'NUM' }`, not a non-finite number that JSON would turn into null. `rows` accepts the rows a host's filter displays, as for `readCells`. What Excel's status bar counts has not been verified: here a formula that displays empty text is not counted as a cell, and an error is counted as a cell and adds to `errors`.

A host that keeps its own texts can find and replace in them as `findCells` and `replaceInCells` do, with `findMatcher(query, opts?)` and `findReplacer(query, by, opts?)` from `xlcalc/text` (`caseSensitive`, `wildcards`, `wholeCell`), without the engine. `compareValues(a, b, { descending })` is the order in which Excel's Sort command sorts values of mixed types, for a host's display filter: numbers, text in any case, FALSE then TRUE, errors, and blanks last in either direction. The Sort command treats all errors as equal, so a stable sort keeps them in their original order; the SORT function orders them by the number ERROR.TYPE returns. To sort many values, `valueComparator({ descending })` validates the options once and returns the comparison function itself, which accepts values without checking them: `values.sort(valueComparator())`. Both also accept a `locale`, as a tag or a locale object, and order text as workbooks in that locale do; ru-RU's order is the same as en-US's, Latin before Cyrillic, as Excel for Windows sorts in every locale (Excel for Mac in ru-RU puts Cyrillic first; see [Locales](locale.md)).

### Values of a rectangle

```js
import { createSession } from 'xlcalc'

const session = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { formula: 'A1/3' }, A2: { value: 'x' }, B2: { formula: '1/0' } } }] }, { clock: { now: '2026-10-16T12:00:00Z', seed: 1 } })
console.log(session.getValues('S', 'A1:C2'))
// → [ [ 1, 0.333333333333333, null ], [ 'x', { error: 'DIV0' }, null ] ]
console.log(session.getValues('S', 'A:A', { maxCells: 1048576 }).length)
// → 1048576
```

`getValues(sheet, range)` returns a rectangle's values as rows, top to bottom, each from left to right, in the form a chart, an export or another engine expects: every value as `readCells` returns it in its `value` field, and null where a cell has none. The rectangle is exactly the one requested, including its empty cells, not trimmed to where the sheet's data ends. An entire column is a million rows, so a rectangle of more than `maxCells` cells (a million by default) is rejected with `CalcError` `limit` rather than truncated. Each call returns new arrays, which a host may modify. No object and no address is created per cell: the 506,594 values of a sheet of 34,001 rows are returned in about 10 ms, against about 155 ms for `readCells` by pages, and a viewport of 1,600 cells in about 0.04 ms.

## Finding as Excel finds

A grid's Find and Replace calls `findCells` with the options of Excel's dialog and gets the cells Excel's Find returns, in the same order, without reading the sheet's viewports itself.

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{
    name: 'Plan',
    rows: { 3: { hidden: true } },
    filtered: true,
    cells: { A1: { value: 'Q1-plan' }, B1: { value: 'Q2-fact' }, A2: { value: 'q3-plan' }, A3: { value: 'Q4-hidden' }, A4: { value: 'Q12' }, B4: { formula: '"Q9-"&A4' } },
  }],
})
const all = [{ sheet: 'Plan', range: 'A1:XFD1048576' }]
const found = await session.findCells(all, 'Q?-*', { wildcards: true, in: 'edit', skipHidden: true, maxCells: 500, count: true, fields: ['text'] })
// in any case; A3 left out, its row hidden on a filtered sheet; B4 found in its formula, ="Q9-"&A4
console.log(found.cells.map((c) => `${c.cell} ${c.text}`), found.total)
// → [ 'A1 Q1-plan', 'B1 Q2-fact', 'A2 q3-plan', 'B4 Q9-Q12' ] 4
```

- `wildcards: true` interprets the query as Find does. `*` stands for any text and `?` for one character; `~` before a character takes that character literally, so `~*` finds an asterisk, `~?` a question mark and `~~` a tilde. `~` before an ordinary character is that character (`~a` finds `a`), and a trailing `~` is ignored (`a~` finds `a`). These are the lookup functions' wildcards, except that `?` stands for one UTF-16 code unit: a character outside the BMP takes two (`a??b` finds `a😀b`, `a?b` does not), and so does a letter with a combining mark (`e?x` finds `e` followed by a combining acute and `x`, and `?x` finds `éx` written as a single character).
- `wholeCell: true` is Match entire cell contents: `a*b` finds `axb` but not `xaxb`.
- Without `caseSensitive` (Match case off), a letter matches in either case: `hello` finds `HELLO`, `straße` finds `STRAẞE`, and `σ` finds `ς` and `Σ`. Each character still matches as one letter: `ss` does not find `ß`, `i` finds `I` but neither `İ` nor `ı`, and `е` does not find `ё`.
- `in: 'edit'` is Look in: Formulas, the text the cell opens with in edit mode (`editText`), in the workbook's locale. A formula is `=SUM(D1:D2)`, or `=СУММ(D1:D2)` in ru-RU; 0.05 displayed as `5%` is `5%`; a date is `3/15/2026` (`15.03.2026`); `$1,234.50` is `1234.5` (`1234,5`); TRUE is `TRUE` (`ИСТИНА`). Text is matched as it is, without the apostrophe the formula bar shows before `'007`. Each cell in the range of a legacy array formula contains its formula; a cell an array spilled into contains none and is not found.
- `in: 'text'` is Look in: Values, the text the cell displays: `$1,234.50`, `#N/A` (`#Н/Д`), `1 234 567` in ru-RU with its no-break space, `##` for a date in a column that is too narrow for it. A formula that displays empty text is not found.
- `byColumns: true` is Search: By Columns, each range's cells column by column. Several ranges are searched one after another, in the order given, like the areas of a multi-area selection.
- `skipHidden: true` skips hidden cells as Find does. Look in: Values never finds a cell in a hidden row or column. Look in: Formulas finds them, unless the sheet is filtered (`SheetInput.filtered`: a filter has criteria, even ones that hide nothing). On a filtered sheet both skip every hidden row and column, whether hidden by a filter or manually, and return the visible cells by the blocks that the hidden rows and columns leave: the blocks of the first rows first, left to right, each block by rows or by columns. With row 2 and column B hidden on a filtered sheet, A1, C1, A3 and C3 come in that order, by rows and by columns alike.
- `after: { sheet, cell }` starts after a given cell, as Find Next continues from the active cell (the `After` argument of `Range.Find`): the cells come in Find's order starting with the first one after that cell, wrapping around to the beginning of the ranges and on to the cell itself, which comes last. Pages continue around the wrap by their cursor, and `count` counts each cell once. A cell outside the ranges is rejected, as Excel rejects it. On a filtered sheet, a cell in a hidden row or column is placed at its position in the range's first block: Find from it continues with that block's cells after the position, then with the second block.
- An empty query finds nothing, and so does `~` alone.

### Cached search text

The session caches the texts it searches: the first search of a sheet reads the text of each cell it searches, in the requested kind (`text`, `edit`, `formula` or `value`), and keeps it, along with the sheet's cells in row order (and column order once a search by columns needs it). Another search of the same sheet, whether the same query again or a longer one as the user types, only matches against the cached texts. After a call that changes the workbook, the next search rereads only the cells listed in the call's result: a typed value, a formula, a cell's number format, the cells whose values depend on it. A change of a sheet's layout (widths, hidden rows and columns, row and column formats, merged areas, tables, conditional formats) drops that sheet's texts, and a change of the styles drops all of them. A structure operation, new names, the undo of either, and a call that was stopped all start from a freshly calculated workbook, which has no cached texts. What conditional formats display is reread after any change of a value, as their rules may depend on any cell. The cache never changes a result: `test/find-texts.test.ts` compares every kind of search after each call of a series with a new session over the same workbook.

A sheet keeps at most two kinds of text, each both as is and case-folded, the two searched most recently (Look in: Formulas and Values for a grid). On the synthetic workbook of 34,000 rows, the data sheet's 506,668 cells take 3.9 MB for their positions, Look in: Formulas 7.0 MB plus 3.9 MB case-folded, Look in: Values 8.7 MB plus 4.1 MB, and the column order 5.8 MB: about 28 MB for a grid's Find, and 42 MB for 774,898 cells (52,000 rows).

| on the data sheet (510,000 cells), Node | 0.29.0 | now |
|---|---:|---:|
| the first search, Look in: Formulas (`maxCells: 500`, `count`) | 0.41 s to 0.48 s | 0.44 s to 0.51 s |
| the same query again, or a longer one | 0.37 s to 0.44 s | 0.02 s to 0.04 s |
| a search after editing a cell | 0.40 s to 0.44 s | 0.02 s to 0.04 s |
| the first search, Look in: Values | 0.43 s to 0.45 s | 0.41 s to 0.45 s |
| the same query again, or a longer one | 0.40 s to 0.45 s | 0.02 s to 0.04 s |
| the rows of a host's filter, 1,726 rows, counted | 19 pages of 10,000 filtered by the host, 5.3 s to 6.0 s | one call: 0.09 s to 0.10 s the first time, under 0.01 s afterwards |
| `1`, all pages of 10,000 | 19 pages, 5.3 s to 6.0 s | 0.16 s to 0.20 s |
| `*`, all pages of 10,000 | 51 pages, 14 s | 0.4 s to 0.5 s |
| the sheet as 50 ranges of rows, counted | 3.8 s to 4.2 s | 0.03 s to 0.04 s |
| the cell after the 500th match (`after`, `maxCells: 1`) | | under 0.01 s |
| a grid's Replace All of 855 cells: pages, viewports with `editText`, `setCellTexts` | the pages and viewports alone 5.6 s to 6.3 s | 0.7 s to 0.85 s, with `replaceInCells` |

The times for a synchronous session, a session on the same thread and a session on a pool's thread differ by less than the spread between runs, at a load average of 5 to 18.

## Replacing as Excel replaces

`replaceInCells(ranges, query, replacement, options?)` is Excel's Replace All: it takes the cells `findCells` finds with the same options in the text they open with in edit mode (Look in: Formulas, which Excel's Replace always uses), replaces the query's matches in each, and writes them in one call with one recalculation, one revision and one undo step. The result is an `EditResult` with `replaced`, the number of cells written, and `refused`, the cells whose new text Excel would not accept, which are left unchanged.

```js
import { createSession } from 'xlcalc'

const session = createSession({
  sheets: [{ name: 'Plan', cells: { A1: { value: 'Q1-plan' }, A2: { value: 'Q2-fact' }, A3: { value: 'Q1' }, A4: { formula: 'A1&"!"' } } }],
})
const done = await session.replaceInCells([{ sheet: 'Plan', range: 'A1:XFD1048576' }], 'Q?-*', 'Q4-', { wildcards: true })
console.log(done.replaced, ['A1', 'A2', 'A3', 'A4'].map((c) => session.getCell('Plan', c).value), session.history)
// → 2 [ 'Q4-', 'Q4-', 'Q1', 'Q4-!' ] { undo: 1, redo: 0 }
```

- The query is interpreted as Find interprets it (`wildcards`, `caseSensitive`, `wholeCell`), with `?` as one UTF-16 code unit: `a??b` replaces `a😀b`, and `a?` replaces the first half of the emoji, leaving the second. `*` matches as little as possible, except the query's last `*`, which matches the rest of the text: `a*b` in `a1b2b` matches `a1b`, `b*b` in `b1b2b` matches `b1b`, `*b` matches `a1b2b` twice, and `a*` and `*` match the rest. Every match is replaced, in any case without `caseSensitive`. The replacement is inserted literally, including `*`, `?` and `~`; an empty replacement removes the matches, and a cell whose entire text is replaced by nothing is emptied.
- Each found cell is written back as its text typed in the workbook's locale, whether it changed or not: `a1` with `a` removed becomes the number 1; `#1+1` with `=` for `#` becomes the formula `=1+1`; `=SUM(D1:D2)` (`=СУММ(D1:D2)` in ru-RU) with `MAX` for `SUM` becomes the formula `=MAX(D1:D2)`, and with `SUMX` a formula that returns `#NAME?`; a date with 2027 for 2026 stays a date. A percentage cell has no percentage entry: `5%` with 6 for 5 becomes 0.06, without `%` it becomes 5 (500%), and `-5` becomes -5. A cell formatted as Text accepts a number, a formula, a boolean and a date as a General cell accepts them when typed, and a date with its date format. A leading apostrophe in the replacement makes the result text.
- A formula that no longer parses (`=SUMD1:D2)`), or a formula in a table's header, leaves its cell unchanged and goes to `refused` with the reason; the other cells are written. Text that would be parsed as something else stays text, as after an apostrophe (see [Known differences](#known-differences)).
- A legacy array formula is replaced as a whole, from its first cell, when any cell of its range is found, even from a search range that covers only part of it. A cell an array spilled into contains no formula and is not found.
- Hidden cells are handled as Find skips them: hidden rows and columns of a sheet without a filter are replaced, and those of a filtered sheet are not, with `skipHidden`. `rows` accepts the requested rows, as for `findCells`.
- No column widens: a number too wide for its column is displayed as Excel displays it there (`1.23E+11`). No data validation rule is checked (not verified in Excel).
- When nothing is found, nothing changes, and the revision and the undo history stay as they were.

## Where Control and an arrow key go

`getEnd(sheet, cell, direction)` returns the cell that Excel's Ctrl and an arrow key move to from `cell`, as `Range.End` returns it: `direction` is `'up'`, `'down'`, `'left'` or `'right'`, and the result is an address. A grid that knows only the cells of the viewports it has read moves its active cell there and reads the viewport around it.

```js
import { createSession } from 'xlcalc'

// A1:A10 and B1:B500 hold values
const cells = {}
for (let r = 1; r <= 500; r++) cells[`B${r}`] = { value: r }
for (let r = 1; r <= 10; r++) cells[`A${r}`] = { value: r }
const session = createSession({ sheets: [{ name: 'S', cells }] })
console.log(await session.getEnd('S', 'A1', 'down'), await session.getEnd('S', 'A10', 'down'), await session.getEnd('S', 'C1', 'left'))
// → A10 A1048576 B1
```

- A cell counts as non-empty when it has a value or a formula, whatever it displays: empty text from a formula (`=""`), empty text from a file or typed as a lone apostrophe, zero, a space, FALSE, an error, or a value hidden by its format (`;;;`). The cells of a spill, the cells of a legacy array formula and the top-left cell of an array that cannot spill (`#SPILL!`) are non-empty too. A cell with only formatting (a style or a number format, its own or from the sheet's `looks`), only a note or only a hyperlink is empty.
- From a non-empty cell whose next cell in the direction is non-empty too, `End` goes to the last cell of that block. Otherwise it goes to the next non-empty cell, skipping empty ones, and if there is none, to the edge of the sheet: row 1 or 1,048,576, column A or XFD. A cell at the edge stays where it is. Blocks separated by one empty cell are two blocks.
- `End` ignores hidden rows and columns, whether hidden manually, by a filter or by a height or width of 0: it never stops on a hidden cell, and a block continues across hidden rows and columns (in A1:A10 with A5 empty and row 5 hidden, A1 goes down to A10; with a value in B12 in hidden row 12 and a value in B15, B11 goes down to B15). With every row from 20 onward hidden, the edge going down is row 19. From a hidden cell, `End` moves as from a visible one, according to what that cell contains.
- A merged area counts as its individual cells: the value is in its top-left cell and the others are empty, so `End` stops on the top-left cell when coming from either side and passes the area's other cells as empty (from G5 going left over `B5:D5` it stops at B5; going up from E12 under `D8:E10` it passes the area).
- The locale makes no difference: ru-RU behaves as en-US.

It reads the cells of the row or column it moves along, up to where its values end, and the sheet's hidden rows and columns, which it reads once per layout change. On the synthetic workbook of 34,000 rows, a call takes 0.2 ms down an entire column of values and a few microseconds across a row, in a synchronous session; with every other row hidden (17,000 hidden rows), 0.6 ms.

## Known differences

- An empty query: Excel's Find finds the empty cells of the range, or of the sheet from its first cell; `findCells` finds only cells with content, and nothing for an empty query. A query of `~` alone, which matches nothing, finds in Excel only the empty cell it starts after; here it finds nothing.
- A range of one cell: Excel's `Range.Find` searches the whole sheet from the next cell, as the dialog does with one cell selected; `findCells` searches only the ranges it is given. A grid passes the sheet's entire range when one cell is selected.
- Searching by columns on a filtered sheet, Excel's `FindNext`, wrapping around the whole sheet, skips the first column of the first block, and did so over `B2:E11` too (`rules-held-3`), but not over `B2:E8` and `A4:E8` (`rules-held-5`): Find from a cell after those cells never returns them. `findCells` returns every cell once per wrap.
- Text typed after an apostrophe is stored as text in the workbook, which keeps no record of the apostrophe; Find reads neither, so this does not change what is found, but `editText` shows the apostrophe only before text that would otherwise be parsed as something else.
- Replace in text that looks like a number and has no prefix character, as when it comes from a file: Excel re-enters it as typed and it becomes a number (`007` with 8 for 7 becomes 8, and with 0 for 0 it becomes 7; `1,5` becomes 1.5 in ru-RU). The workbook keeps no record of an apostrophe, so `replaceInCells` keeps such text as text, as Excel does after a typed apostrophe.
- Replace over a range of one cell: Excel's `Range.Replace` replaces in the whole sheet, as Find searches it; `replaceInCells` replaces only in the ranges it is given.

## Reverse dependencies

`getDependents(ranges, options?)` returns the formulas that directly read the ranges, and their matching source references. A reference to a large range remains a rectangle; no helper vertices or expanded million-cell lists escape the engine. The same page options apply. `maxReferences` defaults to 64 and is bounded by 1000; a truncated reference list is marked separately.

Use `explain` for the sources of a result, `getDependents` for possible consequences, and `previewChanges` for the actual value changes of a proposal. Dependencies are static: an IF branch can be a dependency even when the branch is not currently selected.

Asynchronous calls accept `signal` and `timeoutMs` besides their own options. Rejections for an abort, a deadline or a stale revision leave the workbook and its history unchanged. Authorization, journals of agent actions and publishing database versions remain the host's responsibility.

## Checked against Excel

Find and Ctrl with an arrow key were tested in Excel 16.0.20430 for Windows, en-US and ru-RU, on held-out sets with the results predicted before Excel ran them: `findCells` returns Excel's cells in Excel's order in every case except the known differences above, and `getEnd` returns Excel's cell in every case. The order in which the Sort command sorts values of mixed types, which `compareValues` implements, was tested in Excel for Mac and for Windows on a held-out set: all errors are equal and keep their order in either direction. Not tested: Find in Excel for Mac, and keyboard navigation on screen.
