# Live workbook operations and bounded queries

These additions support a grid or an agent editing the same live session. They work in synchronous sessions, `createInlineCalc` and worker pools, with closures and eval. Existing calls keep their behavior.

## Mixed changes

`session.applyChanges(operations, options?)` accepts plain data, with operations applied in order and calculation performed over their final state. The call is atomic and produces one revision and one undo step. The result is an `EditResult` with the clock used.

```ts
const operations: SessionOperation[] = [
  { op: 'setCells', edits: [{ sheet: 'Plan', cell: 'B2', input: { formula: 'A2*Rate' } }] },
  { op: 'setNames', names: [{ name: 'Rate', ref: '1.2' }] },
  { op: 'patchCells', patches: [{ sheet: 'Plan', range: 'B2', set: { numberFormat: '0.00' } }] },
]
const preview = await session.previewChanges(operations, { baseRevision: session.revision })
const applied = await session.applyChanges(operations, {
  baseRevision: preview.revision,
  clock: preview.clock,
})
```

The operations are `setCells`, `replaceCells`, `clearCells`, `patchCells`, `setNames`, `setStyles` and `setLayout`. Their fields and replacement rules are those of the existing methods. In particular, `setNames` replaces the whole name list. Structure operations and AutoFit are separate calls. AutoFit reads calculated values; run it after the batch. For localized typed text and data validation alerts, use the existing text-entry methods. These changes add no new data validation policy.

Cell and look batches use sparse overlays while their input is checked and keep only touched inputs for rollback and history. Changed names, table headers, entered linked-book formulas, merges and hidden-row calculation require an aside-built state, as the existing operations do. No complete `BookSnapshot` is required. A style-only application without an explicit clock keeps the existing clock.

## Preview

`previewChanges` computes the complete effects, including dependent values and formula errors, then restores the book, clock, revision and history. It returns `changes`, `impacts` (`before`, `after`, and an error's `issue`), `totalChanges`, `truncated` and the clock used. `maxChanges` defaults to 1000 and is bounded by 10000; it limits output, not calculation.

Use both `preview.revision` and `preview.clock` when applying the operations. Revision checks prevent overwriting unseen edits. The clock also fixes NOW, TODAY, RAND and related functions in a live session. A later ordinary live edit resumes the session's normal clock behavior. The library does not hold a lock while the host reviews a preview.

A cell preview recalculates forward and back using a sparse value journal. A preview that needs rebuilding drops its temporary state afterward. Formula changes, layouts and clock-dependent values therefore have their full effects without making a durable fork. `meta` remains host data and does not appear in `CellChange`, as in existing edits.

## Reading and finding

```ts
const ranges = [{ sheet: 'Plan', range: 'A1:D1000' }]
const page = await session.readCells(ranges, {
  fields: ['value', 'formula'], maxCells: 100, baseRevision: session.revision,
})
const hits = await session.findCells(ranges, 'overdue', {
  in: 'text', fields: ['text', 'value'], maxCells: 100, count: true,
})
const totals = await session.summarizeRanges(ranges)
```

`readCells` reads several ranges at one revision, deduplicating overlaps: a cell comes in the first range that has it, and a sheet's cells are read once however many ranges it has. Fields are `value`, `text`, `formula`, `editText`, `style`, `numberFormat`, `note`, `link` and `issue`; value, text and formula are the default. `editText` is the text the cell opens with for editing, the one `getViewport` gives with its option `editText`. Sheet and address are always returned. Input cells and spill cells are included, including unstored large spills, with output limited before they are materialized. Own styles and number formats are returned; use `getViewport` for resolved grid appearance.

`rows` takes the rows a host shows, as `getViewport` takes them: numbers from 1, ascending, each once. `readCells` and `findCells` then give the cells of the ranges in those rows alone, and `total` counts those alone; a row outside a range leaves that range nothing of it. A grid with its own filter of rows searches what it shows in one call.

`findCells` searches the text a cell shows (`in: 'text'`, the default), the text it opens with for editing (`edit`), its stored formula or its value, for a literal part of it by default, or by Excel's Find with the options of [Finding as Excel finds](#finding-as-excel-finds). Matching is case-insensitive unless `caseSensitive` is true. No regular expression or user code is executed. Searching happens where the book lives; the host need not keep a second copy of every shown value. Sorting remains an application policy; a host's filter of rows goes in `rows`. Found cells come range by range, in the order the ranges are given, each range's cells by rows (by columns with `byColumns`), a cell once.

Pages contain `revision`, `clock`, `truncated` and, when another page exists, `cursor`. `maxCells` defaults to 1000 and is bounded by 10000. Use a cursor only with the same query; an edited session rejects it with `stale_revision`. A cursor holds the place of the page's last cell in the query's order, and the next page goes on after it without reading again what came before. `count: true` counts all matches and returns `total`, even when the first page is full; the cursor of a page that counted keeps the count, so a later page with `count` stops when it is full too. Without `count` a page stops when it has enough results to determine that another page exists. The limit counts cells, not serialized bytes; a host may also apply its own tool-output budget.

`summarizeRanges` returns cells, numbers, errors, a compensated sum, min and max with the revision and clock. Null and empty text do not count as cells; text and errors do not add to the numeric sum. An overflowing sum is `{ error: 'NUM' }`, not a non-finite number that JSON would turn into null.

## Finding as Excel finds

A grid's Find and Replace calls `findCells` with the options of Excel's dialog and gets the cells Excel's Find gives, in its order, without reading the sheet's windows itself.

```ts
const found = await session.findCells([{ sheet: 'Plan', range: 'A1:XFD1048576' }], 'Q?-*', {
  wildcards: true, in: 'edit', skipHidden: true, maxCells: 500, count: true, fields: ['text'],
})
```

- `wildcards: true` reads the query as Find does. `*` stands for any text and `?` for one character; `~` before a character takes it as it is, so `~*` finds a star, `~?` a question mark and `~~` a tilde. `~` before an ordinary character is that character (`~a` finds `a`), and a `~` at the end is left out (`a~` finds `a`). These are the lookup functions' wildcards, but `?` stands for one UTF-16 unit: a character outside the BMP takes two (`a??b` finds `a😀b`, `a?b` does not), and so does a letter with a combining mark (`e?x` finds `e` with a combining acute before `x`, `?x` finds `éx` written as one character).
- `wholeCell: true` is Match entire cell contents: `a*b` finds `axb` but not `xaxb`.
- Without `caseSensitive`, Match case off, a letter is found in either case: `hello` finds `HELLO`, `straße` finds `STRAẞE`, `σ` finds `ς` and `Σ`. Each character stays one letter: `ss` does not find `ß`, `i` finds `I` but neither `İ` nor `ı`, and `е` does not find `ё`.
- `in: 'edit'` is Look in Formulas: the text the cell opens with for editing (`editText`), in the book's locale. A formula is `=SUM(D1:D2)`, `=СУММ(D1:D2)` in ru-RU; 0.05 shown as `5%` is `5%`; a date is `3/15/2026` (`15.03.2026`); `$1,234.50` is `1234.5` (`1234,5`); TRUE is `TRUE` (`ИСТИНА`). A text is read as it is, without the apostrophe the formula bar shows before `'007`. Each cell of an old-kind array's area holds its formula; a cell an array spilled into holds none and is not found.
- `in: 'text'` is Look in Values: the text the cell shows, `$1,234.50`, `#N/A` (`#Н/Д`), `1 234 567` in ru-RU with its no-break space, `##` for a date in a column too narrow for it. A formula that shows an empty text is not found.
- `byColumns: true` is Search By Columns: each range's cells column by column. Several ranges come one after another, in the order given, as the areas of a selection do.
- `skipHidden: true` leaves out hidden cells as Find does. Look in Values never finds a cell of a hidden row or column. Look in Formulas finds them, unless the sheet is filtered (`SheetInput.filtered`: a filter has criteria, even ones that hide nothing). On a filtered sheet both leave out every hidden row and column, rows a filter hid and rows hidden by hand alike, and give the visible cells by the blocks the hidden lines leave: the blocks of the first rows first, left to right, each block by rows or by columns. With row 2 and column B hidden on a filtered sheet, A1, C1, A3 and C3 come in that order by rows and by columns alike.
- `after: { sheet, cell }` is Find from a cell, as Find Next goes on from the active cell (`Range.Find`'s `After`): the cells come in Find's order from the first after that cell, round to the ranges' first and on to the cell itself, which comes last. Pages go on through the round by their cursor, and `count` counts each cell once. A cell outside the ranges is refused, as Excel refuses it. On a filtered sheet a cell of a hidden row or column stands at its place in the range's first block: Find from it goes on with that block's cells after the place, then with the second block.
- An empty query finds nothing, and so does `~` alone.

### Texts kept for searching

The session keeps the texts it searches: the first search of a sheet reads the text of each cell it searches, in the kind asked (`text`, `edit`, `formula` or `value`), and keeps it, with the sheet's cells in order by rows (by columns once a search by columns needs them). Another search of the same sheet, the same query again or a longer one as a person types it, only matches the kept texts. A call that changes the book has the cells its answer lists read again at the next search, and nothing else: a typed value, a formula, a cell's format, the cells whose values follow. A change of a sheet's layout (widths, hidden lines, the formats of rows and columns, merged areas, tables, conditional formats) drops the sheet's texts, a change of the styles all of them, and a structure operation, new names, an undo of either and a call that stopped start from a book computed anew, which keeps none. What conditional formats show is read again after any change of a value, as their rules may read any cell. The texts change no answer: `test/find-texts.test.ts` compares every kind of search after each call of a series with a new session over the same book.

A sheet keeps two kinds at most, each in the cells' case and in any case, those searched last (Look in Formulas and Values for a grid). On the synthetic book of 34,000 rows, the data sheet's 506,668 cells take 3.9 MB for their places, Look in Formulas 7.0 MB and 3.9 MB more in any case, Look in Values 8.7 MB and 4.1 MB, and the order by columns 5.8 MB: about 28 MB for a grid's Find; 42 MB for 774,898 cells (52,000 rows).

| on the data sheet (510,000 cells), Node | 0.29.0 | now |
|---|---:|---:|
| the first search, Look in Formulas (`maxCells: 500`, `count`) | 0.41 s to 0.48 s | 0.44 s to 0.51 s |
| the same again, or a longer query | 0.37 s to 0.44 s | 0.02 s to 0.04 s |
| a search after an edit of a cell | 0.40 s to 0.44 s | 0.02 s to 0.04 s |
| the first search, Look in Values | 0.43 s to 0.45 s | 0.41 s to 0.45 s |
| the same again, or a longer query | 0.40 s to 0.45 s | 0.02 s to 0.04 s |
| the rows of a host's filter, 1,726 rows, counted | 19 pages of 10,000 the host sifts, 5.3 s to 6.0 s | one call: 0.09 s to 0.10 s the first, under 0.01 s after |
| `1` in all its pages of 10,000 | 19 pages, 5.3 s to 6.0 s | 0.16 s to 0.20 s |
| `*` in all its pages of 10,000 | 51 pages, 14 s | 0.4 s to 0.5 s |
| the sheet as 50 ranges of rows, counted | 3.8 s to 4.2 s | 0.03 s to 0.04 s |
| the cell after the 500th found (`after`, `maxCells: 1`) | | under 0.01 s |
| a grid's Replace All of 855 cells: pages, windows with `editText`, `setCellTexts` | the pages and windows alone 5.6 s to 6.3 s | 0.7 s to 0.85 s, as `replaceInCells` |

The times are of a synchronous session, of one in the same thread and of one in a pool's thread, which differ by less than the spread of runs, at a load average of 5 to 18.

## Replacing as Excel replaces

`replaceInCells(ranges, query, replacement, options?)` is Excel's Replace All: the cells `findCells` finds with the same options in the text they open with (Look in Formulas, which Excel's Replace always reads), each with the query's matches replaced, written in one call with one recalculation, one revision and one step of the undo. The answer is an `EditResult` with `replaced`, the number of cells written, and `refused`, the cells whose new text Excel would not take, left as they were.

```ts
const done = await session.replaceInCells([{ sheet: 'Plan', range: 'A1:XFD1048576' }], 'Q?-*', 'Q4-', {
  wildcards: true, skipHidden: true, rows: shownRows,
})
```

- The query reads as Find reads it (`wildcards`, `caseSensitive`, `wholeCell`), `?` one UTF-16 unit: `a??b` replaces `a😀b`, and `a?` replaces the first half of the emoji, leaving the second. `*` takes as little as it can and the query's last `*` the rest of the text: `a*b` in `a1b2b` is `a1b`, `b*b` in `b1b2b` is `b1b`, `*b` matches `a1b2b` twice, `a*` and `*` take the rest. Every match is replaced, in any case without `caseSensitive`. What is put goes in as it is, `*`, `?` and `~` too; nothing put removes the matches, and a cell whose whole text is replaced by nothing is emptied.
- Each cell found goes back as its text typed in the book's locale, changed or not: `a1` without `a` is the number 1, `#1+1` with `=` for `#` the formula `=1+1`, `=SUM(D1:D2)` (`=СУММ(D1:D2)` in ru-RU) with `MAX` for `SUM` the formula `=MAX(D1:D2)`, with `SUMX` a formula that answers `#NAME?`, a date with 2027 for 2026 the date. A percent cell has no percent editor: `5%` with 6 for 5 is 0.06, without `%` it is 5 (500%), and `-5` is -5. A cell of the text format takes a number, a formula, a boolean and a date as a General cell takes them typed, a date with its date format. An apostrophe put first makes a text.
- A formula that no longer reads (`=SUMD1:D2)`), or a formula into a table's header, leaves its cell as it was, in `refused` with the reason; the other cells are written. A text that reads as something else stays a text, as after an apostrophe (see [Known differences](#known-differences)).
- An array formula of the old kind is replaced whole from its first cell when any cell of its area is found, also from a range that holds part of it. A cell an array spilled into holds no formula and is not found.
- Hidden cells as Find skips them: hidden rows and columns of a sheet without a filter are replaced, those of a filtered sheet are not with `skipHidden`. `rows` takes the rows asked, as `findCells` does.
- No column widens: a number too wide for its column shows as Excel shows it there (`1.23E+11`). No rule of data validation is checked (not measured).
- Nothing found: nothing changes, and the revision and the undo stay.

## Where Control and an arrow key go

`getEnd(sheet, cell, direction)` is the cell Excel's Control and an arrow key come to from `cell`, as `Range.End` gives it: `direction` is `'up'`, `'down'`, `'left'` or `'right'`, the answer an address. A grid that knows only the cells of the windows it read moves its active cell there and reads the window around it.

```ts
// A1:A10 and B1:B500 hold values
await session.getEnd('S', 'A1', 'down') // 'A10', not B500's row
await session.getEnd('S', 'A10', 'down') // 'A1048576': nothing below
await session.getEnd('S', 'C1', 'left') // 'B1'
```

- A cell holds something when it has a value or a formula, whatever it shows: an empty text from a formula (`=""`), an empty text in the file or typed as an apostrophe alone, zero, a space, FALSE, an error, a value a format hides (`;;;`). So do the cells of a spill, an old-kind array's cells and the top-left cell of an array that cannot spill (`#SPILL!`). A cell with a look alone (a style or a number format, its own or of the sheet's looks), a note alone or a hyperlink alone is empty.
- From a cell that holds something, with the next cell in the direction holding something too, `End` goes to the last cell of that block. Otherwise it goes to the next cell that holds something, past any empty ones, and where there is none to the sheet's edge: row 1 or 1,048,576, column A or XFD. A cell at the edge stays. Blocks one empty cell apart are two blocks.
- Hidden rows and columns are not there for `End`, those hidden by hand, by a filter and by a height or a width of 0 alike: it never stops on a hidden cell, and a block goes on across hidden lines (in A1:A10 with A5 empty and row 5 hidden, A1 goes down to A10; with B12 holding a value in hidden row 12 and B15 a value, B11 goes down to B15). With every row from 20 on hidden, the edge going down is row 19. From a hidden cell `End` goes as from a shown one, by what that cell holds.
- A merged area is its cells: the value is in its top-left cell and the others are empty, so `End` stops on the top-left cell coming from either side and passes the area's other cells as empty ones (from G5 left over `B5:D5` it comes to B5; up from E12 under `D8:E10` it goes past the area).
- The locale changes nothing: ru-RU answered as en-US.

It reads the cells of the line it goes along, up to where the line's values end, and the sheet's hidden lines, which it reads once for each change of the layout. On the synthetic book of 34,000 rows a call takes 0.2 ms down a whole column of values and a few microseconds across a row, in a synchronous session; with every other row hidden, 17,000 hidden rows, 0.6 ms.

## Known differences

- An empty query: Excel's Find finds the empty cells of the range, or of the sheet from its first cell; `findCells` finds what cells hold and nothing for an empty query. A query of `~` alone, which reads as nothing, finds in Excel only the empty cell it starts after; nothing here.
- A range of one cell: Excel's `Range.Find` searches the whole sheet from the cell after it, as the dialog does with one cell selected; `findCells` searches the ranges it is given. A grid passes the sheet's whole range when one cell is selected.
- By columns on a filtered sheet, Excel's `FindNext` going round over the whole sheet passes the first column of the first block, and so did it over `B2:E11` (`rules-held-3`), not over `B2:E8` and `A4:E8` (`rules-held-5`): Find from a cell after those cells never gives them. `findCells` gives every cell once round.
- A text typed after an apostrophe is a text in the book, which keeps no mark of the apostrophe; Find reads neither, so this changes nothing found, but `editText` shows the apostrophe only before a text that would read as something else.
- Replace in a text that reads as a number and has no prefix character, as from a file: Excel enters it again as typed and it becomes a number (`007` with 8 for 7 is 8, and with 0 for 0 it is 7; `1,5` is 1.5 in ru-RU). The book keeps no mark of an apostrophe, so `replaceInCells` keeps such a text a text, as Excel does after an apostrophe typed.
- Replace over a range of one cell: Excel's `Range.Replace` replaces in the whole sheet, as Find searches it; `replaceInCells` replaces in the ranges it is given.

## Reverse dependencies

`getDependents(ranges, options?)` returns immediate formula readers and their matching source references. A reference to a large range remains a rectangle; no helper vertices or expanded million-cell lists escape the engine. The same page options apply. `maxReferences` defaults to 64 and is bounded by 1000; a truncated reference list is marked separately.

Use `explain` for the sources of a result, `getDependents` for possible consequences, and `previewChanges` for the actual value changes of a proposal. Dependencies are static: an IF branch can be a dependency even when the branch is not currently selected.

Asynchronous calls accept `signal` and `timeoutMs` beside their own options. Abort, deadline and stale-revision refusals leave the book and its history unchanged. Authorization, journals of agent actions and database version publication remain responsibilities of the host.

## Checked against Excel

Find and Control with an arrow were measured in Excel 16.0.20430 for Windows, en-US and ru-RU, on held-out sets predicted before Excel ran them: `findCells` gives Excel's cells in Excel's order in all but the known differences above, and `getEnd` gives Excel's cell in every case. Not measured: Excel for Mac's Find, and the keyboard on a screen.
