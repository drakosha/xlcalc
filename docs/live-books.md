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

`readCells` reads several ranges at one revision, deduplicating overlaps. Fields are `value`, `text`, `formula`, `style`, `numberFormat`, `note`, `link` and `issue`; value, text and formula are the default. Sheet and address are always returned. Input cells and spill cells are included, including unstored large spills, with output limited before they are materialized. Own styles and number formats are returned; use `getViewport` for resolved grid appearance.

`findCells` searches the text a cell shows (`in: 'text'`, the default), the text it opens with for editing (`edit`), its stored formula or its value, for a literal part of it by default, or by Excel's Find with the options of [Finding as Excel finds](#finding-as-excel-finds). Matching is case-insensitive unless `caseSensitive` is true. No regular expression or user code is executed. Searching happens where the book lives; the host need not keep a second copy of every shown value. Display filters and sorting remain application policies. Found cells come range by range, in the order the ranges are given, each range's cells by rows (by columns with `byColumns`), a cell once.

Pages contain `revision`, `clock`, `truncated` and, when another page exists, `cursor`. `maxCells` defaults to 1000 and is bounded by 10000. Use a cursor only with the same query; an edited session rejects it with `stale_revision`. `count: true` counts all matches and returns `total`, even when the first page is full. `readCells` without it stops when it has enough results to determine that another page exists; `findCells` reads every cell of its ranges each time, to give them in order. The limit counts cells, not serialized bytes; a host may also apply its own tool-output budget.

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
- An empty query finds nothing, and so does `~` alone.

`findCells` reads each cell of its ranges once, puts the cells in Find's order and tests them in turn, so every page costs a pass over the ranges' cells: on the synthetic book of 34,000 rows, a search of the data sheet's 510,000 cells with Find's options and `count` took 0.4 s to 0.5 s in a pool's thread and 0.4 s in a synchronous session, where the 26 windows of 20,000 cells with `editText` a grid reads to search the same cells itself took 1.1 s and 0.7 s (a load average of 6).

## Known differences

- An empty query: Excel's Find finds the empty cells of the range, or of the sheet from its first cell; `findCells` finds what cells hold and nothing for an empty query. A query of `~` alone, which reads as nothing, finds in Excel only the empty cell it starts after; nothing here.
- A range of one cell: Excel's `Range.Find` searches the whole sheet from the cell after it, as the dialog does with one cell selected; `findCells` searches the ranges it is given. A grid passes the sheet's whole range when one cell is selected.
- A text typed after an apostrophe is a text in the book, which keeps no mark of the apostrophe; Find reads neither, so this changes nothing found, but `editText` shows the apostrophe only before a text that would read as something else.

## Reverse dependencies

`getDependents(ranges, options?)` returns immediate formula readers and their matching source references. A reference to a large range remains a rectangle; no helper vertices or expanded million-cell lists escape the engine. The same page options apply. `maxReferences` defaults to 64 and is bounded by 1000; a truncated reference list is marked separately.

Use `explain` for the sources of a result, `getDependents` for possible consequences, and `previewChanges` for the actual value changes of a proposal. Dependencies are static: an IF branch can be a dependency even when the branch is not currently selected.

Asynchronous calls accept `signal` and `timeoutMs` beside their own options. Abort, deadline and stale-revision refusals leave the book and its history unchanged. Authorization, journals of agent actions and database version publication remain responsibilities of the host.

## How it was checked

Find was measured in Excel 16.0.20430 for Windows through its object model (`Range.Find` and `FindNext`, step `find` of `bench/windows`), in en-US and in ru-RU, on three sets held out with their predictions committed before Excel ran them (`bench/coverage/find/`, answers in `test/excel/win/find/`, test `test/find-win.test.ts`):

| set | cases in each locale | predicted by | confirmed | what it settled |
|---|---:|---|---:|---|
| `rules-held` | 155 | the grid's own model: the lookup functions' wildcards, Look in Formulas `editText`, Values the shown text | 136 | wildcards and `~`, `?` over a character outside the BMP, the case of letters, what Look in reads, merged areas and arrays, order, hidden cells on a filtered sheet |
| `rules-held-2` | 64 | `findCells` | 45 | hidden cells on a sheet without a filter, ranges of several areas, narrow columns, more of Look in, the Greek sigma |
| `rules-held-3` | 32 | `findCells` | 32 | a filter's mode makes Look in Formulas skip hidden cells and take the blocks, even when it hides nothing |

`findCells` gives Excel's cells in Excel's order in all of them but the seven cases of `rules-held` in each locale listed in [Known differences](#known-differences). Excel for Mac's Find was not measured. `test/session-find.test.ts` checks the options in a synchronous session, in the same thread and in a pool's thread, with closures and eval.
