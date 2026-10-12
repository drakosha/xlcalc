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

Cell and look batches use sparse overlays while their input is checked and keep only touched inputs for rollback and history. Changed names, table headers, merges and hidden-row calculation require an aside-built state, as the existing operations do. No complete `BookSnapshot` is required. A style-only application without an explicit clock keeps the existing clock.

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

`findCells` searches a literal substring of text, formula or value. Text is the default and matching is case-insensitive unless `caseSensitive` is true. No regular expression or user code is executed. Searching happens where the book lives; the host need not keep a second copy of every shown value. Display filters and sorting remain application policies.

Pages contain `revision`, `clock`, `truncated` and, when another page exists, `cursor`. `maxCells` defaults to 1000 and is bounded by 10000. Use a cursor only with the same query; an edited session rejects it with `stale_revision`. `count: true` counts all matches and returns `total`, even when the first page is full. Without it, the scan stops when it has enough results to determine that another page exists. The limit counts cells, not serialized bytes; a host may also apply its own tool-output budget.

`summarizeRanges` returns cells, numbers, errors, a compensated sum, min and max with the revision and clock. Null and empty text do not count as cells; text and errors do not add to the numeric sum. An overflowing sum is `{ error: 'NUM' }`, not a non-finite number that JSON would turn into null.

## Reverse dependencies

`getDependents(ranges, options?)` returns immediate formula readers and their matching source references. A reference to a large range remains a rectangle; no helper vertices or expanded million-cell lists escape the engine. The same page options apply. `maxReferences` defaults to 64 and is bounded by 1000; a truncated reference list is marked separately.

Use `explain` for the sources of a result, `getDependents` for possible consequences, and `previewChanges` for the actual value changes of a proposal. Dependencies are static: an IF branch can be a dependency even when the branch is not currently selected.

Asynchronous calls accept `signal` and `timeoutMs` beside their own options. Abort, deadline and stale-revision refusals leave the book and its history unchanged. Authorization, journals of agent actions and database version publication remain responsibilities of the host.
