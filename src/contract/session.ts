/** Sessions: their options, the synchronous `Session`, the live objects over one, and asynchronous computation in portions or in threads with its calculators. */
import type { CellValue, Clock, RangeRef, Scalar } from './values.js'
import type { Book, BookInput, BookLayout, CellInput, ExternalBook, NameInput, NumberFormat } from './book.js'
import type { CellStyle, StylePatch } from './looks.js'
import type { BookSnapshot, CachedValues, CalcLimits, CellIssueDetail, ComputeOptions, ComputeResult, Execution, ExplainOptions, Explanation, Externals } from './results.js'
import type { CellEdit, CellPatch, CellReplace, ChangesOptions, ChangesPreview, ChangesResult, ClearKind, EditOptions, EditResult, EntryCheck, LayoutEdit, PreviewChangesOptions, SessionChange, SessionHistory, SessionOperation, UndoResult } from './edits.js'
import type { CellTextEntry, CellTextOptions, CellTextPreview, CellTextResult, CellTextsPreview, CellTextsResult, FillTextOptions, FillTextPreview, FillTextResult, PasteTextPreview, PasteTextResult } from './typing.js'
import type { SessionStructureResult, StructureOp } from './structure.js'
import type { AutoSumGuess, CellValidation, CellView, CopyTextOptions, DependentsOptions, DependentsResult, EndDirection, FindCellsOptions, RangeSummary, ReadCellsOptions, ReadCellsResult, ReplaceOptions, ReplaceResult, SheetExtent, SummaryOptions, ValuesOptions, Viewport, ViewportOptions } from './reading.js'

/**
 * Session options. A session keeps a single clock, `clock` or by default the time it was opened and a random
 * seed, until `setClock`. A live session (`live: true`) instead takes a fresh clock and a new seed on every
 * recalculation. Edits are compiled with the `execution` the session was opened with (`closures` by
 * default). With `display` (the default), results and changes include what the cells display. `limits` apply to every
 * call; a call that exceeds them throws `CalcError` and leaves the session unchanged. `undo` limits the steps
 * `Session.undo` keeps. `cached`, as `compute` accepts it, gives the formulas stopped by a circular reference the values
 * stored in the workbook's file; from then on, the session's workbook holds them as `held`.
 */
export type SessionOptions = ({ clock?: Clock; live?: false } | { live: true }) & { externals?: Externals; execution?: Execution; display?: boolean; limits?: CalcLimits; undo?: UndoOptions; cached?: CachedValues }

/**
 * The steps a session keeps for `undo`: up to `depth` of them, 100 by default as in Excel (0 keeps none), taking at most
 * `maxMb` megabytes by the engine's estimate, 64 by default, or in a pool with `maxHeapMb`, the session's share of a
 * thread's heap (`maxHeapMb / (8 × maxSessionsPerThread)`, `docs/async.md`). Beyond either limit, the oldest steps are
 * dropped; a step larger than `maxMb` on its own is not kept, and the steps before it are dropped too, just as Excel
 * clears what it cannot undo.
 */
export interface UndoOptions {
  depth?: number
  maxMb?: number
}

/**
 * A session over a workbook: edits recalculate only the cells that depend on them. The session's current clock
 * (`clock`) is what a saved result should record: a full calculation with it produces the same hash.
 */
export interface Session {
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
   * session and later calls change them, so do not modify it; serialize or copy whatever must be preserved.
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

// ---------- live objects over a session (`docs/live-objects.md`) ----------

/**
 * Objects for a synchronous session's sheets and cells (`liveBook`): a cell reads what the session calculated, an
 * assignment to it is an edit of the session (recalculated, one undo step), and a listener is notified of changes to a
 * cell, a sheet or the workbook, whoever makes them. One per session: `liveBook` returns the same object for the same
 * session.
 */
export interface LiveBook {
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

/** A sheet of a `LiveBook`. After the sheet is deleted, every use of the object throws `InputError`. */
export interface LiveSheet {
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

/**
 * A cell of a `LiveSheet`, identified by its address: after rows are inserted above it, it refers to the cell now at
 * that address. Each property reads the session's current state; an assignment edits the session immediately, or as
 * part of a `batch`, and throws whatever the session's call throws, leaving the workbook unchanged.
 */
export interface LiveCell {
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

// ---------- asynchronous and threaded computation (`docs/async.md`) ----------

/** What a calculation needs from an AbortSignal; the platform's `AbortSignal` qualifies. */
export interface AbortSignalLike {
  readonly aborted: boolean
  readonly reason?: unknown
  addEventListener(type: 'abort', listener: () => void, options?: { once?: boolean }): void
  removeEventListener(type: 'abort', listener: () => void): void
}

/**
 * Why a call failed (`CalcError.code`): an asynchronous call rejects with it, and a synchronous session throws `limit`,
 * `closed` and `stale_revision`. `soft_timeout`: its `timeoutMs` elapsed, as detected by a check
 * during the calculation. `aborted`: its `signal` fired. `hard_timeout`: it ran past the pool's `hardTimeoutMs` and its
 * thread was terminated. `out_of_memory`: its thread reached `maxHeapMb`. `thread_failed`: its thread exited or threw
 * outside the calculation. `restore_failed`: the session could not be restored on a thread and has been closed. `limit`:
 * the workbook or the edit exceeds `CalcLimits`. `queue_full`: `maxQueue` tasks are already waiting. `too_many_sessions`:
 * the pool already holds `maxSessions` sessions. `closed`: the calculator or the session is closed. `stale_revision`: an
 * edit's `baseRevision` does not match the session's revision (`EditOptions`); nothing was changed.
 */
export type CalcErrorCode =
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

/** Options of one asynchronous call. */
export interface CalcOptions {
  /** Stops the call: a waiting call immediately, a running one at its next check. */
  signal?: AbortSignalLike
  /** A soft deadline: milliseconds of calculation, counted from when the call starts calculating (not while it waits). */
  timeoutMs?: number
}

/** Options of `compute` on a pool or on `createInlineCalc`: those of `compute`, plus the call's deadline and signal. */
export type AsyncComputeOptions = ComputeOptions & CalcOptions
/** `signal` and `timeoutMs` apply to opening the session (the workbook's first calculation). */
export type AsyncSessionOptions = SessionOptions & CalcOptions

/** `AsyncSession.editStructure`: the same result as `Session.editStructure` returns. */
export type AsyncStructureResult = SessionStructureResult

/** Options of an asynchronous call that changes a session's workbook. */
export type AsyncEditOptions = CalcOptions & EditOptions

/**
 * A session calculated in slices or in a thread: the methods of `Session`, returning promises. Calls are processed in
 * order, each after the previous one has finished. A call that fails (a deadline, an abort, a limit) changes nothing.
 */
export interface AsyncSession {
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
  /** The session's workbook; from a thread, a copy. */
  getBook(): Promise<Book>
  /** Like `Session.getLayout`; from a thread, a copy, without the cells. */
  getLayout(): Promise<BookLayout>
  getUsedRange(sheet: string): Promise<SheetExtent>
  getEnd(sheet: string, cell: string, direction: EndDirection): Promise<string>
  getAutoSum(sheet: string, cell: string): Promise<AutoSumGuess>
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

/** A calculator: `createInlineCalc` (the same thread, in slices) or `createCalcPool` (threads). */
export interface Calc {
  compute(book: BookInput, opts?: AsyncComputeOptions): Promise<ComputeResult>
  createSession(book: BookInput, opts?: AsyncSessionOptions): Promise<AsyncSession>
  /** Rejects further calls (`closed`), including waiting and running ones, and closes the sessions. */
  close(): Promise<void>
}

/** Options of `createInlineCalc`. */
export interface InlineCalcOptions {
  /** Milliseconds of calculation between two yields to the event loop; 10 by default. */
  sliceMs?: number
  /** The soft deadline for every call that does not set its own. */
  timeoutMs?: number
  /** The limits for every calculation; a calculation's own `limits` override them field by field, and the other fields still apply. */
  limits?: CalcLimits
}

/** What the pool needs from a Web Worker (`CalcPoolOptions.spawn`). */
export interface WorkerLike {
  postMessage(message: unknown): void
  terminate(): void
  addEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { data?: unknown; message?: string }) => void): void
}

/** Options of `createCalcPool`. */
export interface CalcPoolOptions {
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
