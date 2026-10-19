/** Sessions: their options, the synchronous `Session`, the live objects over one, and asynchronous computation in portions or in threads with its calculators. */
import type { CellValue, Clock, RangeRef, Scalar } from './values.js'
import type { Book, BookInput, BookLayout, CellInput, NameInput, NumberFormat } from './book.js'
import type { CellStyle, StylePatch } from './looks.js'
import type { BookSnapshot, CachedValues, CalcLimits, CellIssueDetail, ComputeOptions, ComputeResult, Execution, ExplainOptions, Explanation, Externals } from './results.js'
import type { CellEdit, CellPatch, CellReplace, ChangesOptions, ChangesPreview, ChangesResult, ClearKind, EditOptions, EditResult, EntryCheck, LayoutEdit, PreviewChangesOptions, SessionChange, SessionHistory, SessionOperation, UndoResult } from './edits.js'
import type { CellTextEntry, CellTextOptions, CellTextPreview, CellTextResult, CellTextsPreview, CellTextsResult, FillTextOptions, FillTextPreview, FillTextResult, PasteTextPreview, PasteTextResult } from './typing.js'
import type { SessionStructureResult, StructureOp } from './structure.js'
import type { CellValidation, CellView, CopyTextOptions, DependentsOptions, DependentsResult, EndDirection, FindCellsOptions, RangeSummary, ReadCellsOptions, ReadCellsResult, ReplaceOptions, ReplaceResult, SheetExtent, SummaryOptions, ValuesOptions, Viewport, ViewportOptions } from './reading.js'

/**
 * Session options. A session keeps one clock, `clock` or by default the moment of opening and a random
 * seed, until `setClock`. A live session (`live: true`) takes a fresh clock and a new seed on every
 * recomputation instead. Edits are compiled in the `execution` the session was opened with (`closures` by
 * default). With `display` (the default) results and changes carry what the cells show. `limits` bound every call;
 * a call over them throws `CalcError` and leaves the session as it was. `undo` bounds the steps `Session.undo` keeps.
 * `cached`, as `compute` takes it, gives the formulas a circular reference stops the values the book's file stores; the
 * session's book holds them as `held` from then on.
 */
export type SessionOptions = ({ clock?: Clock; live?: false } | { live: true }) & { externals?: Externals; execution?: Execution; display?: boolean; limits?: CalcLimits; undo?: UndoOptions; cached?: CachedValues }

/**
 * The steps a session keeps for `undo`: `depth` of them, 100 by default as in Excel (0 keeps none), holding `maxMb`
 * megabytes at most by the engine's estimate, 64 by default, in a pool with `maxHeapMb` its share of a thread's heap
 * (`maxHeapMb / (8 × maxSessionsPerThread)`, `docs/async.md`). Past either the oldest steps go; a step larger than
 * `maxMb` alone is not kept, and the steps before it go too, as Excel forgets what it cannot undo.
 */
export interface UndoOptions {
  depth?: number
  maxMb?: number
}

/**
 * A session over a book: edits recompute only the cells that depend on them. The current clock of the
 * session (`clock`) is what a saved result should record: a full computation with it gives the same hash.
 */
export interface Session {
  /** Mixed cells, names and looks applied atomically, with one final recalculation. */
  applyChanges(operations: SessionOperation[], opts?: ChangesOptions): ChangesResult
  /** Dependent-value changes and errors without changing the session; use its clock when applying. */
  previewChanges(operations: SessionOperation[], opts?: PreviewChangesOptions): ChangesPreview
  /** One revision, multiple ranges, only the fields and number of cells requested; range by range, each range's cells by rows. */
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
  summarizeRanges(ranges: RangeRef[], opts?: SummaryOptions): RangeSummary
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
  /**
   * The values of a rectangle of a sheet (`B2:D9`, `B:B`, `2:9`), as rows from top to bottom, each from left to right:
   * every cell's value as `readCells` gives it in its field `value` (a number to 15 significant digits, a text, a boolean,
   * an error as `{ error }`), null where a cell holds none. The rectangle is the one asked, its empty cells too, not cut
   * to where the sheet's cells end; one of more than `maxCells` cells is refused (`CalcError` `limit`). Each call gives
   * arrays of its own. It costs the rectangle's cells, not a result object and an address for each; changes nothing.
   */
  getValues(sheet: string, range: string, opts?: ValuesOptions): CellValue[][]
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

// ---------- live objects over a session (`docs/live-objects.md`) ----------

/**
 * Objects of a synchronous session's sheets and cells (`liveWorkbook`): a cell reads what the session computed, an
 * assignment to it is the session's edit (recomputed, one step of the undo), and a listener learns of the changes of a
 * cell, a sheet or the book, whoever makes them. One per session: `liveWorkbook` gives the same object for it again.
 */
export interface LiveWorkbook {
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

/** A sheet of a `LiveWorkbook`. After the sheet is deleted, every use of the object throws `InputError`. */
export interface LiveSheet {
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

/**
 * A cell of a `LiveSheet`, by its address: after rows are inserted above it, it is the cell now at that address. Each
 * property reads the session as it is now; an assignment edits it at once, or as part of a `batch`, and throws what the
 * session's call throws, the book unchanged.
 */
export interface LiveCell {
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

// ---------- asynchronous and threaded computation (`docs/async.md`) ----------

/** What a computation needs of an AbortSignal: the platform's `AbortSignal` is one. */
export interface AbortSignalLike {
  readonly aborted: boolean
  readonly reason?: unknown
  addEventListener(type: 'abort', listener: () => void, options?: { once?: boolean }): void
  removeEventListener(type: 'abort', listener: () => void): void
}

/**
 * Why an asynchronous call failed (`CalcError.code`). `soft_timeout`: its `timeoutMs` passed, seen by a check inside
 * the computation. `aborted`: its `signal` fired. `hard_timeout`: it ran past the pool's `hardTimeoutMs` and its
 * thread was terminated. `out_of_memory`: its thread reached `maxHeapMb`. `thread_failed`: its thread exited or threw
 * outside the computation. `restore_failed`: the session could not be restored on a thread and is closed. `limit`:
 * the book or the edit is over `CalcLimits`. `queue_full`: `maxQueue` tasks already wait. `too_many_sessions`: the pool holds
 * `maxSessions` sessions. `closed`: the calculator or the session is closed. `stale_revision`: an edit's `baseRevision`
 * is not the session's revision (`EditOptions`); it changed nothing.
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
  /** Stops the call: a waiting call at once, a running one at its next check. */
  signal?: AbortSignalLike
  /** A soft deadline: milliseconds of computing, counted from when the call starts computing (not while it waits). */
  timeoutMs?: number
}

/** Options of `compute` of a pool or of `createInlineCalc`: those of `compute` and the call's deadline and signal. */
export type AsyncComputeOptions = ComputeOptions & CalcOptions
/** `signal` and `timeoutMs` bound the opening (the first computation of the book). */
export type AsyncSessionOptions = SessionOptions & CalcOptions

/** `AsyncSession.editStructure`: the same answer as `Session.editStructure`'s. */
export type AsyncStructureResult = SessionStructureResult

/** The options of an asynchronous call that changes a session's book. */
export type AsyncEditOptions = CalcOptions & EditOptions

/**
 * A session computed in portions or in a thread: the methods of `Session`, as promises. Calls are taken in order,
 * each after the one before has finished. A call that fails (a deadline, an abort, a limit) changes nothing.
 */
export interface AsyncSession {
  /** Mixed cells, names and looks applied atomically, with one final recalculation. */
  applyChanges(operations: SessionOperation[], opts?: ChangesOptions & CalcOptions): Promise<ChangesResult>
  /** Dependent-value changes and errors without changing the session; use its clock when applying. */
  previewChanges(operations: SessionOperation[], opts?: PreviewChangesOptions & CalcOptions): Promise<ChangesPreview>
  /** One revision, multiple ranges, only the fields and number of cells requested; range by range, each range's cells by rows. */
  readCells(ranges: RangeRef[], opts?: ReadCellsOptions & CalcOptions): Promise<ReadCellsResult>
  /** Search performed where the book lives, literal or by Excel's Find; output is paged, in Find's order. */
  findCells(ranges: RangeRef[], query: string, opts?: FindCellsOptions & CalcOptions): Promise<ReadCellsResult>
  /** Excel's Replace All, as `Session.replaceInCells`. */
  replaceInCells(ranges: RangeRef[], query: string, replacement: string, opts?: ReplaceOptions & CalcOptions): Promise<ReplaceResult>
  /** Immediate formula dependents; source ranges stay ranges. */
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
  /** As `Session.getCells`; from a thread only the rectangles come back, in one call; `signal` and `timeoutMs` stop it as they stop a calculation. */
  getCells(sheet: string, range: string | string[], opts?: CalcOptions): Promise<Record<string, CellView>>
  /** As `Session.copyText`; the text is prepared in one read in the session's thread. */
  copyText(sheet: string, range: string, opts?: CopyTextOptions & CalcOptions): Promise<string>
  /** As `Session.getValues`; from a thread the rows come back in one call. */
  getValues(sheet: string, range: string, opts?: ValuesOptions & CalcOptions): Promise<CellValue[][]>
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

/** A calculator: `createInlineCalc` (the same thread, in portions) or `createCalcPool` (threads). */
export interface Calc {
  compute(book: BookInput, opts?: AsyncComputeOptions): Promise<ComputeResult>
  createSession(book: BookInput, opts?: AsyncSessionOptions): Promise<AsyncSession>
  /** Refuses further calls (`closed`), the waiting and running ones too, and closes the sessions. */
  close(): Promise<void>
}

/** Options of `createInlineCalc`. */
export interface InlineCalcOptions {
  /** Milliseconds of computing between two yields to the event loop; 10 by default. */
  sliceMs?: number
  /** The soft deadline of every call that gives none. */
  timeoutMs?: number
  /** The limits of every computation; a computation's own `limits` take their place field by field, the others stay. */
  limits?: CalcLimits
}

/** What the pool needs of a Web Worker (`CalcPoolOptions.spawn`). */
export interface WorkerLike {
  postMessage(message: unknown): void
  terminate(): void
  addEventListener(type: 'message' | 'error' | 'messageerror', listener: (event: { data?: unknown; message?: string }) => void): void
}

/** Options of `createCalcPool`. */
export interface CalcPoolOptions {
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
