/** The edits of a session's book: cells set, replaced and patched, the layout changed, operations in a batch and their preview, what a call that changes the book answers and its subscribers get, undo and redo. */
import type { CellRef, CellValue, Clock, JsonValue, RangeRef } from './values.js'
import type { CellInput, CellLink, CellNote, ColumnInput, FrozenPanes, NameInput, NumberFormat, OutlineSettings, RowInput } from './book.js'
import type { CellStyle, ConditionalFormat, DataValidation, StylePatch } from './looks.js'
import type { CellDisplay, CellIssueDetail } from './results.js'
import type { ValidationRefusal } from './typing.js'

/** Fields of a cell to set (`Session.setCells`): a value sets the field, null removes it, a field left out stays. */
export type CellFields = { [K in keyof CellInput]?: CellInput[K] | null }

/**
 * A cell edit (`Session.setCells`): the fields given set the cell's, null removes one, and the fields left out stay, as
 * Excel keeps a cell's look when a value is typed over it. The content is one: a new `value` removes the formula, its
 * `array` and its `held`, a new `formula` the value and the `held` (a formula entered where a circular reference stops it
 * holds 0). `input: null` clears the content and the hyperlink and keeps the rest, as Excel's
 * Delete key.
 */
export interface CellEdit extends CellRef {
  input: CellFields | null
}

/** A cell replaced whole (`Session.replaceCells`): its new input, or null to remove the cell. */
export interface CellReplace extends CellRef {
  input: CellInput | null
}

/** A cell a call changed (`EditResult.changes`): its address, its new value and, where they changed, its text, formula, look, note or link. */
export interface CellChange extends CellRef, Partial<CellDisplay> {
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

/**
 * What a call that changes a session's book answers: the session's revision after it, the cells it changed, and what
 * else a host showing the book reads again.
 */
export interface EditResult {
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

/**
 * How a structure operation moved a book's addresses (`EditResult.moved`). `lines`: the lines from `at` (a row's number,
 * or a column's, A being 1) on move by `count`: down or right by the lines inserted before `at`, up or left when
 * `-count` lines from `at` on were deleted. `range`: the cells of `from` moved so that its top-left cell is at `to`,
 * what was there gone. A sheet renamed, deleted, inserted or moved: `index`, its place among the sheets, from 0.
 */
export type AddressMove =
  | { kind: 'lines'; sheet: string; axis: 'rows' | 'columns'; at: number; count: number }
  | { kind: 'range'; from: RangeRef; to: CellRef }
  | { kind: 'renameSheet'; sheet: string; name: string }
  | { kind: 'deleteSheet'; sheet: string }
  | { kind: 'insertSheet'; sheet: string; index: number }
  | { kind: 'moveSheet'; sheet: string; index: number }

/** The calls of a session that change its book, by name (`SessionChange.method`). */
export type ChangingMethod =
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

/**
 * A change of a session's book as its subscribers get it (`Session.subscribe`): the `EditResult` of the call that made
 * it (the revision after it, the cells it changed, the sheets whose layout changed, the style table, how addresses
 * moved, where to repaint) and the call's name.
 */
export interface SessionChange extends EditResult {
  method: ChangingMethod
}

/** Options of a call that changes a session's book. */
export interface EditOptions {
  /**
   * The revision the caller's view of the book is of (`EditResult.revision`, `Session.revision`): when the session's is
   * another, the call changes nothing and throws `CalcError` with code `stale_revision`, so that a host does not apply
   * an edit over one it has not seen.
   */
  baseRevision?: number
}

/** Serializable operations applied in order, with one final recalculation and one undo step. */
export type SessionOperation =
  | { op: 'setCells'; edits: CellEdit[] }
  | { op: 'replaceCells'; edits: CellReplace[] }
  | { op: 'clearCells'; ranges: RangeRef[]; what: ClearKind }
  | { op: 'patchCells'; patches: CellPatch[] }
  | { op: 'setNames'; names: NameInput[] }
  | { op: 'setStyles'; styles: CellStyle[] }
  | { op: 'setLayout'; edits: LayoutEdit[] }

/** An explicit clock makes a preview and its application reproducible, including in a live session. */
export interface ChangesOptions extends EditOptions { clock?: Clock }
/** Options of `previewChanges`: those of `applyChanges` and how many changed cells to give. */
export interface PreviewChangesOptions extends ChangesOptions {
  /** Maximum changed cells returned, 1000 by default, at most 10000. The computation itself remains complete. */
  maxChanges?: number
}
/** What `applyChanges` answers: the edit's result and the clock it computed with. */
export interface ChangesResult extends EditResult { clock: Clock }
/** A cell a previewed change would change (`ChangesPreview`): its value before and after. */
export interface CellImpact extends CellRef {
  before: CellValue
  after: CellValue
  /** The error after the operation, or the circular reference that stops the cell then, when there is one. */
  issue?: CellIssueDetail
}
/** A complete computation with bounded output; the session, its revision and history are unchanged. */
export interface ChangesPreview extends ChangesResult {
  impacts: CellImpact[]
  totalChanges: number
  truncated: boolean
}

/** What `Session.clearCells` clears: the content and hyperlinks (Excel's ClearContents), the look (ClearFormats), or both and the rest (Clear). */
export type ClearKind = 'contents' | 'formats' | 'all'

/** The calls a session can undo: those that change its book. `setClock` and a live session's new clock are not steps. */
export type UndoCall = 'applyChanges' | 'setCells' | 'replaceCells' | 'clearCells' | 'patchCells' | 'setCellText' | 'setCellTexts' | 'pasteText' | 'fillCellText' | 'replaceInCells' | 'setNames' | 'setStyles' | 'setLayout' | 'editStructure'

/** A step of a session undone or done again (`Session.undo`, `Session.redo`): the edit's result, as the call gave it. */
export interface UndoResult extends EditResult {
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

/** What a session can undo and redo now: the steps kept each way. */
export interface SessionHistory {
  undo: number
  redo: number
}

/** Fields to change on columns or rows: a value sets the field, null removes it, a field left out stays. */
export type LinePatch<T> = { [K in keyof T]?: T[K] | null }

/** A change of a sheet's layout. */
export type LayoutEdit =
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

/**
 * Excel's AutoFit of columns `C` or `B:D` (a double click on a column's border): each column with something to show gets
 * the width of its widest cell as Excel for Mac fits it (`fitColumns`), marked `bestFit` and shown if it was hidden; a
 * column with nothing to show keeps its width. The widths are computed from the values and layout the edits before it
 * leave (`docs/formats.md`, "Fitting columns").
 */
export interface ColumnsFit {
  sheet: string
  columns: string
  fit: true
}

/**
 * Excel's AutoFit of rows `5` or `2:9` (a double click on the border below a row's number): each row gets the height of
 * its highest cell as Excel for Mac fits it (`fitRows`), marked `bestFit` and shown if it was hidden; a row as high as
 * the default gets no height. The heights are computed from the values and layout the edits before it leave
 * (`docs/formats.md`, "Fitting rows").
 */
export interface RowsFit {
  sheet: string
  rows: string
  fit: true
}

/**
 * Whether a value meets a cell's rule (`Session.validateEntry`); when it does not, the rule, whose alert Excel shows,
 * its style (`errorStyle`, stop by default) and its message: the rule's `error`, else the text of Excel's alert in the
 * book's locale.
 */
export type EntryCheck = { ok: true } | ValidationRefusal

/**
 * A change of the look of a cell (`cell`) or a range (`range`: `B2:D9`, whole columns `B:D`, whole rows `2:5`)
 * (`Session.patchCells`): each field given sets the cells' own, null removes it, a field left out stays; values and
 * formulas stay as they are. A style by its fields (`StylePatch`) goes over the style each cell shows, its own or else
 * its row's or column's, as Excel formats a cell; a style by its number of the book's `styles` replaces it.
 */
export type CellPatch = { sheet: string; set: PatchFields } & ({ cell: string; range?: never } | { range: string; cell?: never })

/** The fields `CellPatch` sets; on whole rows and columns, `meta`, `note` and `link` can only be removed. */
export interface PatchFields {
  style?: number | StylePatch | null
  numberFormat?: NumberFormat | null
  meta?: JsonValue | null
  note?: CellNote | null
  link?: CellLink | null
}
