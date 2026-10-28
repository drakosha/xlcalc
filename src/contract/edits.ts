/** The edits of a session's book: cells set, replaced and patched, the layout changed, operations in a batch and their preview, what a call that changes the book answers and its subscribers get, undo and redo. */
import type { CellRef, CellValue, Clock, JsonValue, RangeRef } from './values.js'
import type { AutoFilter, CellInput, CellLink, CellNote, ColumnInput, FrozenPanes, NameInput, NumberFormat, OutlineSettings, RowInput, SheetProtection, SheetView } from './book.js'
import type { CellStyle, Color, ConditionalFormat, DataValidation, StylePatch } from './looks.js'
import type { CellDisplay, CellIssueDetail } from './results.js'
import type { ValidationRefusal } from './typing.js'

/** Fields of a cell to set (`Session.setCells`): a value sets the field, null removes it, and an omitted field is left unchanged. */
export type CellFields = { [K in keyof CellInput]?: CellInput[K] | null }

/**
 * A cell edit (`Session.setCells`): the given fields replace the cell's, null removes a field, and omitted fields are
 * left unchanged, as Excel keeps a cell's formatting when a value is typed over it. A cell has one kind of content: a
 * new `value` removes the formula, its `array` and its `held`, and a new `formula` removes the value and the `held` (a
 * formula entered where a circular reference prevents its calculation holds 0). `input: null` clears the content and
 * the hyperlink and keeps the rest, like Excel's Delete key.
 */
export interface CellEdit extends CellRef {
  input: CellFields | null
}

/** A cell replaced as a whole (`Session.replaceCells`): its new input, or null to remove the cell. */
export interface CellReplace extends CellRef {
  input: CellInput | null
}

/** A cell that a call changed (`EditResult.changes`): its address, its new value and, when they changed, its text, formula, formatting, note or hyperlink. */
export interface CellChange extends CellRef, Partial<CellDisplay> {
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

/**
 * What a call that changes a session's workbook returns: the session's revision after the call, the cells it changed,
 * and anything else a host displaying the workbook needs to read again.
 */
export interface EditResult {
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

/** Rows given one height by a call (`EditResult.fitted`). */
export interface FittedRows {
  sheet: string
  rows: string
  height: number | null
}

/**
 * How a structure operation moved a workbook's addresses (`EditResult.moved`). `lines`: the rows or columns from `at`
 * onward (a row number, or a column number with A as 1) move by `count`: down or right when lines were inserted before
 * `at`, up or left when `-count` lines starting at `at` were deleted. `range`: the cells of `from` moved so that its
 * top-left cell is at `to`, replacing what was there. A sheet renamed, deleted, inserted or moved: `index` is its
 * position among the sheets, from 0.
 */
export type AddressMove =
  | { kind: 'lines'; sheet: string; axis: 'rows' | 'columns'; at: number; count: number }
  | { kind: 'range'; from: RangeRef; to: CellRef }
  | { kind: 'renameSheet'; sheet: string; name: string }
  | { kind: 'deleteSheet'; sheet: string }
  | { kind: 'insertSheet'; sheet: string; index: number }
  | { kind: 'moveSheet'; sheet: string; index: number }

/** The calls of a session that change its workbook, by name (`SessionChange.method`). */
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
  | 'setExternalBooks'
  | 'setClock'
  | 'editStructure'
  | 'setStyles'
  | 'setLayout'
  | 'undo'
  | 'redo'

/**
 * A change to a session's workbook as its subscribers receive it (`Session.subscribe`): the `EditResult` of the call
 * that made it (the revision after the call, the cells it changed, the sheets whose layout changed, the style table,
 * how addresses moved, where to repaint) and the call's name.
 */
export interface SessionChange extends EditResult {
  method: ChangingMethod
}

/** Options of a call that changes a session's workbook. */
export interface EditOptions {
  /**
   * The revision of the workbook that the caller's view reflects (`EditResult.revision`, `Session.revision`): when the
   * session's revision differs, the call changes nothing and throws `CalcError` with code `stale_revision`, so that a
   * host never applies an edit on top of one it has not seen.
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
/** Options of `previewChanges`: those of `applyChanges`, and how many changed cells to return. */
export interface PreviewChangesOptions extends ChangesOptions {
  /** Maximum number of changed cells returned, 1000 by default, at most 10000. The calculation itself is still complete. */
  maxChanges?: number
}
/** What `applyChanges` returns: the edit's result and the clock it calculated with. */
export interface ChangesResult extends EditResult { clock: Clock }
/** A cell that a previewed change would change (`ChangesPreview`): its value before and after. */
export interface CellImpact extends CellRef {
  before: CellValue
  after: CellValue
  /** The error after the operation, or the circular reference that then prevents the cell's calculation, if any. */
  issue?: CellIssueDetail
}
/** A complete calculation with bounded output; the session, its revision and its history are unchanged. */
export interface ChangesPreview extends ChangesResult {
  impacts: CellImpact[]
  totalChanges: number
  truncated: boolean
}

/** What `Session.clearCells` clears: the content and hyperlinks (Excel's ClearContents), the formatting (ClearFormats), or both and everything else (Clear). */
export type ClearKind = 'contents' | 'formats' | 'all'

/** The calls a session can undo: those that change its workbook. `setClock` and a live session's new clock are not steps. */
export type UndoCall = 'applyChanges' | 'setCells' | 'replaceCells' | 'clearCells' | 'patchCells' | 'setCellText' | 'setCellTexts' | 'pasteText' | 'fillCellText' | 'replaceInCells' | 'setNames' | 'setExternalBooks' | 'setStyles' | 'setLayout' | 'editStructure'

/** A session step that was undone or redone (`Session.undo`, `Session.redo`): the edit's result, as the call returned it. */
export interface UndoResult extends EditResult {
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

/** What a session can undo and redo now: the number of steps kept in each direction. */
export interface SessionHistory {
  undo: number
  redo: number
}

/** Fields to change on columns or rows: a value sets the field, null removes it, and an omitted field is left unchanged. */
export type LinePatch<T> = { [K in keyof T]?: T[K] | null }

/** A change of a sheet's layout. */
export type LayoutEdit =
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
  /** The sheet's saved view (`SheetInput.view`), all its fields at once; null removes it. */
  | { sheet: string; view: SheetView | null }
  /** The color of the sheet's tab (`SheetInput.tabColor`); null removes it. */
  | { sheet: string; tabColor: Color | null }
  /** All the conditional formats of the sheet (`SheetInput.conditionalFormats`), in priority order. */
  | { sheet: string; conditionalFormats: ConditionalFormat[] }
  /** All the data validation rules of the sheet (`SheetInput.validations`). */
  | { sheet: string; validations: DataValidation[] }
  /** Where the sheet's summary rows and columns are (`SheetInput.outline`); null for Excel's default, below and to the right. */
  | { sheet: string; outline: OutlineSettings | null }

/**
 * Excel's AutoFit of columns `C` or `B:D` (a double-click on a column's border): each column with content to display
 * gets the width of its widest cell as Excel for Mac fits it (`fitColumns`), is marked `bestFit` and is unhidden if it
 * was hidden; a column with nothing to display keeps its width. The widths are calculated from the values and layout
 * left by the preceding edits (`docs/formats.md`, "Fitting columns").
 */
export interface ColumnsFit {
  sheet: string
  columns: string
  fit: true
}

/**
 * Excel's AutoFit of rows `5` or `2:9` (a double-click on the border below a row's number): each row gets the height of
 * its tallest cell as Excel for Mac fits it (`fitRows`), is marked `bestFit` and is unhidden if it was hidden; a row
 * whose fitted height equals the default gets no height. The heights are calculated from the values and layout left by
 * the preceding edits (`docs/formats.md`, "Fitting rows").
 */
export interface RowsFit {
  sheet: string
  rows: string
  fit: true
}

/**
 * Whether a value satisfies a cell's rule (`Session.validateEntry`); when it does not, the rule whose alert Excel
 * displays, its style (`errorStyle`, stop by default) and its message: the rule's `error`, or else the text of Excel's
 * alert in the workbook's locale.
 */
export type EntryCheck = { ok: true } | ValidationRefusal

/**
 * A formatting change to a cell (`cell`) or a range (`range`: `B2:D9`, entire columns `B:D`, entire rows `2:5`)
 * (`Session.patchCells`): each given field sets the cells' own, null removes it, and an omitted field is left
 * unchanged; values and formulas stay as they are. A style given by its fields (`StylePatch`) is applied over the style
 * each cell uses, its own or else its row's or column's, as Excel formats a cell; a style given by its index in the
 * workbook's `styles` replaces it.
 */
export type CellPatch = { sheet: string; set: PatchFields } & ({ cell: string; range?: never } | { range: string; cell?: never })

/** The fields `CellPatch` sets; on entire rows and columns, `meta`, `note` and `link` can only be removed. */
export interface PatchFields {
  style?: number | StylePatch | null
  numberFormat?: NumberFormat | null
  meta?: JsonValue | null
  note?: CellNote | null
  link?: CellLink | null
}
