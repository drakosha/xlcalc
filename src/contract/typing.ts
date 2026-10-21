/** Text typed into a session's cells as Excel takes it: into a cell, into several, pasted and entered into a range; the options, the refusals, the alerts of data validation, the widths taken, and the previews. */
import type { CellRef, CellValue, Refusal } from './values.js'
import type { CellInput } from './book.js'
import type { DataValidation } from './looks.js'
import type { EditOptions, EditResult } from './edits.js'

/** Options of `Session.setCellText` and `Session.setCellTexts`. */
export interface CellTextOptions extends EditOptions {
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

/** Why `Session.setCellText` rejected the text: a formula Excel would reject (`CheckCode`) or one that refers to a workbook not among the links (`external_ref`), part of an array formula, or a formula in a table's header cell. */
export type CellTextRefusal = 'syntax' | 'empty' | 'too_long' | 'too_deep' | 'external_ref' | 'array' | 'header'

/** A value that a cell's data validation rule does not allow (`Session.setCellText`): the rule, and its alert's style and message. */
export interface ValidationAlert {
  valid: false
  rule: DataValidation
  severity: 'stop' | 'warning' | 'information'
  message: string
}

/** A value rejected because a cell's data validation rule does not allow it (`Session.validateEntry`): the rule, and its alert's style and message. */
export type ValidationRefusal = Refusal<'validation'> & { rule: DataValidation; severity: ValidationAlert['severity'] }

/** Typed text rejected because a cell's data validation rule does not allow its value (`CellTextOptions.refuseInvalid`): the cell, the rule, and its alert's style and message. */
export type CellValidationRefusal = ValidationRefusal & CellRef

/** Text that `Session.setCellText` rejected, and the correction Excel offers for the formula, as in `parseEntry`'s `correction`. */
export type CellTextRefused = Refusal<CellTextRefusal> & { correction?: string }

/**
 * Text typed into a cell (`Session.setCellText`): the edit's result, with the width the column was widened to (`width`)
 * and the data validation rule the value violates (`validation`); or why the text was rejected, with the position in
 * the text, or, with `refuseInvalid`, the rule the value violates.
 */
export type CellTextResult = (EditResult & { ok: true; width?: number; validation?: ValidationAlert }) | CellTextRefused | CellValidationRefusal

/** What `Session.setCellText` would do (`Session.previewCellText`): the cell's input, its value, the width and the alert, without changing anything. */
export type CellTextPreview = { ok: true; input: CellInput | null; value: CellValue; width?: number; validation?: ValidationAlert } | CellTextRefused | CellValidationRefusal

/** Text typed into a cell, one of several typed at once (`Session.setCellTexts`). */
export interface CellTextEntry extends CellRef {
  text: string
}

/** The width a column was widened to (`Session.setCellTexts`): its sheet, its letters (`B`) and the width, as in `ColumnInput.width`. */
export interface ColumnWidth {
  sheet: string
  column: string
  width: number
}

/** A cell whose value its data validation rule does not allow (`Session.setCellTexts`), with the rule's alert. */
export type CellValidationAlert = CellRef & ValidationAlert

/**
 * Why `Session.setCellTexts` wrote none of the texts: the first one rejected, as `setCellText` rejects it, with its
 * cell; with `refuseInvalid`, the cell whose value violates the strictest data validation rule once all are written
 * (stop, then warning, then information; the first cell of that style).
 */
export type CellTextsRefusal = (CellTextRefused & CellRef) | CellValidationRefusal

/**
 * Texts typed into several cells at once (`Session.setCellTexts`): the edit's result, with the widths the columns were
 * widened to (`widths`) and the cells whose values their data validation rules do not allow (`validations`); or the
 * first rejected text, with its cell, and nothing written.
 */
export type CellTextsResult = (EditResult & { ok: true; widths?: ColumnWidth[]; validations?: CellValidationAlert[] }) | CellTextsRefusal

/** A cell of a preview of text written into cells (`Session.previewCellTexts`, `previewPasteText`, `previewFillCellText`): its input and its value as the call would leave them. */
export interface PreviewCell extends CellRef {
  input: CellInput | null
  value: CellValue
}

/** What `Session.setCellTexts` would do (`Session.previewCellTexts`): each cell's input and value, the widths and the alerts, without changing anything. */
export type CellTextsPreview = { ok: true; cells: PreviewCell[]; widths?: ColumnWidth[]; validations?: CellValidationAlert[] } | CellTextsRefusal

/**
 * Why `Session.pasteText` pasted nothing, with the cell: part of an array formula, a formula into a table's header cell
 * (as `setCellText` rejects them), or a rectangle that runs beyond the sheet's last row or column (`off_sheet`).
 */
export type PasteTextRefusal = Refusal<'array' | 'header' | 'off_sheet'> & CellRef

/** Text pasted from the clipboard (`Session.pasteText`): the edit's result with the range the clipboard's rectangle filled (`B1:D2`); or why nothing was pasted. */
export type PasteTextResult = (EditResult & { ok: true; range: string }) | PasteTextRefusal

/** What `Session.pasteText` would do (`Session.previewPasteText`): the range, and the input and value of each of its cells, without changing anything. */
export type PasteTextPreview = { ok: true; range: string; cells: PreviewCell[] } | PasteTextRefusal

/** Options of `Session.fillCellText`. */
export interface FillTextOptions extends CellTextOptions {
  /** The selection's active cell, which the text is typed into: a cell of the range (`B2`); its top-left cell by default. */
  active?: string
}

/**
 * Text entered into every cell of a range at once (`Session.fillCellText`, Excel's Control+Return): the edit's result,
 * with the widths the columns were widened to (`widths`) and the alert of the active cell's data validation rule
 * (`validation`); or why the text was rejected, with the cell, and nothing written.
 */
export type FillTextResult = (EditResult & { ok: true; widths?: ColumnWidth[]; validation?: ValidationAlert }) | CellTextsRefusal

/** What `Session.fillCellText` would do (`Session.previewFillCellText`): each cell's input and value, the widths and the alert, without changing anything. */
export type FillTextPreview = { ok: true; cells: PreviewCell[]; widths?: ColumnWidth[]; validation?: ValidationAlert } | CellTextsRefusal
