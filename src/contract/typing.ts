/** Text typed into a session's cells as Excel takes it: into a cell, into several, pasted and entered into a range; the options, the refusals, the alerts of data validation, the widths taken, and the previews. */
import type { CellRef, CellValue, Refusal } from './values.js'
import type { CellInput } from './book.js'
import type { DataValidation } from './looks.js'
import type { EditOptions, EditResult } from './edits.js'

/** `Session.setCellText`'s and `Session.setCellTexts`' options. */
export interface CellTextOptions extends EditOptions {
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

/** Why `Session.setCellText` did not take the text: a formula Excel does not take (`CheckCode`), part of an array formula, a formula in a table's header cell. */
export type CellTextRefusal = 'syntax' | 'empty' | 'too_long' | 'too_deep' | 'array' | 'header'

/** A value a cell's rule of data validation does not take (`Session.setCellText`): the rule, its alert's style and message. */
export interface ValidationAlert {
  valid: false
  rule: DataValidation
  severity: 'stop' | 'warning' | 'information'
  message: string
}

/** A value a cell's rule of data validation does not take, refused (`Session.validateEntry`): the rule, its alert's style and message. */
export type ValidationRefusal = Refusal<'validation'> & { rule: DataValidation; severity: ValidationAlert['severity'] }

/** Typed text refused for the value a cell's rule of data validation does not take (`CellTextOptions.refuseInvalid`): the cell, the rule, its alert's style and message. */
export type CellValidationRefusal = ValidationRefusal & CellRef

/** Text `Session.setCellText` did not take, and the formula Excel offers to correct it to, as `parseEntry`'s `correction`. */
export type CellTextRefused = Refusal<CellTextRefusal> & { correction?: string }

/**
 * Text typed into a cell (`Session.setCellText`): the edit's result, with the width the column took (`width`) and the
 * rule of data validation the value breaks (`validation`); or why the text was not taken, with its position in the text,
 * or, with `refuseInvalid`, the rule the value breaks.
 */
export type CellTextResult = (EditResult & { ok: true; width?: number; validation?: ValidationAlert }) | CellTextRefused | CellValidationRefusal

/** What `Session.setCellText` would do (`Session.previewCellText`): the cell's input, its value, the width and the alert, with nothing changed. */
export type CellTextPreview = { ok: true; input: CellInput | null; value: CellValue; width?: number; validation?: ValidationAlert } | CellTextRefused | CellValidationRefusal

/** Text typed into a cell, one of several typed at once (`Session.setCellTexts`). */
export interface CellTextEntry extends CellRef {
  text: string
}

/** The width a column took (`Session.setCellTexts`): its sheet, its letters (`B`) and the width, as `ColumnInput.width`. */
export interface ColumnWidth {
  sheet: string
  column: string
  width: number
}

/** A cell whose value its rule of data validation does not take (`Session.setCellTexts`), with the rule's alert. */
export type CellValidationAlert = CellRef & ValidationAlert

/**
 * Why `Session.setCellTexts` took none of the texts: the first one refused, as `setCellText` refuses it, with its cell;
 * with `refuseInvalid`, the cell whose value breaks the strictest rule of data validation once all are written (stop, then
 * warning, then information; the first cell of that style).
 */
export type CellTextsRefusal = (CellTextRefused & CellRef) | CellValidationRefusal

/**
 * Texts typed into cells at once (`Session.setCellTexts`): the edit's result, with the widths the columns took
 * (`widths`) and the cells whose values their rules of data validation do not take (`validations`); or the first text
 * refused, with its cell, and nothing written.
 */
export type CellTextsResult = (EditResult & { ok: true; widths?: ColumnWidth[]; validations?: CellValidationAlert[] }) | CellTextsRefusal

/** A cell of a preview of text written into cells (`Session.previewCellTexts`, `previewPasteText`, `previewFillCellText`): its input and its value as the call would leave them. */
export interface PreviewCell extends CellRef {
  input: CellInput | null
  value: CellValue
}

/** What `Session.setCellTexts` would do (`Session.previewCellTexts`): each cell's input and value, the widths and the alerts, with nothing changed. */
export type CellTextsPreview = { ok: true; cells: PreviewCell[]; widths?: ColumnWidth[]; validations?: CellValidationAlert[] } | CellTextsRefusal

/**
 * Why `Session.pasteText` pasted nothing, with the cell: part of an array formula, a formula into a table's header cell
 * (as `setCellText` refuses them), or a rectangle that runs past the sheet's last row or column (`off_sheet`).
 */
export type PasteTextRefusal = Refusal<'array' | 'header' | 'off_sheet'> & CellRef

/** Text pasted from the clipboard (`Session.pasteText`): the edit's result with the range the clipboard's rectangle filled (`B1:D2`); or why nothing was pasted. */
export type PasteTextResult = (EditResult & { ok: true; range: string }) | PasteTextRefusal

/** What `Session.pasteText` would do (`Session.previewPasteText`): the range and each of its cells' input and value, with nothing changed. */
export type PasteTextPreview = { ok: true; range: string; cells: PreviewCell[] } | PasteTextRefusal

/** `Session.fillCellText`'s options. */
export interface FillTextOptions extends CellTextOptions {
  /** The selection's active cell, which the text is typed into: a cell of the range (`B2`); its top-left cell by default. */
  active?: string
}

/**
 * A text entered into every cell of a range at once (`Session.fillCellText`, Excel's Control+Return): the edit's result,
 * with the widths the columns took (`widths`) and the alert of the active cell's rule of data validation
 * (`validation`); or why it was not taken, with the cell, and nothing written.
 */
export type FillTextResult = (EditResult & { ok: true; widths?: ColumnWidth[]; validation?: ValidationAlert }) | CellTextsRefusal

/** What `Session.fillCellText` would do (`Session.previewFillCellText`): each cell's input and value, the widths and the alert, with nothing changed. */
export type FillTextPreview = { ok: true; cells: PreviewCell[]; widths?: ColumnWidth[]; validation?: ValidationAlert } | CellTextsRefusal
