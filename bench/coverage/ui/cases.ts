/**
 * What Excel does with what a person does in its window, for Excel (`run.ts ui <set> [ru|<tag>]`, `bench/windows/run.ts
 * ui <set> [ru|<tag>]`): each case is a book of one sheet `S` and one action, on the Mac with Excel in front. `bars`
 * selects cells in turn and reads the formula bar, then enters each cell's editing (F2) and reads what it opens with;
 * `paste` puts text on the clipboard (tabs between columns, line breaks between rows) and pastes it into a cell (Cmd+V);
 * `fill` selects a range, types a text and enters it into every cell of the selection (Control+Return). Then the cells the
 * action touched are read (what each holds, shows and its formula), their columns' widths, and the windows Excel showed.
 */
import type { BookInput, CellInput, CellValue, ColumnInput, DataValidation } from '../../../src/contract.js'
import { formatAddress, indexToCol, parseAddress } from '../../../src/model.js'

export type UiAction =
  /** The formula bar of each cell, selected and then in editing (F2). */
  | { do: 'bars'; cells: string[] }
  /** Text pasted from the clipboard at `at`: the rows' fields joined by tabs, the rows by `newline` (`\n` without it), with one more line break after the last with `trailing`. */
  | { do: 'paste'; at: string; rows: string[][]; newline?: '\r\n' | '\r'; trailing?: true }
  /** A text typed into the selected range and entered into each of its cells (Control+Return); the range's first cell is the active one. */
  | { do: 'fill'; range: string; text: string }

export interface UiCase {
  id: string
  /** The sheet's cells before the action, its columns and its rules of data validation. */
  cells: Record<string, CellInput>
  columns?: Record<string, ColumnInput>
  validations?: DataValidation[]
  act: UiAction
  /** Cells read besides those the action selects or writes. */
  read?: string[]
}

/** What Excel's object model and the saved file say of a cell after the action. */
export interface UiCell {
  /** The cell as Excel saved it, read back by `readXlsx` (its value or formula and its format); null when blank. */
  cell: CellInput | null
  /** A formula's value as Excel saved it. */
  computed?: CellValue
  /** What the cell shows (`Range.Text`) and its format codes (`NumberFormat`, `NumberFormatLocal`). */
  text: string
  code: string
  local: string
  /** The prefix character (`'`); absent when empty. */
  prefix?: string
  /** The number the object model gives as the value; absent for anything else. */
  value2?: number
  /** `Formula`, `FormulaLocal`, `Formula2`, `Formula2Local`. */
  formula: string
  formulaLocal: string
  formula2: string
  formula2Local: string
  /** Part of a formula of the old kind over an area (`HasArray`), with that formula (`FormulaArray`). */
  hasArray: boolean
  formulaArray?: string
  /** Excel for Windows: the cell is part of a spilled array (`HasSpill`), whose formula stands in `spillParent`. */
  hasSpill?: boolean
  spillParent?: string
  /**
   * Excel for Mac, a case of `bars`: the formula bar with the cell selected, and once its editing began (F2), each as
   * its description (`Formula Bar. A,1 <text>`) and its value; the element in focus then, its role and value; the
   * windows Excel showed meanwhile.
   */
  bar?: { selected: [string, string]; edit: [string, string]; focused: string; dialog?: string }
}

/** A column read after the action: `ColumnWidth`, `Width` (points), `UseStandardWidth`, and its `<col>` in the saved file. */
export interface UiColumn {
  columnWidth: number
  points: number
  standard: boolean
  saved: { width: number; customWidth?: true; bestFit?: true; hidden?: true } | null
}

export interface UiAnswer {
  /** The windows Excel showed for a paste or a fill, answered (`answerDialog` of `excel-keys.applescript`): their texts. */
  dialog?: string
  cells: Record<string, UiCell>
  columns: Record<string, UiColumn>
}

export interface UiSet {
  buildUiCases(): UiCase[]
}

/** The data formulas read: numbers and a date with the short date format, in column A. */
export const DATA: Readonly<Record<string, CellInput>> = { A1: { value: 2 }, A2: { value: 46096, numberFormat: 14 }, A3: { value: 3 } }

/** The cells of a range `B1:C2` (or one cell), row by row. */
export function cellsOf(range: string): string[] {
  const [a, b = a] = range.split(':')
  const p = parseAddress(a)
  const q = parseAddress(b)
  const out: string[] = []
  for (let r = p.row; r <= q.row; r++) for (let c = p.col; c <= q.col; c++) out.push(formatAddress(r, c))
  return out
}

/** The texts the action puts into cells, as a host passes them to `setCellTexts`: a paste's fields from its cell on, a fill's text in each cell. */
export function entriesOf(act: UiAction): { sheet: string; cell: string; text: string }[] {
  if (act.do === 'fill') return cellsOf(act.range).map((cell) => ({ sheet: 'S', cell, text: act.text }))
  if (act.do === 'bars') return []
  const at = parseAddress(act.at)
  return act.rows.flatMap((row, r) => row.map((text, c) => ({ sheet: 'S', cell: formatAddress(at.row + r, at.col + c), text })))
}

/** The clipboard's text of a paste. */
export const clipboardText = (act: Extract<UiAction, { do: 'paste' }>): string => {
  const nl = act.newline ?? '\n'
  return act.rows.map((r) => r.join('\t')).join(nl) + (act.trailing ? nl : '')
}

/** The cells read after a case's action: those it selects or writes, then `read`, each once. */
export const readCells = (c: UiCase): string[] => [...new Set([...(c.act.do === 'bars' ? c.act.cells : entriesOf(c.act).map((e) => e.cell)), ...(c.read ?? [])])]

/** The columns of the cells read, by letter. */
export const readColumns = (c: UiCase): string[] => [...new Set(readCells(c).map((a) => indexToCol(parseAddress(a).col)))]

/** The book of a case. */
export const uiBook = (c: UiCase, locale?: string): BookInput => ({
  ...(locale ? { locale } : {}),
  sheets: [{ name: 'S', cells: c.cells, ...(c.columns ? { columns: c.columns } : {}), ...(c.validations ? { validations: c.validations } : {}) }],
})

/** Cases of one group: ids `<group>-<n>`. */
export function collector() {
  const out: UiCase[] = []
  const count = new Map<string, number>()
  const add = (group: string, c: Omit<UiCase, 'id'>) => {
    const n = (count.get(group) ?? 0) + 1
    count.set(group, n)
    out.push({ id: `${group}-${n}`, ...c })
  }
  return { out, add }
}
