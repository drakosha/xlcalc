/**
 * What Excel does with what a person does in its window, for Excel (`run.ts ui <set> [ru|<tag>]`, `bench/windows/run.ts
 * ui <set> [ru|<tag>]`): each case is a book of one sheet `S` and one action, on the Mac with Excel in front. `bars`
 * selects cells in turn and reads the formula bar, then enters each cell's editing (F2) and reads what it opens with;
 * `paste` puts text on the clipboard (tabs between columns, line breaks between rows) and pastes it into a cell (Cmd+V);
 * `fill` selects a range, types a text and enters it into every cell of the selection (Control+Return); `ribbon` selects
 * cells in turn and reads the category the Number Format box of the Home tab shows; `choose` applies items of that box's
 * list to cells. Then the cells the action touched are read (what each holds, shows and its formula), their columns'
 * widths, and the windows Excel showed.
 */
import type { BookInput, CellInput, CellValue, ColumnInput, DataValidation, NumberFormat } from '../../../src/contract.js'
import { formatAddress, indexToCol, parseAddress } from '../../../src/model.js'
import { numberFormatCategory, numberFormatChoices } from '../../../src/text.js'

export type UiAction =
  /** The formula bar of each cell, selected and then in editing (F2). */
  | { do: 'bars'; cells: string[] }
  /** Text pasted from the clipboard at `at`: the rows' fields joined by tabs, the rows by `newline` (`\n` without it), with one more line break after the last with `trailing`. */
  | { do: 'paste'; at: string; rows: string[][]; newline?: '\r\n' | '\r'; trailing?: true }
  /** A text typed into the selected range and entered into each of its cells (Control+Return); the range's first cell is the active one. */
  | { do: 'fill'; range: string; text: string }
  /** The category the ribbon's Number Format box (the Home tab) shows for each cell, selected in turn; A1 and A2 stay empty, the script's control cells. */
  | { do: 'ribbon'; cells: string[] }
  /** Items of the ribbon's Number Format list applied, each to its cell selected: `[cell, item]`, the item by its place in the list (1 General to 11 Text); A1 and A2 stay empty. */
  | { do: 'choose'; picks: [string, number][] }
  /**
   * The cell's editing begun with F2, as a person edits a cell in place (the cursor at the end of what it opens with):
   * by default all it opened with erased and `text` typed, or with `keys` Left pressed `left` times, Backspace `back`
   * times and `type` typed; then Return. `text` is what the cell holds then, as a host passes it to `setCellText`.
   */
  | { do: 'edit'; at: string; text: string; keys?: { left: number; back: number; type: string } }

export interface UiCase {
  id: string
  /** The sheet's cells before the action, its columns and its rules of data validation. */
  cells: Record<string, CellInput>
  columns?: Record<string, ColumnInput>
  validations?: DataValidation[]
  act: UiAction
  /** Cells read besides those the action selects or writes. */
  read?: string[]
  /** Cells whose format code is written as a code of the file's own (an id from 164), though the writer stores a built-in format's code by its number. */
  own?: string[]
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
   * windows Excel showed meanwhile; a case of `edit` reads them too, and the bar once its keys were typed (`typed`).
   */
  bar?: { selected: [string, string]; edit: [string, string]; focused: string; dialog?: string; typed?: [string, string] }
  /** Excel for Mac, a case of `ribbon` or `choose`: what the Number Format box shows with the cell selected (after the item applied), and the item's name in the list (its title and sample on two lines). */
  ribbon?: { box: string; item?: string }
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

/** The actions Excel takes only on the Mac's screen: the keyboard's fill and edit, the ribbon's box and list. */
export const macOnly = (act: UiAction): boolean => act.do === 'fill' || act.do === 'edit' || act.do === 'ribbon' || act.do === 'choose'

/** The cells an action selects in turn without writing: the formula bar's and the ribbon's; a choice's cells by its picks. */
export function selectedOf(act: UiAction): string[] {
  if (act.do === 'bars' || act.do === 'ribbon') return act.cells
  return act.do === 'choose' ? act.picks.map(([cell]) => cell) : []
}

/** The texts the action puts into cells, as a host passes them to `setCellTexts`: a paste's fields from its cell on, a fill's text in each cell, an edit's text. */
export function entriesOf(act: UiAction): { sheet: string; cell: string; text: string }[] {
  if (act.do === 'fill') return cellsOf(act.range).map((cell) => ({ sheet: 'S', cell, text: act.text }))
  if (act.do === 'edit') return [{ sheet: 'S', cell: act.at, text: act.text }]
  if (act.do !== 'paste') return []
  const at = parseAddress(act.at)
  return act.rows.flatMap((row, r) => row.map((text, c) => ({ sheet: 'S', cell: formatAddress(at.row + r, at.col + c), text })))
}

/** The clipboard's text of a paste. */
export const clipboardText = (act: Extract<UiAction, { do: 'paste' }>): string => {
  const nl = act.newline ?? '\n'
  return act.rows.map((r) => r.join('\t')).join(nl) + (act.trailing ? nl : '')
}

/** The cells read after a case's action: those it selects or writes, then `read`, each once. */
export const readCells = (c: UiCase): string[] => [...new Set([...selectedOf(c.act), ...entriesOf(c.act).map((e) => e.cell), ...(c.read ?? [])])]

/** The columns of the cells read, by letter. */
export const readColumns = (c: UiCase): string[] => [...new Set(readCells(c).map((a) => indexToCol(parseAddress(a).col)))]

/** The window from A1 that holds the cells given, for `getViewport`. */
export function windowOver(cells: string[]): string {
  const at = cells.map(parseAddress)
  return `A1:${formatAddress(Math.max(...at.map((p) => p.row)), Math.max(...at.map((p) => p.col)))}`
}

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

/** What the ribbon shows in a cell of a case: the box's category; for an item applied, the item's name, the format it applies and the box's category after. */
export type RibbonCell = string | { name: string; format: NumberFormat; box: string }

/** The ribbon of each cell of a case of `ribbon` or `choose` by `numberFormatCategory` (a cell's format, General without one) and `numberFormatChoices`. */
export function ribbonOf(c: UiCase, tag: string): Record<string, RibbonCell> {
  const act = c.act
  if (act.do === 'ribbon') return Object.fromEntries(act.cells.map((a) => [a, numberFormatCategory(c.cells[a]?.numberFormat ?? 0, tag).name]))
  if (act.do !== 'choose') throw new Error(`${c.id} is no case of the ribbon`)
  const choices = numberFormatChoices(tag)
  return Object.fromEntries(act.picks.map(([a, k]) => [a, { name: choices[k - 1].name, format: choices[k - 1].format, box: numberFormatCategory(choices[k - 1].format, tag).name }]))
}
