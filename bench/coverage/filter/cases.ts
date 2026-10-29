/**
 * The AutoFilter under what is entered on its sheet, for Excel (`run.ts filter <set>`, `bench/windows/run.ts filter
 * <set>`): each case is a book of one sheet `S` with an AutoFilter, its criteria applied by Excel where the case has
 * some, and steps taken in turn, the AutoFilter's range read after each. On the Mac the steps are a person's, Excel in
 * front: a text typed into a cell and Return, text pasted from the clipboard (Cmd+V), a range copied (Cmd+C) and pasted
 * (Cmd+V), Fill Down (Cmd+D), cells cleared (Delete) and Undo (Cmd+Z); `set` writes a cell through the object model.
 */
import type { BookInput, CellInput, Session } from '../../../src/contract.js'
import { createSession } from '../../../src/index.js'
import { groupCollector } from '../collect.js'

export type FilterStep =
  /** A text typed into the cell selected, then Return. */
  | { type: string; at: string }
  /** A text written into the cell by the object model (`formula`, `Formula`). */
  | { set: string; at: string }
  /** A text put on the clipboard (tabs between columns, line breaks between rows) and pasted at the cell. */
  | { paste: string; at: string }
  /** The range copied and pasted at the target range. */
  | { copy: string; to: string }
  /** Fill Down over the range: its first row copied into the others. */
  | { fill: string }
  /** The range's contents cleared (Delete). */
  | { clear: string }
  /** Undo of the last step. */
  | { undo: true }

export interface FilterCase {
  id: string
  /** The sheet's cells and its AutoFilter's range. */
  cells: Record<string, CellInput>
  filter: string
  /** Rows hidden in the book (from 1). */
  hidden?: number[]
  /** Criteria Excel applies before the steps (`AutoFilter` with a field and a value), and the rows they hide, the sheet filtered in the engine. */
  criteria?: { field: number; value: string; hides: number[] }
  steps: FilterStep[]
}

export interface FilterAnswer {
  /** The AutoFilter's range after each step, as `A1:C7` (empty when there is none). */
  after: string[]
  /** The AutoFilter's range in the file Excel saved (null for none). */
  saved: string | null
  /** The windows Excel showed meanwhile, answered. */
  dialog?: string
}

export interface FilterSet {
  buildFilterCases(): FilterCase[]
}

/** Cases of one group: ids `<group>-<n>`. */
export const collector = () => groupCollector<FilterCase>()

/** A table of `rows` rows over the columns `cols` from row 1: a header of the column's letter, then the row's number in each cell but the first column's, which holds North or South in turn. */
export function data(cols: string, rows: number, from = 'A'): Record<string, CellInput> {
  const out: Record<string, CellInput> = {}
  const at = from.charCodeAt(0)
  for (let r = 1; r <= rows; r++)
    for (let k = 0; k < cols.length; k++) {
      const c = String.fromCharCode(at + k)
      out[`${c}${r}`] = { value: r === 1 ? cols[k] : k === 0 ? (r % 2 ? 'South' : 'North') : r }
    }
  return out
}

/** The book of a case: the hidden rows its criteria or the case hide, the sheet filtered with criteria. */
export function filterBook(c: FilterCase): BookInput {
  const rows = Object.fromEntries([...(c.hidden ?? []), ...(c.criteria?.hides ?? [])].map((r) => [String(r), { hidden: true }]))
  return { sheets: [{ name: 'S', cells: c.cells, autoFilter: { range: c.filter }, ...(Object.keys(rows).length ? { rows } : {}), ...(c.criteria ? { filtered: true } : {}) }] }
}

/** A step taken by a session as a host takes it: typed and set texts by `setCellText`, a paste by `pasteText`, a copy and Fill Down by `editStructure`. */
export function stepInEngine(s: Session, step: FilterStep): void {
  if ('type' in step) s.setCellText('S', step.at, step.type)
  else if ('set' in step) s.setCellText('S', step.at, step.set)
  else if ('paste' in step) s.pasteText('S', step.at, step.paste)
  else if ('copy' in step) s.editStructure({ op: 'copyRange', from: { sheet: 'S', range: step.copy }, to: { sheet: 'S', range: step.to } })
  else if ('fill' in step) {
    const [a, b] = step.fill.split(':')
    const first = `${a}:${b.replace(/\d+$/, a.replace(/^[A-Z]+/, ''))}`
    s.editStructure({ op: 'fillRange', sheet: 'S', range: first, to: step.fill, type: 'copy' })
  } else if ('clear' in step) s.clearCells([{ sheet: 'S', range: step.clear }], 'contents')
  else s.undo()
}

/** The engine's answer to a case: the range its AutoFilter acts on after each step (`getAutoFilterRange`), and its own range at the end, which a file saves. */
export function filterInEngine(c: FilterCase): FilterAnswer {
  const s = createSession(filterBook(c))
  const after = c.steps.map((step) => {
    stepInEngine(s, step)
    return s.getAutoFilterRange('S') ?? ''
  })
  const saved = s.getLayout().sheets[0].autoFilter?.range ?? null
  s.close()
  return { after, saved }
}
