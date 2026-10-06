/**
 * Display cases for Excel (`run.ts display <set> [ru]`): a cell (a value or a formula, a number format), the
 * width of its column and an optional merged area; Excel shows it and the text it shows is read back.
 */
import type { CellInput } from '../../../src/contract.js'

/**
 * What Excel shows in a cell: its text (`Range.Text`) and the colour of its format's section as `#RRGGBB`, left
 * out when black (Excel reads the automatic colour and `[Black]` alike).
 */
export interface DisplayAnswer {
  text: string
  color?: string
}

export interface DisplayCase {
  id: string
  /** The cell as the engine takes it: a value or a formula (errors come from formulas), with its `format`. */
  cell: CellInput
  /** The column's width as xlsx stores it; the sheet's default width without it. */
  width?: number
  /** A merged area of this many columns from the cell, all of `width`. */
  span?: number
  /** A held-out case: the answer predicted before Excel showed it. */
  predicted?: DisplayAnswer
}

export interface DisplaySet {
  buildDisplayCases(): DisplayCase[]
}

/** A cell from a value, or a formula written `=…`, with a format. */
export const cellOf = (v: number | string | boolean | null, format?: string): CellInput => ({
  ...(typeof v === 'string' && v.startsWith('=') ? { formula: v.slice(1) } : { value: v }),
  ...(format !== undefined ? { format } : {}),
})

/** Cases of one group: ids `<group>#<n>`. */
export function collector() {
  const out: DisplayCase[] = []
  const count = new Map<string, number>()
  const add = (group: string, v: number | string | boolean | null, format?: string, extra: Omit<DisplayCase, 'id' | 'cell'> = {}) => {
    const n = (count.get(group) ?? 0) + 1
    count.set(group, n)
    out.push({ id: `${group}#${n}`, cell: cellOf(v, format), ...extra })
  }
  return { out, add }
}
