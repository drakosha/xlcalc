/**
 * Display cases for Excel (`run.ts display <set> [ru]`): a cell (a value or a formula, a number format), the
 * width of its column and an optional merged area; Excel shows it and the text it shows is read back.
 */
import type { StoredCell } from '../stored.js'

/**
 * What Excel shows in a cell: its text (`Range.Text`) and the colour of its format's section as `#RRGGBB`, left
 * out when black (Excel reads the automatic colour and `[Black]` alike).
 */
export interface DisplayAnswer {
  text: string
  color?: string
  /** For a case with `numFmtId`: the cell's format code as Excel reports it (`NumberFormat`) and in the interface's language (`NumberFormatLocal`). */
  code?: string
  local?: string
}

export interface DisplayCase {
  id: string
  /** The cell as the engine takes it: a value or a formula (errors come from formulas), with its `format`. */
  cell: StoredCell
  /** The column's width as xlsx stores it; the sheet's default width without it. */
  width?: number
  /** A merged area of this many columns from the cell, all of `width`. */
  span?: number
  /** The format is written by its number when it is one of Excel's built-in formats (else as a custom format). */
  builtin?: true
  /** The cell's format is the built-in format of this number, written by its number alone (`numFmtId` without a code); `cell.format` is left out. */
  numFmtId?: number
  /** A held-out case: the answer predicted before Excel showed it. */
  predicted?: DisplayAnswer
}

export interface DisplaySet {
  buildDisplayCases(): DisplayCase[]
}

/** A cell from a value, or a formula written `=…`, with a format. */
export const cellOf = (v: number | string | boolean | null, format?: string): StoredCell => ({
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
