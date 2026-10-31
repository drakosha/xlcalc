/**
 * The look Excel keeps in the cells of a PivotTable's area, which `placePivots` writes with its values: a compact form's
 * row labels left aligned and indented by their level; with Merge and center cells with labels (`mergeLabels`), the
 * headers centered, the labels of the tabular form's rows and of the columns centered both ways and wrapped, and merged
 * over the lines they label, the outline form's row labels centered. Measured on Excel's books (sets `basic`, `layout`,
 * `merge-1`).
 */
import type { AlignmentStyle, PivotLayout, PivotTable } from '../contract.js'
import { parseRange, type Rect } from '../structure/area.js'
import type { LaidCell, PivotLaidOut } from './layout.js'

/** The rows of the row labels and the columns of the column labels of a laid out PivotTable. */
interface Frame {
  /** The row labels: from this row, up to this column. */
  rows: { r1: number; c2: number }
  /** The column labels: these rows, from this column. */
  columns: { r1: number; r2: number; c1: number }
  r2: number
  c2: number
}

function frameOf(laid: PivotLaidOut): Frame | null {
  const x = parseRange(laid.location.range)
  if (!x) return null
  const { firstHeaderRow, firstDataRow, firstDataColumn } = laid.location
  return { rows: { r1: x.r1 + firstDataRow, c2: x.c1 + firstDataColumn - 1 }, columns: { r1: x.r1 + firstHeaderRow, r2: x.r1 + firstDataRow - 1, c1: x.c1 + firstDataColumn }, r2: x.r2, c2: x.c2 }
}

const CENTER: AlignmentStyle = { horizontal: 'center' }
const MERGED: AlignmentStyle = { horizontal: 'center', vertical: 'center', wrap: true }

/** A label's form: the PivotTable's, which its fields' own do not change in Excel's cells (`layout/form-per-field`). */
const formOf = (t: PivotTable): PivotLayout => t.layout ?? 'compact'

/**
 * The indent of a compact form's row label: an item's its level's, a subtotal's at the bottom its field's level, at the
 * top and the grand total's none.
 */
function labelIndent(t: PivotTable, c: LaidCell): number {
  const part = c.part!
  if (part.kind === 'item') return c.indent ?? 0
  if (part.kind !== 'subtotal') return 0
  const level = (t.rows ?? []).indexOf(part.field)
  return level > 0 && typeof part.field === 'number' && t.fields[part.field]?.subtotalTop === false ? level * Math.max(1, t.indent ?? 1) : 0
}

/** Where a cell lies: among the row labels, the column labels, or neither. */
function regionOf(f: Frame, c: LaidCell): 'rows' | 'columns' | null {
  if (c.row >= f.rows.r1 && c.col <= f.rows.c2) return 'rows'
  if (c.row >= f.columns.r1 && c.row <= f.columns.r2 && c.col >= f.columns.c1) return 'columns'
  return null
}

/** The alignment Excel keeps in a cell of the area, none for the others (values, a compact form's headers without `mergeLabels`). */
export function cellAlignment(t: PivotTable, laid: PivotLaidOut, c: LaidCell): AlignmentStyle | undefined {
  const f = frameOf(laid)
  const kind = c.part?.kind
  if (!f || !kind) return undefined
  const label = kind === 'item' || kind === 'subtotal' || kind === 'grand'
  const region = label ? regionOf(f, c) : null
  if (region === 'rows' && formOf(t) === 'compact') {
    const indent = labelIndent(t, c)
    return { horizontal: 'left', ...(indent ? { indent } : {}) }
  }
  if (!t.mergeLabels) return undefined
  if (!label) return kind === 'page' ? undefined : CENTER
  if (region === 'columns') return MERGED
  if (region === 'rows') return formOf(t) === 'tabular' ? MERGED : CENTER
  return undefined
}

/**
 * The merged areas of `mergeLabels`: an item of the tabular form's rows, but the last field's, down over the rows of its
 * items; a subtotal's and the grand total's row label across the row labels' columns; an item of the columns, but the
 * last field's, across its items' columns; a subtotal's and the grand total's column label down the column labels' rows.
 */
export function labelMerges(t: PivotTable, laid: PivotLaidOut): Rect[] {
  const f = frameOf(laid)
  if (!t.mergeLabels || !f) return []
  const at = new Set(laid.cells.map((c) => `${c.row},${c.col}`))
  const taken = (r: number, c: number) => at.has(`${r},${c}`)
  const out: Rect[] = []
  const lastRow = (t.rows ?? []).filter((x) => x !== 'values').at(-1)
  const lastColumn = (t.columns ?? []).filter((x) => x !== 'values').at(-1)
  for (const c of laid.cells) {
    const part = c.part
    if (!part || (part.kind !== 'item' && part.kind !== 'subtotal' && part.kind !== 'grand')) continue
    const region = regionOf(f, c)
    if (region === 'rows' && formOf(t) === 'tabular') {
      if (part.kind === 'item' && 'field' in part && part.field !== lastRow) {
        let r = c.row + 1
        while (r <= f.r2 && !taken(r, c.col)) r++
        if (r - 1 > c.row) out.push({ r1: c.row, r2: r - 1, c1: c.col, c2: c.col })
      } else if (part.kind !== 'item') {
        let k = c.col + 1
        while (k <= f.rows.c2 && !taken(c.row, k)) k++
        if (k - 1 > c.col) out.push({ r1: c.row, r2: c.row, c1: c.col, c2: k - 1 })
      }
    } else if (region === 'columns') {
      if (part.kind === 'item' && 'field' in part && part.field !== lastColumn) {
        let k = c.col + 1
        while (k <= f.c2 && !taken(c.row, k)) k++
        if (k - 1 > c.col) out.push({ r1: c.row, r2: c.row, c1: c.col, c2: k - 1 })
      } else if (part.kind !== 'item') {
        let r = c.row + 1
        while (r <= f.columns.r2 && !taken(r, c.col)) r++
        if (r - 1 > c.row) out.push({ r1: c.row, r2: r - 1, c1: c.col, c2: c.col })
      }
    }
  }
  return out
}
