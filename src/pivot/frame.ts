/**
 * Where a PivotTable's labels go, without the engine: the columns of the rows' labels, the header rows, the Filters
 * area's places, and which part of the PivotTable each label is (`placeParts`), over the lines of its axes as the engine
 * makes them (`layout.ts`) or as the workbook saved them (`savedLines`, for what a cell is: `cells.ts`).
 */
import type { PivotAxisField, PivotLine, PivotTable } from '../contract.js'
import { formOf, hasItems } from './layout-lines.js'

/**
 * The parts of a PivotTable's cells: a field's name (`field`, on the Filters area too), an item's label on an axis
 * (`item`, of a value field for `values`) or on the Filters area (`page`), a subtotal's label, the grand total's, the
 * header of the rows or columns in compact form, a value field's name (`value`; `total` the classic layout's `Total`).
 */
export type PivotPart =
  | { kind: 'field'; field: PivotAxisField }
  | { kind: 'item' | 'subtotal'; field: PivotAxisField; item: number }
  | { kind: 'page'; field: number }
  | { kind: 'grand' | 'rowHeader' | 'columnHeader' }
  | { kind: 'value'; value: number; total?: true }

/** A line of an axis as its labels need it: its kind, the items of its fields up to its own, where its labels start, its value field. */
export interface LineView {
  /** `data`: an item's line (with its subtotal when the item is not the last field's); `item`: an item over its group, values empty. */
  kind: 'data' | 'item' | 'subtotal' | 'grand' | 'blank'
  items: number[]
  /** The field from which on the line shows a label. */
  from: number
  value?: number
}

/**
 * The lines an axis of `fields` saved (`rowLines`, `columnLines`) as views: each with the items it repeats from the line
 * before, and its value field (0 where the line leaves it out, as the file does).
 */
export function savedLines(lines: readonly PivotLine[] | undefined, fields: readonly PivotAxisField[]): LineView[] {
  const at = fields.indexOf('values')
  let last: number[] = []
  return (lines ?? []).map((l) => {
    const kind = l.type === 'grand' || l.type === 'blank' ? l.type : l.type ? 'subtotal' : 'data'
    const items = kind === 'grand' ? [] : [...last.slice(0, l.repeat ?? 0), ...l.items]
    if (kind !== 'grand') last = items
    // a line under the value fields takes its own; a total over them is of each in turn
    const value = l.value ?? (at < 0 ? undefined : at < items.length ? items[at] : kind === 'subtotal' || kind === 'grand' ? 0 : undefined)
    return { kind, items, from: l.repeat ?? 0, ...(value === undefined ? {} : { value }) }
  })
}

/** A PivotTable's frame: the columns of its rows' labels (and the level of each compact field), its header rows, its Filters area's runs. */
export interface PivotFrame {
  labelCol: number[]
  level: number[]
  /** The columns of the rows' labels, left of the values. */
  width: number
  /** The classic layout's row of the columns' drop zone, with no field on columns. */
  classic: boolean
  /** A caption row over the column fields. */
  captionRow: number
  /** The rows above the values. */
  height: number
  /** Report filters per run, and the rows and columns of the Filters area. */
  run: number
  pageRows: number
  pageColumns: number
}

/** The frame of a PivotTable with `rowFields` and `colFields` on its axes. */
export function pivotFrame(table: PivotTable, rowFields: readonly PivotAxisField[], colFields: readonly PivotAxisField[]): PivotFrame {
  const values = table.values ?? []
  // a compact field shares the column of a compact field before it
  const labelCol: number[] = []
  const level: number[] = []
  rowFields.forEach((f, k) => {
    const shared = k > 0 && formOf(table, f) === 'compact' && formOf(table, rowFields[k - 1]) === 'compact'
    labelCol.push(k === 0 ? 0 : labelCol[k - 1] + (shared ? 0 : 1))
    level.push(shared ? level[k - 1] + 1 : 0)
  })
  const colItems = hasItems(colFields)
  // without a field on rows, a column of its own holds the name of a single value field under column labels (the set areas:
  // none for several value fields, whose names head their columns, nor for none)
  const width = rowFields.length ? labelCol[rowFields.length - 1] + 1 : colItems && values.length === 1 ? 1 : 0
  // the classic layout keeps the row of the columns' drop zone with no column field: the value field's name over it, `Total`
  // over its values (the set corpus-shapes)
  const classic = table.classicLayout === true && !colItems && values.length > 0
  // the caption row over the column fields; over value fields alone only with the Values row shown
  const captionRow = colItems || classic || (table.valuesRow && colFields.includes('values')) ? 1 : 0
  // the Filters area: a field and its item per pair of columns, down then across (or across then down), `pagesWrap` per run
  const pages = table.pages?.length ?? 0
  const run = table.pagesWrap || pages
  const pageRows = pages ? (table.pagesAcrossFirst ? Math.ceil(pages / run) : Math.min(pages, run)) : 0
  const pageColumns = pages ? (table.pagesAcrossFirst ? Math.min(pages, run) : Math.ceil(pages / run)) : 0
  return { labelCol, level, width, classic, captionRow, height: captionRow + Math.max(colFields.length, 1), run, pageRows, pageColumns }
}

/** Where a label goes: the line it is of, the report filter (its place among `pages`), its indent in compact form. */
export interface PartAt {
  line?: LineView
  rows?: boolean
  page?: number
  indent?: number
}

/**
 * Each label of a PivotTable laid out with its top-left cell at `top`, `left`: the Filters area's fields and items, the
 * headers, the labels of the columns' and the rows' lines, given to `put` with the part it is and where it goes.
 */
export function placeParts(table: PivotTable, f: PivotFrame, top: number, left: number, rowFields: readonly PivotAxisField[], colFields: readonly PivotAxisField[], rowLines: readonly LineView[], columnLines: readonly LineView[], put: (row: number, col: number, part: PivotPart, at: PartAt) => void) {
  const values = table.values ?? []
  const { labelCol, level, width, captionRow, height, run } = f
  ;(table.pages ?? []).forEach((p, k) => {
    const [r, pair] = table.pagesAcrossFirst ? [Math.floor(k / run), k % run] : [k % run, Math.floor(k / run)]
    const row = top - f.pageRows - 1 + r
    const col = left + pair * 3
    put(row, col, { kind: 'field', field: p.field }, { page: k })
    put(row, col + 1, { kind: 'page', field: p.field }, { page: k })
  })

  // the headers: the value field's name at the corner, `Column Labels` (the column fields' names in outline and tabular forms), `Row Labels` (the row fields')
  const headers = table.fieldHeaders !== false
  const compact = (table.layout ?? 'compact') === 'compact'
  if (captionRow) {
    if (values.length === 1 && rowFields.length) put(top, left, { kind: 'value', value: 0 }, {})
    if (headers) {
      if (compact) put(top, left + width, { kind: 'columnHeader' }, {})
      else colFields.forEach((c, k) => put(top, left + width + k, { kind: 'field', field: c }, {}))
    }
  }
  if (rowFields.length && headers) {
    rowFields.forEach((r, k) => {
      if (k > 0 && labelCol[k] === labelCol[k - 1]) return
      put(top + height - 1, left + labelCol[k], k === 0 && compact ? { kind: 'rowHeader' } : { kind: 'field', field: r }, {})
    })
  }
  if (!colFields.length && values.length === 1 && !rowFields.includes('values')) put(top + height - 1, left + width, f.classic ? { kind: 'value', value: 0, total: true } : { kind: 'value', value: 0 }, {})
  if (!rowFields.length && width === 1 && values.length === 1) put(top + height, left, { kind: 'value', value: 0 }, {})

  // the labels of the columns' lines, a field per row under the caption row
  columnLines.forEach((a, j) => {
    const col = left + width + j
    const d = a.items.length - 1
    if (a.kind === 'grand') put(top + captionRow, col, { kind: 'grand' }, { line: a, rows: false })
    else if (a.kind === 'subtotal') put(top + captionRow + d, col, { kind: 'subtotal', field: colFields[d], item: a.items[d] }, { line: a, rows: false })
    else for (let k = a.from; k <= d; k++) put(top + captionRow + k, col, { kind: 'item', field: colFields[k], item: a.items[k] }, { line: a, rows: false })
  })

  // the labels of the rows' lines, the deepest of a shared column last; outline and tabular fields repeat theirs on the
  // lines of their group when set to, a subtotal's line of an inner field too
  const repeats = (r: PivotAxisField) => formOf(table, r) !== 'compact' && (r === 'values' ? table.repeatLabels : (table.fields[r]?.repeatLabels ?? table.repeatLabels)) === true
  rowLines.forEach((a, i) => {
    const row = top + height + i
    const d = a.items.length - 1
    if (a.kind === 'grand' || a.kind === 'blank') {
      if (a.kind === 'grand') put(row, left, { kind: 'grand' }, { line: a, rows: true })
      return
    }
    const subtotal = a.kind === 'subtotal'
    for (let k = 0; k <= d; k++) {
      const r = rowFields[k]
      if (subtotal && k === d) put(row, left + labelCol[d], { kind: 'subtotal', field: r, item: a.items[d] }, { line: a, rows: true })
      else if ((!subtotal && k >= a.from) || repeats(r)) put(row, left + labelCol[k], { kind: 'item', field: r, item: a.items[k] }, { line: a, rows: true, indent: level[k] * Math.max(1, table.indent ?? 1) })
    }
  })
}
