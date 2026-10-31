/**
 * A PivotTable laid out from its calculation, as Excel lays it out on refresh: the lines of its rows and columns
 * (`rowLines`, `columnLines`), its place (`location`, the Filters area above it), and the cells of its area: the
 * headers, the items' labels, the values with their value field's number format, the subtotals and grand totals.
 */
import type { BookInput, CellValue, ErrorCode, NumberFormat, PivotAxisField, PivotCache, PivotLine, PivotLocation, PivotShowAs, PivotTable } from '../contract.js'
import { dateToSerial } from '../date-serial.js'
import { epochOf } from '../dates.js'
import { displayRules } from '../locale/rules.js'
import { formatAddress, parseAddress } from '../model.js'
import { cornersText, type Rect } from '../structure/area.js'
import { pivotArea } from './area.js'
import { wordsOf } from './captions.js'
import { groupedItems } from './items.js'
import { itemFormat } from '../pivot-format.js'
import { axisLines, formOf, hasItems, type AxisLine } from './layout-lines.js'
import type { PivotResult } from './result.js'
import { shownValues, type PivotValueAt } from './show.js'

/** A cell of a PivotTable's area: its place on the sheet (zero-based), its value or error, its number format, its indent in compact form. */
export interface PivotCell {
  row: number
  col: number
  value?: string | number | boolean | { date: string }
  error?: ErrorCode
  numberFormat?: NumberFormat
  /** Levels of indent of a label in a column of several fields (compact form). */
  indent?: number
  /** The part of the PivotTable the cell shows, which an entry into it changes (`pivotEntry`); a value without it. */
  part?: PivotPart
}

/**
 * The parts of a PivotTable's cells: a field's name (`field`, on the Filters area too), an item's label on an axis
 * (`item`, of a value field for `values`) or on the Filters area (`page`), a subtotal's label, the grand total's, the
 * header of the rows or columns in compact form, a value field's name at the corner (`value`).
 */
export type PivotPart =
  | { kind: 'field'; field: PivotAxisField }
  | { kind: 'item' | 'subtotal'; field: PivotAxisField; item: number }
  | { kind: 'page'; field: number }
  | { kind: 'grand' | 'rowHeader' | 'columnHeader' }
  | { kind: 'value'; value: number }

/** A PivotTable laid out: its saved layout and place, the rectangle of its cells, its Filters area included, and the cells. */
export interface PivotLaidOut {
  location: PivotLocation
  rowLines: PivotLine[]
  columnLines: PivotLine[]
  /** The whole area on the sheet, the Filters area included. */
  area: Rect
  cells: PivotCell[]
}

/** The cells of a PivotTable with no field anywhere: Excel keeps the place of its drop zones, 18 rows by 3 columns. */
const EMPTY_ROWS = 18
const EMPTY_COLS = 3

/** The calculations Excel shows as a percentage (0.00%, built-in format 10) for a value field without a format of its own. */
const PERCENT = new Set<PivotShowAs>(['percentOfTotal', 'percentOfColumn', 'percentOfRow', 'percent', 'percentOfParentRow', 'percentOfParentColumn', 'percentOfParent', 'percentDiff', 'percentOfRunningTotal'])

type Label = { value: PivotCell['value']; error?: ErrorCode; numberFormat?: NumberFormat }

/** A PivotTable's labels: of its fields, value fields and items, as cells and as text. */
export function labeler(table: PivotTable, cache: PivotCache, book: Pick<BookInput, 'locale' | 'dateSystem'>) {
  const words = wordsOf(book.locale)
  const rules = displayRules(book.locale, epochOf(book.dateSystem))
  const values = table.values ?? []
  const fieldName = (f: PivotAxisField) => (f === 'values' ? (table.captions?.values ?? words.values) : (table.fields[f]?.name ?? cache.fields[f]?.name ?? ''))
  const valueName = (v: number) => {
    const d = values[v]
    return d?.name ?? (d ? words.valueName(d.function ?? 'sum', cache.fields[d.field]?.name ?? '') : '')
  }
  // an item's label: its caption, else its value in the cache (`(blank)` for none), a date with its field's format
  const itemLabel = (f: PivotAxisField, item: number): Label => {
    if (f === 'values') return { value: valueName(item) }
    const fi = table.fields[f]?.items?.[item]
    const cf = cache.fields[f]
    const v = fi?.item === undefined ? undefined : groupedItems(cf)?.[fi.item]
    const numberFormat = typeof v === 'number' || (typeof v === 'object' && v !== null && 'date' in v) ? (table.fields[f]?.numberFormat ?? itemFormat(cache, f, v)) : undefined
    // a caption keeps the item's format (the set edit-held: a date item named `first`)
    if (fi?.caption !== undefined) return numberFormat === undefined ? { value: fi.caption } : { value: fi.caption, numberFormat }
    if (v === null || v === undefined) return { value: words.blank }
    // an error and a boolean are the locale's words for them, as text
    if (typeof v === 'boolean' || (typeof v === 'object' && 'error' in v)) return { value: rules.display(v, undefined, Infinity).text }
    if (typeof v === 'object') return { value: v, numberFormat }
    return numberFormat === undefined ? { value: v } : { value: v, numberFormat }
  }
  const labelText = (f: PivotAxisField, item: number): string => {
    const l = itemLabel(f, item)
    const v = l.value
    const shown: CellValue = typeof v === 'object' && v ? dateToSerial(v.date, { dateSystem: book.dateSystem }) : (v ?? null)
    return rules.display(shown, l.numberFormat, Infinity).text
  }
  return { words, rules, fieldName, valueName, itemLabel, labelText }
}

/**
 * Lays out the PivotTable of a calculation; `book` gives the language of its words and the date system. The values
 * come as their value fields show them (`shown`, Show Values As over the calculation's).
 */
export function layoutPivot(result: PivotResult, book: Pick<BookInput, 'locale' | 'dateSystem'>, shown: PivotValueAt = shownValues(result)): PivotLaidOut {
  const { table, cache } = result
  const values = table.values ?? []
  const rowFields = result.rows.fields
  const colFields = result.columns.fields
  const rowLines = axisLines(table, result.rows, true, table.grandTotals?.rows !== false)
  const columnLines = axisLines(table, result.columns, false, table.grandTotals?.columns !== false)
  const { words, rules, fieldName, valueName, itemLabel, labelText } = labeler(table, cache, book)
  // a subtotal's label: the field's own caption, else its item's with `Total`, the function's name, or the value field's
  const subtotalLabel = (fields: readonly PivotAxisField[], a: AxisLine): string => {
    const d = a.items.length - 1
    const f = fields[d]
    const own = f === 'values' ? undefined : table.fields[f]?.subtotalCaption
    const item = labelText(f, a.items[d])
    // a caption of its own stands `?` for the item's label (`? all` is `East all`)
    if (own !== undefined) return own.replaceAll('?', item)
    const under = fields.indexOf('values') > d && a.value !== undefined
    return `${item} ${under ? valueName(a.value!) : a.fn ? words.functions[a.fn] : words.total}`
  }
  // the grand total's label takes the format of its axis's first field where the field has one of its own (the sets
  // filter-held and sort-held: `Grand Total` under items of `yyyy-mm-dd hh:mm`)
  const ownFormat = (f: PivotAxisField | undefined) => (f === undefined || f === 'values' ? undefined : (table.fields[f]?.numberFormat ?? cache.fields[f]?.numberFormat))
  const grandLabel = (a: AxisLine, rows: boolean): Label => {
    const value = a.value !== undefined ? `${words.total} ${valueName(a.value)}` : (table.captions?.grandTotal ?? words.grandTotal)
    const numberFormat = ownFormat((rows ? rowFields : colFields)[0])
    return numberFormat === undefined ? { value } : { value, numberFormat }
  }

  // the columns of the rows' labels: a compact field shares the column of a compact field before it
  const labelCol: number[] = []
  const level: number[] = []
  rowFields.forEach((f, k) => {
    const shared = k > 0 && formOf(table, f) === 'compact' && formOf(table, rowFields[k - 1]) === 'compact'
    labelCol.push(k === 0 ? 0 : labelCol[k - 1] + (shared ? 0 : 1))
    level.push(shared ? level[k - 1] + 1 : 0)
  })
  const colItems = hasItems(colFields)
  const width = rowFields.length ? labelCol[rowFields.length - 1] + 1 : colItems ? 1 : 0
  // the caption row over the column fields; over value fields alone only with the Values row shown
  const captionRow = colItems || (table.valuesRow && colFields.includes('values')) ? 1 : 0
  const height = captionRow + Math.max(colFields.length, 1)
  const empty = !rowFields.length && !colFields.length && !values.length
  const dataCols = empty ? EMPTY_COLS - width : !values.length && !colFields.length ? 0 : columnLines.length

  // the Filters area: a field and its item per pair of columns, down then across (or across then down), `pagesWrap` per run
  const pages = result.pages
  const run = table.pagesWrap || pages.length
  const pageRows = pages.length ? (table.pagesAcrossFirst ? Math.ceil(pages.length / run) : Math.min(pages.length, run)) : 0
  const pageCols = pages.length ? (table.pagesAcrossFirst ? Math.min(pages.length, run) : Math.ceil(pages.length / run)) : 0

  // the table keeps its top-left cell, moved down only when the Filters area needs the rows above it
  const at = parseAddress(table.location.range.split(':')[0])
  const top = pageRows ? Math.max(at.row, pageRows + 1) : at.row
  const left = at.col
  const rows = empty ? EMPTY_ROWS : height + rowLines.length
  const range: Rect = { r1: top, c1: left, r2: top + rows - 1, c2: left + Math.max(width + dataCols, 1) - 1 }
  const location: PivotLocation = {
    range: cornersText(range),
    firstHeaderRow: colFields.length ? height - colFields.length : 1,
    firstDataRow: empty ? 1 : height,
    firstDataColumn: width,
  }
  if (pageRows) Object.assign(location, { pageRows, pageColumns: pageCols })

  const cells = new Map<string, PivotCell>()
  const put = (row: number, col: number, l: Label | string, part?: PivotPart, indent = 0) => {
    const cell: PivotCell = { row, col }
    if (part) cell.part = part
    const label = typeof l === 'string' ? { value: l } : l
    if (label.value !== undefined) cell.value = label.value
    if (label.error) cell.error = label.error
    if (label.numberFormat !== undefined) cell.numberFormat = label.numberFormat
    if (indent) cell.indent = indent
    cells.set(formatAddress(row, col), cell)
  }
  const out: PivotLaidOut = { location, rowLines: rowLines.map((a) => a.line), columnLines: columnLines.map((a) => a.line), area: pivotArea({ location })!, cells: [] }
  if (empty) return out

  pages.forEach((p, k) => {
    const [r, pair] = table.pagesAcrossFirst ? [Math.floor(k / run), k % run] : [k % run, Math.floor(k / run)]
    const row = top - pageRows - 1 + r
    const col = left + pair * 3
    put(row, col, fieldName(p.field), { kind: 'field', field: p.field })
    put(row, col + 1, p.all ? words.all : p.items.length === 1 ? itemLabel(p.field, p.items[0]) : words.multiple, { kind: 'page', field: p.field })
  })

  // the headers: the value field's name at the corner, `Column Labels` (the column fields' names in outline and tabular forms), `Row Labels` (the row fields')
  const headers = table.fieldHeaders !== false
  const compact = (table.layout ?? 'compact') === 'compact'
  if (captionRow) {
    if (values.length === 1 && rowFields.length) put(top, left, valueName(0), { kind: 'value', value: 0 })
    if (headers) {
      if (compact) put(top, left + width, table.captions?.columnHeader ?? words.columnHeader, { kind: 'columnHeader' })
      else colFields.forEach((f, k) => put(top, left + width + k, fieldName(f), { kind: 'field', field: f }))
    }
  }
  if (rowFields.length && headers) {
    rowFields.forEach((f, k) => {
      if (k > 0 && labelCol[k] === labelCol[k - 1]) return
      const header = k === 0 && compact
      put(top + height - 1, left + labelCol[k], header ? (table.captions?.rowHeader ?? words.rowHeader) : fieldName(f), header ? { kind: 'rowHeader' } : { kind: 'field', field: f })
    })
  }
  if (!colFields.length && values.length === 1 && !rowFields.includes('values')) put(top + height - 1, left + width, valueName(0), { kind: 'value', value: 0 })
  if (!rowFields.length && width === 1 && values.length === 1) put(top + height, left, valueName(0), { kind: 'value', value: 0 })

  // the labels of the columns' lines, a field per row under the caption row
  columnLines.forEach((a, j) => {
    const col = left + width + j
    const d = a.items.length - 1
    if (a.kind === 'grand') put(top + captionRow, col, grandLabel(a, false), { kind: 'grand' })
    else if (a.kind === 'subtotal') put(top + captionRow + d, col, subtotalLabel(colFields, a), { kind: 'subtotal', field: colFields[d], item: a.items[d] })
    else for (let k = a.from; k <= d; k++) put(top + captionRow + k, col, itemLabel(colFields[k], a.items[k]), { kind: 'item', field: colFields[k], item: a.items[k] })
  })

  // the labels of the rows' lines, the deepest of a shared column last; outline and tabular fields repeat theirs on the
  // lines of their group when set to, a subtotal's line of an inner field too
  const repeats = (f: PivotAxisField) => formOf(table, f) !== 'compact' && (f === 'values' ? table.repeatLabels : (table.fields[f]?.repeatLabels ?? table.repeatLabels)) === true
  rowLines.forEach((a, i) => {
    const row = top + height + i
    const d = a.items.length - 1
    if (a.kind === 'grand' || a.kind === 'blank') {
      if (a.kind === 'grand') put(row, left, grandLabel(a, true), { kind: 'grand' })
      return
    }
    const subtotal = a.kind === 'subtotal'
    for (let k = 0; k <= d; k++) {
      const f = rowFields[k]
      if (subtotal && k === d) put(row, left + labelCol[d], subtotalLabel(rowFields, a), { kind: 'subtotal', field: f, item: a.items[d] })
      else if ((!subtotal && k >= a.from) || repeats(f)) put(row, left + labelCol[k], itemLabel(f, a.items[k]), { kind: 'item', field: f, item: a.items[k] }, level[k] * Math.max(1, table.indent ?? 1))
    }
  })

  // the text shown for an empty cell or an error, a number as a number (`0`), as an entry takes it; a dash alone, the
  // accounting format's zero, is 0 too (the set basic: `-` and ` - ` are 0, `--` and `n/a` text)
  const typed = (text: string) => (text.trim() === '-' ? 0 : (rules.numberOnly(text) ?? text))
  // the values: each line with values against each column's, by the value field of either and a subtotal's function
  if (values.length && dataCols)
    rowLines.forEach((ra, i) => {
      if (ra.kind === 'blank') return
      columnLines.forEach((ca, j) => {
        const v = ra.value ?? ca.value ?? 0
        const numberFormat = values[v].numberFormat ?? (PERCENT.has(values[v].showAs ?? 'normal') ? 10 : undefined)
        const row = top + height + i
        const col = left + width + j
        // an item's line over its group shows no value, but the format of the first value field (the set show-items)
        if (ra.kind === 'item') {
          if (numberFormat !== undefined) put(row, col, { value: undefined, numberFormat })
          return
        }
        const x = shown(ra.path, ca.path, v, ra.fn ?? ca.fn)
        const cell: Label = { value: undefined }
        if (numberFormat !== undefined) cell.numberFormat = numberFormat
        if (x === null) {
          if (table.empty?.show !== false && table.empty?.text) cell.value = typed(table.empty.text)
        } else if (typeof x === 'object') {
          if (table.errors?.show && table.errors.text !== undefined) cell.value = typed(table.errors.text)
          else cell.error = x.error
        } else cell.value = x
        if (cell.value !== undefined || cell.error || cell.numberFormat !== undefined) put(row, col, cell)
      })
    })
  out.cells = [...cells.values()]
  return out
}
