/**
 * A PivotTable laid out from its calculation, as Excel lays it out on refresh: the lines of its rows and columns
 * (`rowLines`, `columnLines`), its place (`location`, the Filters area above it), and the cells of its area: the
 * headers, the items' labels, the values with their value field's number format, the subtotals and grand totals.
 */
import type { BookInput, CellValue, ErrorCode, NumberFormat, PivotAxisField, PivotLine, PivotLocation, PivotTable } from '../contract.js'
import { dateToSerial } from '../date-serial.js'
import { epochOf } from '../dates.js'
import { displayRules } from '../locale/rules.js'
import { formatAddress, parseAddress } from '../model.js'
import { cornersText, type Rect } from '../structure/area.js'
import { wordsOf, type PivotWords } from './captions.js'
import { axisLines, formOf, hasItems, type AxisLine } from './layout-lines.js'
import type { PivotResult } from './result.js'

/** A cell of a PivotTable's area: its place on the sheet (zero-based), its value or error, its number format, its indent in compact form. */
export interface PivotCell {
  row: number
  col: number
  value?: string | number | boolean | { date: string }
  error?: ErrorCode
  numberFormat?: NumberFormat
  /** Levels of indent of a label in a column of several fields (compact form). */
  indent?: number
}

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

type Label = { value: PivotCell['value']; error?: ErrorCode; numberFormat?: NumberFormat }

/** Lays out the PivotTable of a calculation; `book` gives the language of its words and the date system. */
export function layoutPivot(result: PivotResult, book: Pick<BookInput, 'locale' | 'dateSystem'>): PivotLaidOut {
  const { table, cache } = result
  const words = wordsOf(book.locale)
  const rules = displayRules(book.locale, epochOf(book.dateSystem))
  const values = table.values ?? []
  const rowFields = result.rows.fields
  const colFields = result.columns.fields
  const rowLines = axisLines(table, result.rows, true, table.grandTotals?.columns !== false)
  const columnLines = axisLines(table, result.columns, false, table.grandTotals?.rows !== false)
  const fieldName = (f: PivotAxisField) => (f === 'values' ? (table.captions?.values ?? words.values) : (table.fields[f]?.name ?? cache.fields[f]?.name ?? ''))
  const valueName = (v: number) => {
    const d = values[v]
    return d?.name ?? (d ? words.valueName(d.function ?? 'sum', cache.fields[d.field]?.name ?? '') : '')
  }
  const formatOf = (f: number, date: boolean): NumberFormat | undefined => table.fields[f]?.numberFormat ?? cache.fields[f]?.numberFormat ?? (date ? 14 : undefined)
  // an item's label: its caption, else its value in the cache (`(blank)` for none), a date with its field's format
  const itemLabel = (f: PivotAxisField, item: number): Label => {
    if (f === 'values') return { value: valueName(item) }
    const fi = table.fields[f]?.items?.[item]
    if (fi?.caption !== undefined) return { value: fi.caption }
    const cf = cache.fields[f]
    const v = fi?.item === undefined ? undefined : (cf?.items ?? cf?.group?.items)?.[fi.item]
    if (v === null || v === undefined) return { value: words.blank }
    if (typeof v === 'object') return 'error' in v ? { value: undefined, error: v.error } : { value: v, numberFormat: formatOf(f, true) }
    // a boolean is the locale's word for it, as text
    if (typeof v === 'boolean') return { value: rules.display(v, undefined, Infinity).text }
    const numberFormat = typeof v === 'number' ? formatOf(f, false) : undefined
    return numberFormat === undefined ? { value: v } : { value: v, numberFormat }
  }
  const labelText = (f: PivotAxisField, item: number): string => {
    const l = itemLabel(f, item)
    if (l.error) return rules.display({ error: l.error }, undefined, Infinity).text
    const v = l.value
    const shown: CellValue = typeof v === 'object' && v ? dateToSerial(v.date, { dateSystem: book.dateSystem }) : (v ?? null)
    return rules.display(shown, l.numberFormat, Infinity).text
  }
  // a subtotal's label: the field's own caption, else its item's with `Total`, the function's name, or the value field's
  const subtotalLabel = (fields: readonly PivotAxisField[], a: AxisLine): string => {
    const d = a.items.length - 1
    const f = fields[d]
    const own = f === 'values' ? undefined : table.fields[f]?.subtotalCaption
    if (own !== undefined) return own
    const item = labelText(f, a.items[d])
    const under = fields.indexOf('values') > d && a.value !== undefined
    return `${item} ${under ? valueName(a.value!) : a.fn ? words.functions[a.fn] : words.total}`
  }
  const grandLabel = (a: AxisLine) => (a.value !== undefined ? `${words.total} ${valueName(a.value)}` : (table.captions?.grandTotal ?? words.grandTotal))

  // the columns of the rows' labels: a compact field shares the column of the field before it
  const labelCol: number[] = []
  const level: number[] = []
  rowFields.forEach((f, k) => {
    const shared = k > 0 && formOf(table, f) === 'compact' && formOf(table, rowFields[k - 1]) !== 'tabular'
    labelCol.push(k === 0 ? 0 : labelCol[k - 1] + (shared ? 0 : 1))
    level.push(shared ? level[k - 1] + 1 : 0)
  })
  const colItems = hasItems(colFields)
  const width = rowFields.length ? labelCol[rowFields.length - 1] + 1 : colItems ? 1 : 0
  const captionRow = colItems ? 1 : 0
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
  const put = (row: number, col: number, l: Label | string, indent = 0) => {
    const cell: PivotCell = { row, col }
    const label = typeof l === 'string' ? { value: l } : l
    if (label.value !== undefined) cell.value = label.value
    if (label.error) cell.error = label.error
    if (label.numberFormat !== undefined) cell.numberFormat = label.numberFormat
    if (indent) cell.indent = indent
    cells.set(formatAddress(row, col), cell)
  }
  const out: PivotLaidOut = { location, rowLines: rowLines.map((a) => a.line), columnLines: columnLines.map((a) => a.line), area: { ...range, r1: range.r1 - (pageRows ? pageRows + 1 : 0) }, cells: [] }
  if (empty) return out

  pages.forEach((p, k) => {
    const [r, pair] = table.pagesAcrossFirst ? [Math.floor(k / run), k % run] : [k % run, Math.floor(k / run)]
    const row = top - pageRows - 1 + r
    const col = left + pair * 3
    put(row, col, fieldName(p.field))
    put(row, col + 1, p.all ? words.all : p.items.length === 1 ? itemLabel(p.field, p.items[0]) : words.multiple)
  })

  // the headers: the value field's name at the corner, `Column Labels` (the column fields' names in outline and tabular forms), `Row Labels` (the row fields')
  const headers = table.fieldHeaders !== false
  const compact = (table.layout ?? 'compact') === 'compact'
  if (colItems) {
    if (values.length === 1 && rowFields.length) put(top, left, valueName(0))
    if (headers) {
      if (compact) put(top, left + width, table.captions?.columnHeader ?? words.columnHeader)
      else colFields.forEach((f, k) => put(top, left + width + k, fieldName(f)))
    }
  }
  if (rowFields.length && headers) {
    if (compact) put(top + height - 1, left, table.captions?.rowHeader ?? words.rowHeader)
    else rowFields.forEach((f, k) => (k === 0 || labelCol[k] !== labelCol[k - 1]) && put(top + height - 1, left + labelCol[k], fieldName(f)))
  }
  if (!colFields.length && values.length === 1 && !rowFields.includes('values')) put(top + height - 1, left + width, valueName(0))
  if (!rowFields.length && width === 1 && values.length === 1) put(top + height, left, valueName(0))

  // the labels of the columns' lines, a field per row under the caption row
  columnLines.forEach((a, j) => {
    const col = left + width + j
    const d = a.items.length - 1
    if (a.kind === 'grand') put(top + captionRow, col, grandLabel(a))
    else if (a.kind === 'subtotal') put(top + captionRow + d, col, subtotalLabel(colFields, a))
    else for (let k = a.from; k <= d; k++) put(top + captionRow + k, col, itemLabel(colFields[k], a.items[k]))
  })

  // the labels of the rows' lines, the deepest of a shared column last; outline and tabular fields repeat theirs when set to
  rowLines.forEach((a, i) => {
    const row = top + height + i
    const d = a.items.length - 1
    if (a.kind === 'grand') put(row, left, grandLabel(a))
    else if (a.kind === 'subtotal') put(row, left + labelCol[d], subtotalLabel(rowFields, a))
    else if (a.kind !== 'blank')
      for (let k = 0; k <= d; k++) {
        const f = rowFields[k]
        const repeat = k < a.from && formOf(table, f) !== 'compact' && (f === 'values' ? table.repeatLabels : (table.fields[f]?.repeatLabels ?? table.repeatLabels))
        if (k >= a.from || repeat) put(row, left + labelCol[k], itemLabel(f, a.items[k]), level[k] * Math.max(1, table.indent ?? 1))
      }
  })

  // the values: each line with values against each column's, by the value field of either and a subtotal's function
  if (values.length && dataCols)
    rowLines.forEach((ra, i) => {
      if (ra.kind === 'item' || ra.kind === 'blank') return
      columnLines.forEach((ca, j) => {
        const v = ra.value ?? ca.value ?? 0
        const fn = ra.fn ?? ca.fn
        const x = result.value(ra.path, ca.path, v, fn)
        const numberFormat = values[v].numberFormat
        const row = top + height + i
        const col = left + width + j
        const cell: Label = { value: undefined }
        if (numberFormat !== undefined) cell.numberFormat = numberFormat
        if (x === null) {
          if (table.empty?.show !== false && table.empty?.text) cell.value = table.empty.text
        } else if (typeof x === 'object') {
          if (table.errors?.show && table.errors.text !== undefined) cell.value = table.errors.text
          else cell.error = x.error
        } else cell.value = x
        if (cell.value !== undefined || cell.error || cell.numberFormat !== undefined) put(row, col, cell)
      })
    })
  out.cells = [...cells.values()]
  return out
}
