/**
 * What each cell of a PivotTable is (`pivotCells`, `pivotCellAt`), as Excel's `Range.PivotCell` tells it, from the
 * workbook's model alone (its place, its saved lines, its fields and value fields): no engine, so that a host's sheet
 * knows its menus and filter buttons without loading `xlcalc/pivot`.
 */
import type { PivotAxisField, PivotCell, PivotCells, PivotItemRef, PivotTable, SheetInput } from '../contract.js'
import { caseKey } from '../collate.js'
import { InputError } from '../input-error.js'
import { checkString, object } from '../input/basic.js'
import { formatAddress } from '../model.js'
import { hasCell, parseRange } from '../structure/area.js'
import { pivotParts } from './area.js'
import { pivotFrame, placeParts, savedLines, type LineView } from './frame.js'

/** A workbook or `getLayout()`'s, of which the PivotTables are read. */
type Sheets = { sheets: readonly Pick<SheetInput, 'name' | 'pivotTables'>[] }

/** The path of a line's items up to its field `k`, the value fields left out (`value` names them). */
const pathOf = (fields: readonly PivotAxisField[], line: LineView | undefined, k = (line?.items.length ?? 0) - 1): PivotItemRef[] =>
  (line?.items ?? []).slice(0, k + 1).flatMap((item, j) => (fields[j] === 'values' ? [] : [{ field: fields[j], item }]))

/** What a PivotTable's every cell is, and its cells with filter buttons. */
function cellsOf(t: PivotTable): PivotCells {
  const cells: Record<string, PivotCell> = {}
  const buttons: string[] = []
  const put = (row: number, col: number, c: Omit<PivotCell, 'pivot'>) => (cells[formatAddress(row, col)] = { pivot: t.name, ...c })
  const x = parseRange(t.location.range)
  if (!x) return { cells, buttons }
  // the table's cells; of the Filters area only its fields' names and items, the cells between them none of its own
  for (let r = x.r1; r <= x.r2; r++) for (let c = x.c1; c <= x.c2; c++) put(r, c, { kind: 'blank' })
  const rowFields = t.rows ?? []
  const colFields = t.columns ?? []
  const values = t.values ?? []
  const laid = !!t.rowLines && t.location.firstDataRow > 0 && (rowFields.length || colFields.length || values.length)
  // the frame as the place saved it: the rows above the values and the columns of the labels
  const frame = { ...pivotFrame(t, rowFields, colFields), height: t.location.firstDataRow, width: t.location.firstDataColumn }
  const { height, width, captionRow, labelCol } = frame
  const [top, left] = [x.r1, x.c1]
  const rowLines = laid ? savedLines(t.rowLines, rowFields) : []
  const columnLines = laid ? savedLines(t.columnLines, colFields) : []
  // the Filters area and the headers: a field's name, Σ Values', a value field's (`Total` in the classic layout is the grand total's); Row and Column Labels name no field
  placeParts(t, frame, top, left, rowFields, colFields, [], [], (row, col, part, at) => {
    if (!laid && at.page === undefined) return
    const axis = at.page !== undefined ? 'pages' : part.kind === 'field' && colFields.includes(part.field) ? 'columns' : 'rows'
    if (part.kind === 'page') put(row, col, { kind: 'pageItem', axis, field: part.field })
    else if (part.kind === 'field') put(row, col, part.field === 'values' ? { kind: 'valuesField', axis, field: 'values' } : { kind: 'field', axis, field: part.field })
    else if (part.kind === 'value') put(row, col, part.total ? { kind: 'grand', axis: 'columns' } : { kind: 'valueName', value: part.value })
    // Excel's buttons (`pivotButton`): the fields' names and the headers of the rows and columns, the report filters' names
    if (part.kind === 'rowHeader' || part.kind === 'columnHeader' || part.kind === 'field') buttons.push(formatAddress(row, col))
  })
  if (laid) {
    // a single value field's name at the corner over column labels, whether a row field shows it there or not, with a
    // button beside row fields; with no field on columns and the value fields on the rows, the head of their column is the
    // grand total's
    if (captionRow && values.length === 1) {
      put(top, left, { kind: 'valueName', value: 0 })
      if (rowFields.length) buttons.push(formatAddress(top, left))
    }
    if (!colFields.length && rowFields.includes('values')) put(top + height - 1, left + width, { kind: 'grand', axis: 'columns' })
    const one = values.length ? (rowFields.includes('values') || colFields.includes('values') ? undefined : 0) : undefined
    // every cell of a line's labels is of its item on the field of that row or column the line reaches, else of the
    // line's own; a total's cells are the total's, a blank line's blank
    const label = (fields: readonly PivotAxisField[], a: LineView, k: number | undefined, axis: 'rows' | 'columns'): Omit<PivotCell, 'pivot'> => {
      const d = a.items.length - 1
      const value = (a.value ?? one) === undefined ? {} : { value: a.value ?? one }
      const path = (to: number) => {
        const p = pathOf(fields, a, to)
        return p.length ? { [axis === 'rows' ? 'rowItems' : 'columnItems']: p } : {}
      }
      if (a.kind === 'grand') return { kind: 'grand', axis, ...(values.length ? { value: a.value ?? 0 } : {}) }
      if (a.kind === 'blank') return { kind: 'blank' }
      if (a.kind === 'subtotal') return { kind: 'subtotal', axis, field: fields[d], item: a.items[d], ...value, ...path(d) }
      const at = k === undefined || k > d ? d : k
      if (at < 0) return values.length === 1 ? { kind: 'valueName', value: 0 } : { kind: 'blank' }
      const f = fields[at]
      return f === 'values' ? { kind: 'valueName', axis, value: a.items[at], ...path(at) } : { kind: 'item', axis, field: f, item: a.items[at], ...path(at) }
    }
    rowLines.forEach((a, i) => {
      for (let c = 0; c < width; c++) {
        // the deepest field of this column
        let k: number | undefined
        labelCol.forEach((lc, j) => lc === c && j <= a.items.length - 1 && (k = j))
        if (k === undefined && labelCol.includes(c)) k = labelCol.indexOf(c)
        put(top + height + i, left + c, label(rowFields, a, k, 'rows'))
      }
    })
    columnLines.forEach((a, j) => {
      for (let r = captionRow; r < height; r++) if (r - captionRow < colFields.length || a.kind === 'grand') put(top + r, left + width + j, label(colFields, a, r - captionRow, 'columns'))
    })
    // the values, of no value field too: each line against each column's, of the value field of either
    rowLines.forEach((ra, i) => {
      if (ra.kind === 'blank' || top + height + i > x.r2) return
      columnLines.forEach((ca, j) => {
        if (left + width + j > x.c2) return
        const rowItems = pathOf(rowFields, ra)
        const columnItems = pathOf(colFields, ca)
        const value = ra.value ?? ca.value ?? one
        put(top + height + i, left + width + j, { kind: 'value', ...(value === undefined ? {} : { value }), ...(rowItems.length ? { rowItems } : {}), ...(columnItems.length ? { columnItems } : {}) })
      })
    })
  }
  const unique = [...new Set(buttons)]
  for (const a of unique) cells[a] = { ...cells[a], button: true }
  return { cells, buttons: unique }
}

/** The sheet of `book` by its name, compared as Excel compares names. */
function sheetOf(book: Sheets, sheet: string) {
  const b = object(book, 'book', 'a workbook such as getLayout() gives')
  if (!Array.isArray(b.sheets)) throw new InputError('book.sheets', 'expected the sheets of a workbook')
  const key = caseKey(checkString(sheet, 'sheet'))
  return book.sheets.find((s) => caseKey(s.name) === key) ?? null
}

/**
 * What every cell of the PivotTable `pivot` on `sheet` is, by address, its Filters area included, and the cells with
 * Excel's filter buttons; null when the sheet has no such PivotTable. `book` is a workbook or `getLayout()`'s.
 */
export function pivotCells(book: Sheets, sheet: string, pivot: string): PivotCells | null {
  const s = sheetOf(book, sheet)
  const key = caseKey(checkString(pivot, 'pivot'))
  const t = s?.pivotTables?.find((p) => caseKey(p.name) === key)
  return t ? cellsOf(t) : null
}

/** What the cell `cell` of `sheet` is in the PivotTable whose cells hold it (`pivotCells`); null in no PivotTable. */
export function pivotCellAt(book: Sheets, sheet: string, cell: string): PivotCell | null {
  const s = sheetOf(book, sheet)
  const x = parseRange(checkString(cell, 'cell'))
  if (!x || x.r1 !== x.r2 || x.c1 !== x.c2) throw new InputError('cell', `expected the address of a cell such as "B4", got ${JSON.stringify(cell)}`)
  const t = s?.pivotTables?.find((p) => pivotParts(p).some((y) => hasCell(y, x.r1, x.c1)))
  return t ? (cellsOf(t).cells[formatAddress(x.r1, x.c1)] ?? null) : null
}
