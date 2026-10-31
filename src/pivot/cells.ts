/**
 * What each cell of a PivotTable is (`pivotCells`, `pivotCellAt`), as Excel's `Range.PivotCell` tells it, from the
 * workbook's model alone (its place, its saved lines, its fields and value fields): no engine, so that a host's sheet
 * knows its menus and filter buttons without loading `xlcalc/pivot`.
 */
import type { BookInput, PivotAxisField, PivotCell, PivotCells, PivotItemRef, PivotTable } from '../contract.js'
import { caseKey } from '../collate.js'
import { InputError } from '../input-error.js'
import { checkString, object } from '../input/basic.js'
import { formatAddress } from '../model.js'
import { hasCell, parseRange } from '../structure/area.js'
import { pivotParts } from './area.js'
import { pivotFrame, placeParts, savedLines, type LineView, type PartAt, type PivotPart } from './frame.js'

type Sheets = Pick<BookInput, 'sheets'>

/** The path of a line's items up to its field `k`, the value fields left out (`value` names them). */
const pathOf = (fields: readonly PivotAxisField[], line: LineView | undefined, k = (line?.items.length ?? 0) - 1): PivotItemRef[] =>
  (line?.items ?? []).slice(0, k + 1).flatMap((item, j) => (fields[j] === 'values' ? [] : [{ field: fields[j], item }]))

/** What a PivotTable's every cell is, and its cells with filter buttons. */
function cellsOf(t: PivotTable): PivotCells {
  const cells: Record<string, PivotCell> = {}
  const buttons: string[] = []
  const put = (row: number, col: number, c: Omit<PivotCell, 'pivot'>) => (cells[formatAddress(row, col)] = { pivot: t.name, ...c })
  for (const x of pivotParts(t)) for (let r = x.r1; r <= x.r2; r++) for (let c = x.c1; c <= x.c2; c++) put(r, c, { kind: 'blank' })
  const x = parseRange(t.location.range)
  if (!x) return { cells, buttons }
  const rowFields = t.rows ?? []
  const colFields = t.columns ?? []
  const values = t.values ?? []
  const laid = !!t.rowLines && t.location.firstDataRow > 0 && (rowFields.length || colFields.length || values.length)
  // the frame as the place saved it: the rows above the values and the columns of the labels
  const frame = { ...pivotFrame(t, rowFields, colFields), height: t.location.firstDataRow, width: t.location.firstDataColumn }
  const rowLines = laid ? savedLines(t.rowLines, rowFields) : []
  const columnLines = laid ? savedLines(t.columnLines, colFields) : []
  const axisOf = (f: PivotAxisField, at: PartAt): PivotCell['axis'] => (at.page !== undefined ? 'pages' : at.rows !== undefined ? (at.rows ? 'rows' : 'columns') : rowFields.includes(f) ? 'rows' : 'columns')
  const first = (fields: readonly PivotAxisField[]) => fields.find((f) => f !== 'values')
  const kindOf = (part: PivotPart, at: PartAt): Omit<PivotCell, 'pivot'> => {
    // a label's path goes down to its own field, a total's to its line's
    const fields = at.rows ? rowFields : colFields
    const depth = part.kind === 'item' ? fields.indexOf(part.field) : undefined
    const path = at.line ? pathOf(fields, at.line, depth) : []
    const paths = path.length ? (at.rows ? { rowItems: path } : { columnItems: path }) : {}
    const value = at.line?.value === undefined ? {} : { value: at.line.value }
    switch (part.kind) {
      case 'field':
        return part.field === 'values' ? { kind: 'valuesField', axis: axisOf('values', at), field: 'values' } : { kind: 'field', axis: axisOf(part.field, at), field: part.field }
      case 'rowHeader':
      case 'columnHeader': {
        const rows = part.kind === 'rowHeader'
        const f = first(rows ? rowFields : colFields)
        return f === undefined ? { kind: 'blank' } : { kind: 'field', axis: rows ? 'rows' : 'columns', field: f }
      }
      case 'value':
        return { kind: 'valueName', value: part.value }
      case 'page':
        return { kind: 'pageItem', axis: 'pages', field: part.field }
      case 'grand':
        return { kind: 'grand', axis: at.rows ? 'rows' : 'columns', ...value, ...paths }
      case 'item':
        if (part.field === 'values') return { kind: 'valueName', axis: at.rows ? 'rows' : 'columns', value: part.item, ...paths }
        return { kind: 'item', axis: at.rows ? 'rows' : 'columns', field: part.field, item: part.item, ...paths }
      case 'subtotal':
        return { kind: 'subtotal', axis: at.rows ? 'rows' : 'columns', field: part.field, item: part.item, ...value, ...paths }
    }
  }
  placeParts(t, frame, x.r1, x.c1, rowFields, colFields, rowLines, columnLines, (row, col, part, at) => {
    const cell = kindOf(part, at)
    if (laid || at.page !== undefined) put(row, col, cell)
    // Excel's filter buttons: the headers of the rows and columns and the report filters' items
    if ((laid || at.page !== undefined) && (part.kind === 'rowHeader' || part.kind === 'columnHeader' || part.kind === 'page' || (part.kind === 'field' && part.field !== 'values' && at.page === undefined))) buttons.push(formatAddress(row, col))
  })
  // the values: each line with values against each column's, by the value field of either
  if (laid && values.length)
    rowLines.forEach((ra, i) => {
      if (ra.kind === 'blank') return
      columnLines.forEach((ca, j) => {
        const rowItems = pathOf(rowFields, ra)
        const columnItems = pathOf(colFields, ca)
        put(x.r1 + frame.height + i, x.c1 + frame.width + j, { kind: 'value', value: ra.value ?? ca.value ?? 0, ...(rowItems.length ? { rowItems } : {}), ...(columnItems.length ? { columnItems } : {}) })
      })
    })
  for (const a of buttons) cells[a] = { ...cells[a], button: true }
  return { cells, buttons }
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
