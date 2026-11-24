import type { PivotTable, TableInput } from '../contract.js'
import { pivotFrame, placeParts, savedLines } from '../pivot/frame.js'
import { hasCell, parseRange, type Rect } from '../structure/area.js'
import { DEFAULT_TABLE_STYLE } from '../tables/check.js'

export interface Region {
  type: string
  box: Rect
}

type SizeOf = (type: string) => number
export interface StyledArea {
  regions(r: number, c: number, sizeOf: SizeOf): Region[]
  bounds: Rect[]
}

const rows = (x: Rect, r1: number, r2 = r1): Rect => ({ ...x, r1, r2 })
const cols = (x: Rect, c1: number, c2 = c1): Rect => ({ ...x, c1, c2 })
const at = (r: number, c: number): Rect => ({ r1: r, r2: r, c1: c, c2: c })

function stripe(out: Region[], x: Rect, index: number, axis: 'Row' | 'Column', sizeOf: SizeOf): void {
  const first = `first${axis}Stripe`
  const second = `second${axis}Stripe`
  const [low, high] = axis === 'Row' ? [x.r1, x.r2] : [x.c1, x.c2]
  const a = sizeOf(first)
  const b = sizeOf(second)
  const offset = (index - low) % (a + b)
  const start = index - offset + (offset < a ? 0 : a)
  const end = Math.min(high, start + (offset < a ? a : b) - 1)
  out.push({ type: offset < a ? first : second, box: axis === 'Row' ? rows(x, start, end) : cols(x, start, end) })
}

export function tableArea(table: TableInput): StyledArea {
  const x = parseRange(table.range)!
  const style = table.style ?? DEFAULT_TABLE_STYLE
  const data = rows(x, x.r1 + Number(table.headers !== false), x.r2 - Number(table.totals === true))
  return {
    bounds: [x],
    regions(r, c, sizeOf) {
      if (!hasCell(x, r, c)) return []
      const out: Region[] = [{ type: 'wholeTable', box: x }]
      if (hasCell(data, r, c)) {
        if (style.rowStripes) stripe(out, data, r, 'Row', sizeOf)
        if (style.columnStripes) stripe(out, data, c, 'Column', sizeOf)
      }
      if (style.lastColumn && c === x.c2) out.push({ type: 'lastColumn', box: cols(x, c) })
      if (style.firstColumn && c === x.c1) out.push({ type: 'firstColumn', box: cols(x, c) })
      if (table.totals && r === x.r2) {
        out.push({ type: 'totalRow', box: rows(x, r) })
        if (style.lastColumn && c === x.c2) out.push({ type: 'lastTotalCell', box: at(r, c) })
        if (style.firstColumn && c === x.c1) out.push({ type: 'firstTotalCell', box: at(r, c) })
      }
      if (table.headers !== false && r === x.r1) {
        out.push({ type: 'headerRow', box: rows(x, r) })
        // Excel applies the corner formats when the matching first/last column option is selected.
        if (style.lastColumn && c === x.c2) out.push({ type: 'lastHeaderCell', box: at(r, c) })
        if (style.firstColumn && c === x.c1) out.push({ type: 'firstHeaderCell', box: at(r, c) })
      }
      return out
    },
  }
}

const levels = ['first', 'second', 'third']
const levelOf = (depth: number) => levels[depth <= 0 ? 0 : depth % 2 ? 1 : 2]

/** The saved axes classify the visible rows and columns without laying out or scanning the cells of a PivotTable. */
export function pivotArea(table: PivotTable): StyledArea {
  const x = parseRange(table.location.range)!
  const style = table.style!
  const rowFields = table.rows ?? []
  const colFields = table.columns ?? []
  const rowLines = savedLines(table.rowLines, rowFields)
  const colLines = savedLines(table.columnLines, colFields)
  const dataTop = x.r1 + table.location.firstDataRow
  const dataLeft = x.c1 + table.location.firstDataColumn
  const grandRow = rowLines.at(-1)?.kind === 'grand'
  const grandCol = colLines.at(-1)?.kind === 'grand'
  const data = { r1: dataTop, c1: dataLeft, r2: x.r2 - Number(grandRow), c2: x.c2 - Number(grandCol) }
  const captions = new Map<number, Map<number, Region>>()
  const bounds = [x]
  const frame = { ...pivotFrame(table, rowFields, colFields), height: table.location.firstDataRow, width: table.location.firstDataColumn }
  if (table.pages?.length) {
    placeParts(table, frame, x.r1, x.c1, rowFields, colFields, [], [], (r, c, part, place) => {
      if (place.page === undefined) return
      const box = at(r, c)
      const region = { box, type: part.kind === 'page' ? 'pageFieldValues' : 'pageFieldLabels' }
      let row = captions.get(r)
      if (!row) captions.set(r, (row = new Map()))
      row.set(c, region)
      bounds.push(box)
    })
  }

  return {
    bounds,
    regions(r, c, sizeOf) {
      const page = captions.get(r)?.get(c)
      if (page) return [{ type: 'wholeTable', box: page.box }, page]
      if (!hasCell(x, r, c)) return []
      const out: Region[] = [{ type: 'wholeTable', box: x }]
      if (hasCell(data, r, c)) {
        if (style.rowStripes) stripe(out, data, r, 'Row', sizeOf)
        if (style.columnStripes) stripe(out, data, c, 'Column', sizeOf)
      }
      const row = rowLines[r - dataTop]
      const col = colLines[c - dataLeft]
      if (row?.kind === 'blank') out.push({ type: 'blankRow', box: rows(x, r) })
      if (row?.kind === 'subtotal') out.push({ type: `${levelOf(row.items.length - 1)}SubtotalRow`, box: rows(x, r) })
      if (col?.kind === 'subtotal') out.push({ type: `${levelOf(col.items.length - 1)}SubtotalColumn`, box: cols(x, c) })
      if (style.rowHeaders !== false) {
        if (c < dataLeft) out.push({ type: 'firstColumn', box: cols(x, x.c1, dataLeft - 1) })
        const depth = row ? row.items.length - 1 : -1
        if (row?.kind === 'data' && depth >= 0 && rowFields.slice(depth + 1).some((field) => field !== 'values')) {
          out.push({ type: `${levelOf(depth)}RowSubheading`, box: rows(x, r) })
        }
      }
      if (style.lastColumn !== false && grandCol && c === x.c2) out.push({ type: 'lastColumn', box: cols(x, c) })
      if (grandRow && r === x.r2) out.push({ type: 'totalRow', box: rows(x, r) })
      if (style.columnHeaders !== false && r < dataTop) {
        out.push({ type: 'headerRow', box: rows(x, x.r1, dataTop - 1) })
        const depth = r - x.r1 - frame.captionRow
        if (c >= dataLeft && depth >= 0 && depth < colFields.length && colFields[depth] !== 'values') {
          out.push({ type: `${levelOf(depth)}ColumnSubheading`, box: rows(cols(x, dataLeft, x.c2), r) })
        }
        if (c < dataLeft && style.rowHeaders !== false)
          out.push({ type: 'firstHeaderCell', box: cols(rows(x, x.r1, dataTop - 1), x.c1, dataLeft - 1) })
      }
      return out
    },
  }
}
