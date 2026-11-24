/** Positions of sheet artwork and text in a paginated region's coordinate systems. */
import { checkNumber } from '../../input/basic.js'
import { parseRange } from '../../structure/area.js'
import type { PrintDeviceGeometry, PrintPage, PrintRect, PrintSheetGeometry } from './types.js'

/** A part of a sheet rectangle on a page: in sheet points and in paper points. */
export interface PrintRectProjection {
  source: PrintRect
  target: PrintRect
}

function intersection(a: PrintRect, b: PrintRect): PrintRect | undefined {
  const x = Math.max(a.x, b.x)
  const y = Math.max(a.y, b.y)
  const right = Math.min(a.x + a.width, b.x + b.width)
  const bottom = Math.min(a.y + a.height, b.y + b.height)
  if (right > x && bottom > y) return { x, y, width: right - x, height: bottom - y }
}

/** An offset in sheet points projected inside its line, when printed line sizes differ from screen sizes. */
function inLine(point: number, edge: (line: number) => number, printed: (line: number) => number, first: number, last: number): number {
  if (point >= edge(last + 1)) return printed(last + 1)
  let low = first
  let high = last
  while (low < high) {
    const middle = Math.ceil((low + high) / 2)
    if (edge(middle) <= point) low = middle
    else high = middle - 1
  }
  const start = edge(low)
  const fraction = (point - start) / (edge(low + 1) - start)
  return printed(low) + fraction * (printed(low + 1) - printed(low))
}

/** Map the intersections of a rectangle with a page's sheet regions. The renderer applies the page-body clip. */
export function projectPrintRect(
  rect: PrintRect, page: PrintPage, geometry: PrintSheetGeometry, device?: PrintDeviceGeometry,
): PrintRectProjection[] {
  checkNumber(rect.x, 'rect.x', -Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER)
  checkNumber(rect.y, 'rect.y', -Number.MAX_SAFE_INTEGER, Number.MAX_SAFE_INTEGER)
  checkNumber(rect.width, 'rect.width', 0, Number.MAX_SAFE_INTEGER)
  checkNumber(rect.height, 'rect.height', 0, Number.MAX_SAFE_INTEGER)
  const rowEdge = device?.rowEdge
  const columnEdge = device?.columnEdge
  const printedRows = rowEdge && ((line: number) => rowEdge(line, page.scale))
  const printedColumns = columnEdge && ((line: number) => columnEdge(line, page.scale))
  const sheetRows = (line: number) => geometry.rows.edge(line)
  const sheetColumns = (line: number) => geometry.columns.edge(line)
  const projections: PrintRectProjection[] = []

  for (const region of page.regions) {
    const source = intersection(rect, region.source)
    if (!source) continue
    const range = parseRange(region.range)!
    const x = printedColumns
      ? (point: number) => region.target.x + inLine(point, sheetColumns, printedColumns, range.c1 + 1, range.c2 + 1) - printedColumns(range.c1 + 1)
      : (point: number) => region.target.x + (point - region.source.x) * region.target.width / region.source.width
    const y = printedRows
      ? (point: number) => region.target.y + inLine(point, sheetRows, printedRows, range.r1 + 1, range.r2 + 1) - printedRows(range.r1 + 1)
      : (point: number) => region.target.y + (point - region.source.y) * region.target.height / region.source.height
    const left = x(source.x)
    const top = y(source.y)
    const width = x(source.x + source.width) - left
    const height = y(source.y + source.height) - top
    if (width > 0 && height > 0) projections.push({ source, target: { x: left, y: top, width, height } })
  }
  return projections
}
