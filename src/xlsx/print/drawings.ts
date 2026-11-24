import type { DrawingAnchor, DrawingMarker, XlsxDrawing } from '../drawings/types.js'
import { parseRange } from '../../structure/area.js'
import { projectPrintRect, type PrintRectProjection } from './project.js'
import type { PrintDeviceGeometry, PrintPage, PrintRect, PrintSetup, PrintSheetGeometry } from './types.js'

/** A drawing on a page: its full box in sheet points and its projected regions. */
export interface PrintDrawingProjection {
  drawing: XlsxDrawing
  /** The unrotated bounds in sheet points. */
  bounds: PrintRect
  /** Positions and clips on this page. A note is projected as a whole box. */
  regions: PrintRectProjection[]
}

/** The sheet's note print mode and the device geometry for `projectPrintDrawings`. */
export interface PrintDrawingOptions {
  setup?: Pick<PrintSetup, 'cellComments'>
  device?: PrintDeviceGeometry
}

/** A drawing anchor's rectangle in sheet points, for print or for display on screen; undefined when it has no extent. */
export function printDrawingBounds(anchor: DrawingAnchor, geometry: PrintSheetGeometry): PrintRect | undefined {
  return drawingBounds(anchor, geometry, 1)
}

function drawingBounds(anchor: DrawingAnchor, geometry: PrintSheetGeometry, scale: number, device?: PrintDeviceGeometry): PrintRect | undefined {
  if (anchor.kind === 'absolute') {
    const { x, y, width, height } = anchor.bounds
    return { x: x * scale, y: y * scale, width: width * scale, height: height * scale }
  }
  const mark = ({ row, col, x, y }: DrawingMarker) => {
    if (row < 1 || row > geometry.rows.count + 1 || col < 1 || col > geometry.columns.count + 1) return
    return {
      x: (device?.columnEdge?.(col, scale) ?? geometry.columns.edge(col) * scale) + x * scale,
      y: (device?.rowEdge?.(row, scale) ?? geometry.rows.edge(row) * scale) + y * scale,
    }
  }
  const from = mark(anchor.from)
  if (!from) return
  if (anchor.kind === 'oneCell') return { ...from, width: anchor.size.width * scale, height: anchor.size.height * scale }
  const to = mark(anchor.to)
  if (to && to.x > from.x && to.y > from.y) return { ...from, width: to.x - from.x, height: to.y - from.y }
}

function rotatedBounds(bounds: PrintRect, rotation: number): PrintRect {
  const angle = rotation * Math.PI / 180
  const sin = Math.abs(Math.sin(angle))
  const cos = Math.abs(Math.cos(angle))
  const width = cos * bounds.width + sin * bounds.height
  const height = sin * bounds.width + cos * bounds.height
  return { x: bounds.x + (bounds.width - width) / 2, y: bounds.y + (bounds.height - height) / 2, width, height }
}

function noteRegions(
  drawing: Extract<XlsxDrawing, { kind: 'note' }>, bounds: PrintRect,
  page: PrintPage, geometry: PrintSheetGeometry, device?: PrintDeviceGeometry,
): PrintRectProjection[] {
  const { scale } = page
  const box = drawingBounds(drawing.anchor, geometry, scale, device)
  if (!box) return []
  const regions: PrintRectProjection[] = []
  for (const region of page.regions) {
    const range = parseRange(region.range)!
    const originX = device?.columnEdge?.(range.c1 + 1, scale) ?? region.source.x * scale
    const originY = device?.rowEdge?.(range.r1 + 1, scale) ?? region.source.y * scale
    const x = box.x - originX
    const y = box.y - originY
    if (x >= region.target.width || x + box.width <= 0 || y >= region.target.height || y + box.height <= 0) continue
    // Excel draws the whole box where it intersects the printed grid, even on a page without its cell.
    const target = { ...box, x: region.target.x + x, y: region.target.y + y }
    if (!regions.some((previous) => previous.target.x === target.x && previous.target.y === target.y)) {
      regions.push({ source: bounds, target })
    }
  }
  return regions
}

/** Lay out drawings of the page's sheet in their stored drawing order. */
export function projectPrintDrawings(
  drawings: readonly XlsxDrawing[], page: PrintPage, geometry: PrintSheetGeometry, options: PrintDrawingOptions = {},
): PrintDrawingProjection[] {
  const out: PrintDrawingProjection[] = []
  for (const drawing of drawings) {
    if (drawing.hidden || !drawing.printable) continue
    if ('sheet' in page && page.sheet !== drawing.sheet) continue
    if (drawing.kind === 'note' && options.setup?.cellComments !== 'asDisplayed') continue
    const bounds = printDrawingBounds(drawing.anchor, geometry)
    if (!bounds) continue
    const regions = drawing.kind === 'note'
      ? noteRegions(drawing, bounds, page, geometry, options.device)
      : projectPrintRect(drawing.kind === 'picture' && drawing.rotation ? rotatedBounds(bounds, drawing.rotation) : bounds,
        page, geometry, options.device)
    if (regions.length) out.push({ drawing, bounds, regions })
  }
  return out
}
