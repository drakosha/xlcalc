import { MAX_COLS, MAX_ROWS } from '../../model.js'
import { attrsOf, childrenOf, type Placed } from '../keep/xml.js'
import { element } from '../xml/scan.js'
import type { DrawingAnchor, DrawingMarker } from './types.js'

const EMU_PER_POINT = 12700
const POINTS_PER_PIXEL = 0.75

export function xmlNumber(text: string | null | undefined): number | undefined {
  if (!text?.trim()) return undefined
  const value = Number(text)
  return Number.isFinite(value) ? value : undefined
}

function marker(row: number | undefined, col: number | undefined, x: number | undefined, y: number | undefined, units: number): DrawingMarker | undefined {
  if (row === undefined || col === undefined || x === undefined || y === undefined) return
  if (!Number.isInteger(row) || row < 0 || row >= MAX_ROWS || !Number.isInteger(col) || col < 0 || col >= MAX_COLS) return
  return { row: row + 1, col: col + 1, x: x * units, y: y * units }
}

function drawingMarker(xml: string, child: Placed | undefined): DrawingMarker | undefined {
  if (!child) return
  const text = xml.slice(child.open, child.end)
  const at = (name: string) => xmlNumber(element(text, name)?.body)
  return marker(at('row'), at('col'), at('colOff'), at('rowOff'), 1 / EMU_PER_POINT)
}

export function anchorOf(xml: string, anchor: Placed): DrawingAnchor | undefined {
  const children = childrenOf(xml, anchor)
  const child = (name: string) => children.find((node) => node.local === name)
  const from = drawingMarker(xml, child('from'))
  if (anchor.local === 'twoCellAnchor') {
    const to = drawingMarker(xml, child('to'))
    if (!from || !to) return
    const edit = attrsOf(xml, anchor).editAs
    const editAs = edit === 'oneCell' || edit === 'absolute' ? edit : 'twoCell'
    return { kind: 'twoCell', from, to, editAs }
  }
  const extent = child('ext')
  if (!extent) return
  const extentAttrs = attrsOf(xml, extent)
  const cx = xmlNumber(extentAttrs.cx)
  const cy = xmlNumber(extentAttrs.cy)
  if (cx === undefined || cy === undefined || cx <= 0 || cy <= 0) return
  const size = { width: cx / EMU_PER_POINT, height: cy / EMU_PER_POINT }
  if (anchor.local === 'oneCellAnchor') return from && { kind: 'oneCell', from, size }
  const position = child('pos')
  if (!position) return
  const positionAttrs = attrsOf(xml, position)
  const x = xmlNumber(positionAttrs.x)
  const y = xmlNumber(positionAttrs.y)
  if (x !== undefined && y !== undefined) return { kind: 'absolute', bounds: { x: x / EMU_PER_POINT, y: y / EMU_PER_POINT, ...size } }
}

export function vmlAnchor(text: string | null | undefined): DrawingAnchor | undefined {
  const values = text?.split(',').map(xmlNumber)
  if (values?.length !== 8) return
  const [left, x, top, y, right, endX, bottom, endY] = values
  const from = marker(top, left, x, y, POINTS_PER_PIXEL)
  const to = marker(bottom, right, endX, endY, POINTS_PER_PIXEL)
  if (from && to) return { kind: 'twoCell', from, to, editAs: 'twoCell' }
}
