import { formatAddress, MAX_COLS, MAX_ROWS } from '../../model.js'
import { attrsOf, childrenOf, rootOf } from '../keep/xml.js'
import type { ReadXlsxOptions } from '../types.js'
import { attr, attributes, attrsAt, element, elements, flag } from '../xml/scan.js'
import { decodeXml } from '../xml/text.js'
import { anchorOf, vmlAnchor, xmlNumber } from './anchors.js'
import { drawingParts, openDrawingPackage, pictureTypes, type DrawingPart } from './package.js'
import type { DrawingReadIssue, NoteDrawing, PictureDrawing, XlsxDrawing, XlsxDrawings } from './types.js'

export function frames(part: DrawingPart, issues: DrawingReadIssue[], onlyCharts: boolean): (PictureDrawing | Extract<XlsxDrawing, { kind: 'chart' }>)[] {
  const { xml, sheet } = part
  const root = rootOf(xml)
  if (!root) return []
  const drawings: (PictureDrawing | Extract<XlsxDrawing, { kind: 'chart' }>)[] = []
  const report = (reason: DrawingReadIssue['reason']) => issues.push({ sheet: sheet.name, part: part.part, reason })
  for (const [index, item] of childrenOf(xml, root).entries()) {
    if (item.local === 'extLst') continue
    if (!['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'].includes(item.local)) {
      if (!onlyCharts) report('unsupportedDrawing')
      continue
    }
    const children = childrenOf(xml, item)
    const object = children.find((node) => node.local === 'pic' || node.local === 'graphicFrame')
    const content = object && xml.slice(object.open, object.end)
    const kind = object?.local === 'pic' ? 'picture' : content && element(content, 'chart') ? 'chart' : undefined
    if (onlyCharts && kind !== 'chart') continue
    if (!kind || !content) {
      report('unsupportedDrawing')
      continue
    }
    const anchor = anchorOf(xml, item)
    if (!anchor) {
      report('invalidAnchor')
      continue
    }
    const info = attrsAt(content, 'cNvPr')
    const client = children.find((node) => node.local === 'clientData')
    const frame = {
      id: sheet.part + '#' + part.part + '#' + index, sheet: sheet.name, sheetKind: sheet.kind, anchor,
      ...(info.name === undefined ? {} : { name: decodeXml(info.name) }),
      hidden: flag(info.hidden), printable: !client || flag(attrsOf(xml, client).fPrintsWithSheet, true),
    }
    const image = kind === 'picture' ? attrsAt(content, 'blip') : undefined
    const id = image ? attr(image, 'embed') ?? attr(image, 'link') : attr(attrsAt(content, 'chart'), 'id')
    const rel = part.rels.find((link) => link.id === id && link.kind === (image ? 'image' : 'chart'))
    if (!rel || rel.external) {
      report(rel?.external && image ? 'externalImage' : 'missingPart')
      continue
    }
    if (!image) {
      drawings.push({ ...frame, kind: 'chart', part: rel.target })
      continue
    }
    const transform = attrsAt(content, 'xfrm')
    const rotation = xmlNumber(transform.rot)
    const crop = attrsAt(content, 'srcRect')
    const fraction = (name: string) => (xmlNumber(crop[name]) ?? 0) / 100000
    drawings.push({
      ...frame, kind: 'picture', part: rel.target,
      ...(info.descr === undefined ? {} : { alt: decodeXml(info.descr) }),
      ...(rotation ? { rotation: rotation / 60000 } : {}),
      ...(flag(transform.flipH) ? { flipX: true } : {}),
      ...(flag(transform.flipV) ? { flipY: true } : {}),
      ...(Object.keys(crop).length ? { crop: { left: fraction('l'), top: fraction('t'), right: fraction('r'), bottom: fraction('b') } } : {}),
    })
  }
  return drawings
}

function noteDrawings(part: DrawingPart, issues: DrawingReadIssue[]): NoteDrawing[] {
  const out: NoteDrawing[] = []
  for (const [index, shape] of elements(part.xml, 'shape').entries()) {
    const client = element(shape.body ?? '', 'ClientData')
    if (!client || attributes(client.attrs).ObjectType !== 'Note') continue
    const body = client.body ?? ''
    const row = xmlNumber(element(body, 'Row')?.body)
    const col = xmlNumber(element(body, 'Column')?.body)
    const anchor = vmlAnchor(element(body, 'Anchor')?.body)
    if (!anchor || row === undefined || col === undefined || !Number.isInteger(row) || !Number.isInteger(col) ||
        row < 0 || row >= MAX_ROWS || col < 0 || col >= MAX_COLS) {
      issues.push({ sheet: part.sheet.name, part: part.part, reason: 'invalidAnchor' })
      continue
    }
    const attrs = attributes(shape.attrs)
    const cell = formatAddress(row, col)
    const print = element(body, 'PrintObject')
    const visible = element(body, 'Visible')
    const hidden = visible ? /^(false|0)$/i.test(visible.body?.trim() ?? '')
      : !/visibility\s*:\s*visible\b/i.test(attrs.style ?? '')
    out.push({
      id: part.sheet.part + '#' + part.part + '#' + index, sheet: part.sheet.name, sheetKind: part.sheet.kind, kind: 'note', cell, anchor, hidden,
      printable: !print || !/^(false|0)$/i.test(print.body?.trim() ?? ''),
      ...(attrs.fillcolor ? { fill: decodeXml(attrs.fillcolor) } : {}),
      ...(attrs.strokecolor ? { stroke: decodeXml(attrs.strokecolor) } : {}),
    })
  }
  return out
}

export async function readXlsxDrawings(data: Uint8Array | ArrayBuffer, opts?: ReadXlsxOptions): Promise<XlsxDrawings> {
  const pkg = await openDrawingPackage(data, opts)
  const issues: DrawingReadIssue[] = []
  const drawings: XlsxDrawing[] = []
  const parts = new Set(pkg.archive.names.map((part) => part.toLowerCase()))
  const pictures = new Set<string>()
  let contentType: Awaited<ReturnType<typeof pictureTypes>> | undefined
  for await (const part of drawingParts(pkg, issues, true)) {
    if (part.kind === 'vmlDrawing') {
      for (const drawing of noteDrawings(part, issues)) drawings.push(drawing)
      continue
    }
    for (const drawing of frames(part, issues, false)) {
      if (!parts.has(drawing.part.toLowerCase())) {
        issues.push({ sheet: drawing.sheet, part: drawing.part, reason: 'missingPart' })
        continue
      }
      if (drawing.kind === 'picture') {
        contentType ??= await pictureTypes(pkg.archive)
        drawing.contentType = contentType(drawing.part)
        pictures.add(drawing.part.toLowerCase())
      }
      drawings.push(drawing)
    }
  }
  const images = new Map<string, Promise<Uint8Array | undefined>>()
  return {
    drawings, issues,
    async readImage(part) {
      const key = part.toLowerCase()
      if (!pictures.has(key)) return
      // The same image can be requested by several pages. Inflate it once against the archive's byte budget.
      let loaded = images.get(key)
      if (!loaded) {
        loaded = pkg.archive.bytes(part)
        images.set(key, loaded)
      }
      return loaded
    },
  }
}
