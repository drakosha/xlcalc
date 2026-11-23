import { escapeAttr } from '../xlsx/xml/text.js'
import { differential } from './dxf.js'
import { colorAttrs } from './fields.js'
import type { Part } from './parts.js'
import { records, type Payload } from './records.js'

export const HORIZONTAL = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed']
export const VERTICAL = ['top', 'center', 'bottom', 'justify', 'distributed']
export const LINES = ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed',
  'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
export const PATTERNS = ['none', 'solid', 'mediumGray', 'darkGray', 'lightGray', 'darkHorizontal', 'darkVertical',
  'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp',
  'lightGrid', 'lightTrellis', 'gray125', 'gray0625']
export const UNDERLINE: Readonly<Partial<Record<number, string>>> = { 1: 'single', 2: 'double', 33: 'singleAccounting', 34: 'doubleAccounting' }

const GROUPS = [
  [615, 616, 'numFmts'],
  [611, 612, 'fonts'],
  [603, 604, 'fills'],
  [613, 614, 'borders'],
  [626, 627, 'cellStyleXfs'],
  [617, 618, 'cellXfs'],
  [619, 620, 'cellStyles'],
  [505, 506, 'dxfs'],
  [473, 474, 'colors'],
  [565, 566, 'indexedColors'],
] as const

const OPENS = new Map<number, string>(GROUPS.map(([start, , tag]) => [start, tag]))
const CLOSES = new Map<number, string>(GROUPS.map(([, end, tag]) => [end, tag]))

export function styles({ bytes, path, xml, context }: Part): void {
  let fonts = 0
  let fills = 0
  let borders = 0
  let parents = 0
  let inCellXfs = false
  context.styles = 0
  xml.add('<styleSheet>')
  for (const { type, payload } of records(bytes, path)) {
    const open = OPENS.get(type)
    if (open) {
      if (type !== 473 && type !== 565) payload.u32()
      payload.done()
      if (type === 617) inCellXfs = true
      xml.add(`<${open}>`)
      continue
    }
    const close = CLOSES.get(type)
    if (close) {
      if (type === 618) inCellXfs = false
      payload.done()
      xml.add(`</${close}>`)
      continue
    }
    switch (type) {
      case 43: { // BrtFont
        const f = font(payload)
        xml.add(`<font><name val="${f.name}"/>${f.properties}</font>`)
        context.fonts.push(`<rFont val="${f.name}"/>${f.properties}`)
        fonts++
        break
      }
      case 44: {
        const id = payload.u16()
        const code = escapeAttr(payload.wide())
        payload.done()
        xml.add(`<numFmt numFmtId="${id}" formatCode="${code}"/>`)
        break
      }
      case 45: xml.add(fill(payload)); fills++; break
      case 46: xml.add(border(payload)); borders++; break
      case 47:
        xml.add(xf(payload, fonts, fills, borders, parents))
        if (inCellXfs) context.styles++
        else parents++
        break
      case 48:
        xml.add(namedStyle(payload))
        break
      case 475: {
        const rgb = [payload.u8(), payload.u8(), payload.u8()].map((n) => n.toString(16).padStart(2, '0')).join('').toUpperCase()
        payload.u8()
        payload.done()
        xml.add(`<rgbColor rgb="FF${rgb}"/>`)
        break
      }
      case 507:
        xml.add(`<dxf>${differential(payload)}</dxf>`)
        break
      case 1171:
        context.dxfs14.push(differential(payload, true))
        break
      case 508: {
        const count = payload.u32()
        const table = payload.nullableWide()
        const pivot = payload.nullableWide()
        payload.done()
        xml.add(`<tableStyles count="${count}"${table ? ` defaultTableStyle="${escapeAttr(table)}"` : ''}${pivot ? ` defaultPivotStyle="${escapeAttr(pivot)}"` : ''}>`)
        break
      }
      case 509: xml.add('</tableStyles>'); break
      case 510: {
        const flags = payload.u16()
        const count = payload.u32()
        const name = payload.nullableWide()
        payload.done()
        xml.add(`<tableStyle name="${escapeAttr(name ?? '')}" table="${(flags >>> 2) & 1}" pivot="${(flags >>> 1) & 1}" count="${count}"/>`)
        break
      }
    }
  }
  xml.add('</styleSheet>')
}

function font(payload: Payload): { name: string; properties: string } {
  const size = payload.u16() / 20
  const flags = payload.u16()
  const bold = payload.u16() >= 700
  const align = payload.u16()
  const underline = UNDERLINE[payload.u8()]
  const family = payload.u8()
  const charset = payload.u8()
  payload.u8()
  const color = colorAttrs(payload)
  const scheme = payload.u8()
  const name = escapeAttr(payload.wide())
  payload.done()
  let properties = `<sz val="${size}"/>`
  if (bold) properties += '<b/>'
  for (const [bit, tag] of [[2, 'i'], [8, 'strike'], [16, 'outline'], [32, 'shadow'], [64, 'condense'], [128, 'extend']] as const) {
    if (flags & bit) properties += `<${tag}/>`
  }
  if (underline) properties += `<u val="${underline}"/>`
  if (align) properties += `<vertAlign val="${align === 1 ? 'superscript' : 'subscript'}"/>`
  if (family) properties += `<family val="${family}"/>`
  if (charset) properties += `<charset val="${charset}"/>`
  if (scheme) properties += `<scheme val="${scheme === 1 ? 'major' : 'minor'}"/>`
  if (color !== 'auto="1"') properties += `<color ${color}/>`
  return { name, properties }
}

function fill(payload: Payload): string {
  const pattern = payload.u32()
  const fg = colorAttrs(payload)
  const bg = colorAttrs(payload)
  const gradientType = payload.u32()
  const degree = payload.f64()
  const left = payload.f64()
  const right = payload.f64()
  const top = payload.f64()
  const bottom = payload.f64()
  const stops = payload.u32()
  if (stops * 16 !== payload.remaining) payload.fail(`invalid gradient stop count ${stops}`)
  if (pattern !== 40) {
    if (!PATTERNS[pattern]) payload.fail(`unknown fill pattern ${pattern}`)
    payload.done()
    return `<fill><patternFill patternType="${PATTERNS[pattern]}"><fgColor ${fg}/><bgColor ${bg}/></patternFill></fill>`
  }
  const type = gradientType === 0 ? `degree="${degree}"` : `type="path" left="${left}" right="${right}" top="${top}" bottom="${bottom}"`
  let fill = `<fill><gradientFill ${type}>`
  for (let i = 0; i < stops; i++) {
    const color = colorAttrs(payload)
    const position = payload.f64()
    fill += `<stop position="${position}"><color ${color}/></stop>`
  }
  return fill + '</gradientFill></fill>'
}

function border(payload: Payload): string {
  const flags = payload.u8()
  const diagonal = (flags & 1 ? ' diagonalDown="1"' : '') + (flags & 2 ? ' diagonalUp="1"' : '')
  let xml = `<border${diagonal}>`
  for (const side of ['top', 'bottom', 'left', 'right', 'diagonal']) {
    const style = LINES[payload.u8()]
    payload.u8()
    const color = colorAttrs(payload)
    if (style && style !== 'none') xml += `<${side} style="${style}"><color ${color}/></${side}>`
  }
  payload.done()
  return xml + '</border>'
}

function xf(payload: Payload, fonts: number, fills: number, borders: number, parents: number): string {
  const parent = payload.u16()
  const format = payload.u16()
  const font = payload.u16()
  const fill = payload.u16()
  const border = payload.u16()
  const rotation = payload.u8()
  const indent = payload.u8()
  const flags = payload.u32()
  payload.done()
  if (font >= fonts || fill >= fills || border >= borders || (parent !== 0xffff && parent >= parents)) {
    payload.fail('cell format index out of bounds')
  }
  const horizontal = HORIZONTAL[flags & 7]
  const vertical = VERTICAL[(flags >>> 3) & 7] ?? 'bottom'
  const wrap = (flags >>> 6) & 1
  const justifyLastLine = (flags >>> 7) & 1
  const shrink = (flags >>> 8) & 1
  const readingOrder = (flags >>> 10) & 3
  const locked = (flags >>> 12) & 1
  const hidden = (flags >>> 13) & 1
  const prefix = flags & 0x8000 ? ' quotePrefix="1"' : ''
  const pivot = flags & 0x4000 ? ' pivotButton="1"' : ''
  const style = parent === 0xffff ? '' : ` xfId="${parent}"`
  const attrs = ['NumberFormat', 'Font', 'Alignment', 'Border', 'Fill', 'Protection'].map((key, i) => {
    const bit = (flags >>> (16 + i)) & 1
    return ` apply${key}="${parent === 0xffff ? 1 - bit : bit}"`
  }).join('')
  return `<xf numFmtId="${format}" fontId="${font}" fillId="${fill}" borderId="${border}"${style}${prefix}${pivot}${attrs}>`
    + `<alignment horizontal="${horizontal}" vertical="${vertical}" wrapText="${wrap}" justifyLastLine="${justifyLastLine}" shrinkToFit="${shrink}"`
    + ` readingOrder="${readingOrder}" textRotation="${rotation}" indent="${indent}"/><protection locked="${locked}" hidden="${hidden}"/></xf>`
}

function namedStyle(payload: Payload): string {
  const parent = payload.u32()
  const flags = payload.u16()
  const builtin = payload.u8()
  const level = payload.u8()
  const name = escapeAttr(payload.nullableWide() ?? '')
  payload.done()
  return `<cellStyle name="${name}" xfId="${parent}"${flags & 1 ? ` builtinId="${builtin}" iLevel="${level}"` : ''}${flags & 2 ? ' hidden="1"' : ''}${flags & 4 ? ' customBuiltin="1"' : ''}/>`
}
