import { escapeAttr } from '../xlsx/xml/text.js'
import { differential } from '../xlsb/dxf.js'
import type { Payload } from '../xlsb/records.js'
import { HORIZONTAL, LINES, PATTERNS, UNDERLINE, VERTICAL } from '../xlsb/styles.js'
import type { Xml } from '../xlsb/xml.js'
import type { LegacyEncoding, LegacyVersion } from './encoding.js'
import { unicode } from './records.js'
import { TableStyles } from './table-styles.js'

export interface Font {
  name: string
  size: number
  props: string
}

export function color(index: number): string {
  return index === 0x7fff ? 'auto="1"' : `indexed="${index}"`
}

function font(payload: Payload, encoding?: LegacyEncoding): Font {
  const size = payload.u16() / 20
  const flags = payload.u16()
  const early = encoding !== undefined && encoding.version < 5
  const index = encoding?.version === 2 ? 0x7fff : payload.u16()
  const bold = early ? Boolean(flags & 1) : payload.u16() >= 700
  const vert = early ? 0 : payload.u16()
  const underline = early ? flags & 4 ? 'single' : undefined : UNDERLINE[payload.u8()]
  const family = early ? 0 : payload.u8()
  const charset = early ? 0 : payload.u8()
  if (!early) payload.skip(1)
  const name = encoding ? encoding.text(payload, payload.u8()) : unicode(payload, payload.u8())
  let props = `<sz val="${size}"/>`
  if (bold) props += '<b/>'
  for (const [flag, tag] of [[2, 'i'], [8, 'strike'], [16, 'outline'], [32, 'shadow'], [64, 'condense'], [128, 'extend']] as const) {
    if (flags & flag) props += `<${tag}/>`
  }
  if (underline) props += `<u val="${underline}"/>`
  if (vert) props += `<vertAlign val="${vert === 1 ? 'superscript' : 'subscript'}"/>`
  if (family) props += `<family val="${family}"/>`
  if (charset) props += `<charset val="${charset}"/>`
  if (index !== 0x7fff) props += `<color ${color(index)}/>`
  return { name, size, props }
}

export function colorAttrs(kind: number, raw: number, tint: number): string {
  const suffix = tint ? ` tint="${tint}"` : ''
  if (kind === 1) return color(raw) + suffix
  if (kind === 2) {
    const digits = [raw >>> 24, raw & 255, (raw >>> 8) & 255, (raw >>> 16) & 255]
    return `rgb="${digits.map((n) => n.toString(16).padStart(2, '0')).join('').toUpperCase()}"` + suffix
  }
  if (kind === 3) return `theme="${raw}"` + suffix
  return 'auto="1"'
}

export function fullColor(payload: Payload): string {
  const type = payload.u32()
  const value = payload.u32()
  return colorAttrs(type, value, payload.f64())
}

export function extensions(payload: Payload, count: number): Map<number, string | number> {
  const props = new Map<number, string | number>()
  for (let i = 0; i < count; i++) {
    const type = payload.u16()
    const size = payload.u16()
    const data = payload.section(size - 4)
    if ((type >= 4 && type <= 11 && type !== 6) || type === 13) {
      const kind = data.u16()
      const tint = data.i16()
      props.set(type, colorAttrs(kind, data.u32(), Math.max(-1, tint / 32767)))
    } else if (type === 15) props.set(type, data.remaining === 1 ? data.u8() : data.u16())
    else if (type === 6) {
      const mode = data.u32()
      const angle = data.f64()
      const left = data.f64(), right = data.f64(), top = data.f64(), bottom = data.f64()
      const attrs = mode === 0 ? `degree="${angle}"` : `type="path" left="${left}" right="${right}" top="${top}" bottom="${bottom}"`
      let gradient = `<fill><gradientFill ${attrs}>`
      const stops = data.u32()
      for (let j = 0; j < stops; j++) {
        const kind = data.u16()
        const raw = data.u32()
        const position = data.f64()
        const tint = data.f64()
        gradient += `<stop position="${position}"><color ${colorAttrs(kind, raw, tint)}/></stop>`
      }
      props.set(type, gradient + '</gradientFill></fill>')
    }
  }
  return props
}

interface Xf {
  font: number
  format: number
  protection: number
  alignment: number
  rotation: number
  indent: number
  borders: number
  moreBorders: number
  fill: number
  ext?: Map<number, string | number>
}

const ROTATION = [0, 255, 90, 180]

function xf2(font: number, format: number, protection: number, style: number): Xf {
  const borders = ((style & 8) >>> 3) | ((style & 16) << 0) | ((style & 32) << 3) | ((style & 64) << 6)
  return {
    font, format, protection, alignment: (style & 7) | 32, rotation: 0, indent: 0, borders,
    moreBorders: style & 128 ? 2 << 26 : 0, fill: 1 << 7,
  }
}

function legacyXf(payload: Payload, version: LegacyVersion): Xf {
  const font = payload.u8()
  if (version === 2) {
    payload.skip(1)
    const formatAndProtection = payload.u8()
    return xf2(font, formatAndProtection & 63, formatAndProtection >>> 6, payload.u8())
  }
  const format = payload.u8()
  const protection = payload.u16() & 7
  const align = payload.u8()
  payload.skip(1)
  const area = payload.u16()
  const lines = payload.u32()
  const alignment = (align & 15) | (version === 3 ? 32 : align & 48)
  const rotation = version === 3 ? 0 : ROTATION[align >>> 6]
  const borders = ((lines >>> 8) & 7) | (((lines >>> 24) & 7) << 4) | ((lines & 7) << 8)
    | (((lines >>> 16) & 7) << 12) | (((lines >>> 11) & 31) << 16) | (((lines >>> 27) & 31) << 23)
  const moreBorders = ((lines >>> 3) & 31) | (((lines >>> 19) & 31) << 7) | ((area & 63) << 26)
  const fill = ((area >>> 6) & 31) | ((area >>> 11) << 7)
  return { font, format, protection, alignment, rotation, indent: 0, borders, moreBorders, fill }
}

function xf(payload: Payload, old: boolean): Xf {
  const font = payload.u16()
  const format = payload.u16()
  const protection = payload.u16()
  const alignment = payload.u8()
  const rotation = payload.u8()
  if (old) {
    const fill = payload.u32()
    const border = payload.u32()
    // BIFF5 encodes the same four sides and fill in a different bit order.
    const borders = ((border >>> 3) & 7) | (((border >>> 6) & 7) << 4) | ((border & 7) << 8) |
      (((fill >>> 22) & 7) << 12) | (((border >>> 16) & 127) << 16) | (((border >>> 23) & 127) << 23)
    const moreBorders = ((border >>> 9) & 127) | ((fill >>> 25) << 7) | (((fill >>> 16) & 63) << 26)
    return { font, format, protection, alignment, rotation: ROTATION[rotation & 3], indent: 0, borders,
      moreBorders, fill: fill & 0x3fff }
  }
  const indent = payload.u8()
  payload.skip(1)
  return { font, format, protection, alignment, rotation, indent, borders: payload.u32(), moreBorders: payload.u32(), fill: payload.u16() }
}

function alignment(xf: Xf): string {
  const horizontal = HORIZONTAL[xf.alignment & 7]
  const vertical = VERTICAL[(xf.alignment >>> 4) & 7] ?? 'bottom'
  const wrap = (xf.alignment >>> 3) & 1
  const justify = xf.alignment >>> 7
  const shrink = (xf.indent >>> 4) & 1
  const order = xf.indent >>> 6
  return `<alignment horizontal="${horizontal}" vertical="${vertical}" wrapText="${wrap}" justifyLastLine="${justify}"`
    + ` shrinkToFit="${shrink}" readingOrder="${order}" textRotation="${xf.rotation}" indent="${xf.ext?.get(15) ?? (xf.indent & 15)}"/>`
}

function border(xf: Xf): string {
  const down = (xf.borders >>> 30) & 1
  const up = xf.borders >>> 31
  let xml = `<border diagonalDown="${down}" diagonalUp="${up}">`
  const sides = [
    ['left', 9, xf.borders & 15, (xf.borders >>> 16) & 127],
    ['right', 10, (xf.borders >>> 4) & 15, (xf.borders >>> 23) & 127],
    ['top', 7, (xf.borders >>> 8) & 15, xf.moreBorders & 127],
    ['bottom', 8, (xf.borders >>> 12) & 15, (xf.moreBorders >>> 7) & 127],
    ['diagonal', 11, (xf.moreBorders >>> 21) & 15, (xf.moreBorders >>> 14) & 127],
  ] as const
  for (const [side, property, line, index] of sides) {
    if (line && LINES[line]) xml += `<${side} style="${LINES[line]}"><color ${xf.ext?.get(property) ?? color(index)}/></${side}>`
  }
  return xml + '</border>'
}

function fill(xf: Xf): string {
  const gradient = xf.ext?.get(6)
  if (typeof gradient === 'string') return gradient
  const pattern = PATTERNS[xf.moreBorders >>> 26] ?? 'none'
  const foreground = xf.ext?.get(4) ?? color(xf.fill & 127)
  const background = xf.ext?.get(5) ?? color((xf.fill >>> 7) & 127)
  return `<fill><patternFill patternType="${pattern}"><fgColor ${foreground}/><bgColor ${background}/></patternFill></fill>`
}

export class Styles {
  private readonly fonts: Font[] = []
  private readonly formats = new Map<number, string>()
  private readonly xfs: Xf[] = []
  private readonly mapped = new Map<number, number>()
  private readonly palette: string[] = []
  private readonly dxfs: string[] = []
  private readonly savedDxfs: number[] = []
  private readonly tableStyles = new TableStyles()
  private legacyFormats = 0
  private biff2Xfs = 0
  private readonly inlineXfs = new Map<number, number>()
  shortenedFormat = false

  private store(style: Xf): number {
    const index = this.xfs.length
    if (!(style.protection & 4)) this.mapped.set(index, this.mapped.size)
    return this.xfs.push(style) - 1
  }

  /** BIFF2 can hold cell format bits without a referenced XF record. */
  biff2Index(payload: Payload, extendedIndex?: number): number | undefined {
    const flags = payload.u8()
    const format = payload.u8()
    const style = payload.u8()
    const short = flags & 63
    const index = short === 63 ? extendedIndex : short
    if (index !== undefined && index < this.biff2Xfs) return this.index(index)
    const protection = flags >>> 6
    if (format === 0 && style === 0 && protection < 2) return undefined
    const key = (protection << 16) | (format << 8) | style
    let custom = this.inlineXfs.get(key)
    if (custom === undefined) {
      custom = this.store(xf2(format >>> 6, format & 63, protection, style))
      this.inlineXfs.set(key, custom)
    }
    return this.index(custom)
  }

  biff2ColumnStyles(payload: Payload): { first: number; styles: (number | undefined)[] } {
    const first = payload.u16()
    // Excel stops before the last column in COLUMNDEFAULT, like the sheet's dimensions.
    const end = payload.u16()
    if (first > end || end > 256) payload.fail('invalid default column range')
    const styles: (number | undefined)[] = []
    for (let column = first; column < end; column++) styles.push(this.biff2Index(payload))
    return { first, styles }
  }

  take(type: number, payload: Payload, encoding?: LegacyEncoding): boolean {
    switch (type) {
      case 0x088d:
        payload.skip(12)
        this.savedDxfs.push(this.dxf(differential(payload)))
        return true
      case 0x0031: case 0x0231: this.fonts.push(font(payload, encoding)); return true
      case 0x0045: {
        if (encoding?.version !== 2) return false
        const font = this.fonts.at(-1)
        const index = payload.u16()
        if (font && index !== 0x7fff) font.props += `<color ${color(index)}/>`
        return true
      }
      case 0x00e0: this.store(xf(payload, encoding !== undefined)); return true
      case 0x0043: case 0x0243: case 0x0443: {
        if (!encoding || encoding.version >= 5) return false
        this.store(legacyXf(payload, encoding.version))
        if (encoding.version === 2) this.biff2Xfs++
        return true
      }
      case 0x087d: {
        payload.skip(14)
        const index = payload.u16()
        payload.skip(2)
        const count = payload.u16()
        const style = this.xfs[index]
        if (style) style.ext = extensions(payload, count)
        return true
      }
      case 0x001e: case 0x041e: {
        if (encoding) {
          const declared = type === 0x041e ? payload.u16() : undefined
          const id = encoding.version < 5 ? this.legacyFormats++ : declared
          if (id === undefined) return false
          this.formats.set(id, encoding.text(payload, payload.u8()))
        } else {
          if (type !== 0x041e) return false
          const id = payload.u16()
          const count = payload.u16()
          const wide = Boolean(payload.u8() & 1)
          const available = Math.floor(payload.remaining / (wide ? 2 : 1))
          const length = Math.min(count, available)
          this.shortenedFormat ||= length !== count
          this.formats.set(id, wide ? payload.wide(length) : payload.narrow(length))
        }
        return true
      }
      case 0x0092: {
        const count = payload.u16()
        for (let i = 0; i < count; i++) {
          // LongRGB stores R, G, B, then a reserved byte.
          const value = [payload.u8(), payload.u8(), payload.u8()]
          payload.skip(1)
          this.palette.push('FF' + value.map((n) => n.toString(16).padStart(2, '0')).join('').toUpperCase())
        }
        return true
      }
      default: return this.tableStyles.take(type, payload)
    }
  }

  private fontIndex(index: number): number { return index > 4 ? index - 1 : index }

  normalFont(): Font | undefined {
    return this.fonts[0]
  }

  run(index: number, text: string): string {
    const font = this.fonts[this.fontIndex(index)]
    const props = font ? `<rPr><rFont val="${escapeAttr(font.name)}"/>${font.props}</rPr>` : ''
    return '<r>' + props + text + '</r>'
  }

  dxf(value: string): number {
    return this.dxfs.push(`<dxf>${value}</dxf>`) - 1
  }

  savedDxf(index: number): number | undefined {
    return this.savedDxfs[index]
  }

  index(index: number): number | undefined {
    return this.mapped.get(index)
  }

  write(xml: Xml): string {
    xml.add('<styleSheet><numFmts>')
    for (const [id, code] of this.formats) {
      const format = /^general$/i.test(code) ? 'General' : code
      xml.add(`<numFmt numFmtId="${id}" formatCode="${escapeAttr(format)}"/>`)
    }
    xml.add('</numFmts><fonts>')
    const cells = [...this.mapped.keys()].map((index) => this.xfs[index])
    const putFont = (font: Font, color?: string | number) => {
      const props = color ? font.props.replace(/<color [^>]*\/>/, '') + `<color ${color}/>` : font.props
      xml.add(`<font><name val="${escapeAttr(font.name)}"/>${props}</font>`)
    }
    const baseColor = cells[0]?.font === 0 ? cells[0].ext?.get(13) : undefined
    const normal = this.fonts[0]
    // The model's workbook font carries only its name and size. Match cell color against
    // ordinary black so a colored Normal font remains explicit in the cell styles.
    const explicitNormal = normal && (baseColor !== undefined || normal.props.includes('<color '))
    const normalId = explicitNormal ? this.fonts.length : 0
    for (const [i, font] of this.fonts.entries()) putFont(font, i === 0 && explicitNormal ? 'rgb="FF000000"' : undefined)
    if (explicitNormal) putFont(normal, baseColor)
    const firstColorFont = this.fonts.length + (explicitNormal ? 1 : 0)
    const colorFonts = new Map<string, number>()
    const fontIds = cells.map((style) => {
      const index = this.fontIndex(style.font)
      const font = this.fonts[index]
      const color = style.ext?.get(13)
      if (index === 0 && (!color || color === baseColor)) return normalId
      if (!font || !color) return index
      const key = `${index}:${color}`
      const found = colorFonts.get(key)
      if (found !== undefined) return found
      const id = firstColorFont + colorFonts.size
      colorFonts.set(key, id)
      putFont(font, color)
      return id
    })
    xml.add('</fonts><fills><fill/><fill/>')
    for (const style of cells) xml.add(fill(style))
    xml.add('</fills><borders>')
    for (const style of cells) xml.add(border(style))
    xml.add('</borders><cellXfs>')
    for (let i = 0; i < cells.length; i++) {
      const style = cells[i]
      const locked = style.protection & 1
      const hidden = (style.protection >>> 1) & 1
      xml.add(`<xf numFmtId="${style.format}" fontId="${fontIds[i]}" fillId="${i + 2}" borderId="${i}">`
        + alignment(style) + `<protection locked="${locked}" hidden="${hidden}"/></xf>`)
    }
    xml.add('</cellXfs><dxfs>')
    for (const dxf of this.dxfs) xml.add(dxf)
    xml.add('</dxfs>')
    this.tableStyles.write(xml, (id) => this.savedDxf(id))
    if (this.palette.length) {
      // BIFF's palette starts at index 8. The preceding entries are the fixed system colors.
      xml.add('<colors><indexedColors>')
      for (const rgb of ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF']) xml.add(`<rgbColor rgb="FF${rgb}"/>`)
      for (const rgb of this.palette) xml.add(`<rgbColor rgb="${rgb}"/>`)
      xml.add('</indexedColors></colors>')
    }
    xml.add('</styleSheet>')
    return xml.text()
  }
}
