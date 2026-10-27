/**
 * styles.xml from the book's styles: a cell format (`<xf>`) for each pair of a style and a number format that
 * cells, rows and columns use, fonts, fills and borders shared between them. Format 0 is the default (Calibri 11,
 * General). Number formats Excel shows by the regional settings are written as codes of their own, so the file
 * shows what `numberFormat` says in every locale.
 */
import type { BorderLine, CellStyle, ConditionalStyle, FontStyle, NumberFormat } from '../../contract.js'
import { BUILTIN, BUILTIN_ID } from '../styles/numfmt.js'
import { escapeAttr } from '../xml/text.js'
import { HEAD, MAIN, rgb } from './xml.js'

/** The book's default font, as the model assumes it. */
export const DEFAULT_FONT = { name: 'Calibri', size: 11 }

/** A font's underline, as a cell's font and a dxf's write it. */
const underlineXml = (u: FontStyle['underline']) => (u ? (u === 'double' ? '<u val="double"/>' : '<u/>') : '')

/** The default font a style's font is relative to: its size and name, and its colour's XML, written where the style has none. */
export interface BaseFont {
  name: string
  size: number
  color?: string
  /** The family and theme scheme of the default font, which a font of its name keeps. */
  rest?: string
}

export function fontXml(f: FontStyle, base: BaseFont = DEFAULT_FONT): string {
  const name = f.name ?? base.name
  const out = [
    f.bold ? '<b/>' : '',
    f.italic ? '<i/>' : '',
    f.strike ? '<strike/>' : '',
    underlineXml(f.underline),
    `<sz val="${f.size ?? base.size}"/>`,
    f.color ? `<color rgb="${rgb(f.color)}"/>` : (base.color ?? ''),
    `<name val="${escapeAttr(name)}"/>`,
    name === base.name ? (base.rest ?? '') : '',
  ]
  return `<font>${out.join('')}</font>`
}

const line = (side: string, l: BorderLine | undefined) => (l ? `<${side} style="${l.style}">${l.color ? `<color rgb="${rgb(l.color)}"/>` : ''}</${side}>` : `<${side}/>`)

export function borderXml(b: NonNullable<CellStyle['border']>): string {
  return `<border>${line('left', b.left)}${line('right', b.right)}${line('top', b.top)}${line('bottom', b.bottom)}<diagonal/></border>`
}

/** A cell's solid fill. */
export const fillXml = (f: NonNullable<CellStyle['fill']>) => `<fill><patternFill patternType="solid"><fgColor rgb="${rgb(f.color)}"/><bgColor indexed="64"/></patternFill></fill>`

/** A cell format's alignment, empty without one. */
function alignmentXml(a: CellStyle['alignment']): string {
  return a ? `<alignment${a.horizontal ? ` horizontal="${a.horizontal}"` : ''}${a.vertical ? ` vertical="${a.vertical}"` : ''}${a.wrap ? ' wrapText="1"' : ''}${a.indent ? ` indent="${a.indent}"` : ''}/>` : ''
}

/** A cell format's protection, empty without one. */
function cellProtectionXml(p: CellStyle['protection']): string {
  return p && (p.locked === false || p.hidden) ? `<protection${p.locked === false ? ' locked="0"' : ''}${p.hidden ? ' hidden="1"' : ''}/>` : ''
}

/** A cell format (`<xf>`) of a number format and a font, fill and border by index, with the style's alignment and protection. */
export function xfXml(numFmt: number, font: number, fill: number, border: number, s: CellStyle | undefined): string {
  const align = alignmentXml(s?.alignment)
  const protect = cellProtectionXml(s?.protection)
  const applied = [numFmt ? ' applyNumberFormat="1"' : '', font ? ' applyFont="1"' : '', fill ? ' applyFill="1"' : '', border ? ' applyBorder="1"' : '', align ? ' applyAlignment="1"' : '', protect ? ' applyProtection="1"' : ''].join('')
  const open = `<xf numFmtId="${numFmt}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"${applied}`
  return align || protect ? `${open}>${align}${protect}</xf>` : `${open}/>`
}

/** A conditional format's differential format, its number format by `numFmt`'s id. */
export function dxfXml(s: ConditionalStyle, numFmt: (code: NumberFormat) => number): string {
  const f = s.font
  const flag = (tag: string, v: boolean | undefined) => (v === undefined ? '' : v ? `<${tag}/>` : `<${tag} val="0"/>`)
  const font = f
    ? `<font>${flag('b', f.bold)}${flag('i', f.italic)}${flag('strike', f.strike)}${underlineXml(f.underline)}${f.color ? `<color rgb="${rgb(f.color)}"/>` : ''}</font>`
    : ''
  const code = typeof s.numberFormat === 'number' ? (s.numberFormat === 0 ? 'General' : BUILTIN[s.numberFormat]) : s.numberFormat
  const numFmtXml = s.numberFormat === undefined ? '' : `<numFmt numFmtId="${numFmt(s.numberFormat)}" formatCode="${escapeAttr(code!)}"/>`
  // a dxf's solid fill is its background colour
  const fill = s.fill ? `<fill><patternFill><bgColor rgb="${rgb(s.fill.color)}"/></patternFill></fill>` : ''
  const border = s.border ? borderXml(s.border) : ''
  return `<dxf>${font}${numFmtXml}${fill}${border}</dxf>`
}

/** A number format's id among the built-in ones and those a file defines (`custom`, which a new code joins from 164 on). */
export function numFmtId(code: NumberFormat | undefined, custom: (code: string) => number): number {
  // a built-in format given by number is written as that number, as Excel stores it
  if (typeof code === 'number') return code
  if (code === undefined || code === 'General') return 0
  return BUILTIN_ID.get(code) ?? custom(code)
}

/** An index in a list of distinct XML texts. */
class Table {
  readonly items: string[] = []
  private readonly index = new Map<string, number>()
  constructor(...first: string[]) {
    for (const x of first) this.of(x)
  }
  of(xml: string): number {
    let i = this.index.get(xml)
    if (i === undefined) this.index.set(xml, (i = this.items.push(xml) - 1))
    return i
  }
}

/** What the sheet writers ask of a file's styles: the cell format of a style and a number format, and the differential format of a rule's style. */
export interface CellFormats {
  /** The cell format of a style (by index) and a number format; 0 is the default. */
  xf(style: number | undefined, format: NumberFormat | undefined): number
  /** The differential format of a conditional format's style: its index in `<dxfs>`. */
  dxf(s: ConditionalStyle): number
  /** The body of a dxf, as an Excel 2010 rule holds it (`x14:dxf`). */
  dxfBody(i: number): string
}

export class StyleWriter implements CellFormats {
  private readonly fonts: Table
  private readonly fills = new Table('<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>')
  private readonly borders = new Table('<border><left/><right/><top/><bottom/><diagonal/></border>')
  private readonly formats = new Map<string, number>()
  private readonly xfs = new Table('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')
  private readonly byKey = new Map<string, number>([['\n0', 0]])
  private readonly dxfs = new Table()

  /** `base`: the book's default font, the first of the fonts, which a style's font is relative to. */
  constructor(
    private readonly styles: CellStyle[] = [],
    private readonly base: BaseFont = DEFAULT_FONT,
  ) {
    this.fonts = new Table(fontXml({}, base))
  }

  private numFmt(code: NumberFormat | undefined): number {
    return numFmtId(code, (c) => {
      let id = this.formats.get(c)
      if (id === undefined) this.formats.set(c, (id = 164 + this.formats.size))
      return id
    })
  }

  /** The cell format of a style (by index) and a number format; 0 is the default. */
  xf(style: number | undefined, format: NumberFormat | undefined): number {
    // by the format's id: 0 (General) and '0' (whole numbers) are two formats, 14 and '14' too; an empty code stays
    // General, as it was written before (whether Excel opens an empty formatCode is not measured)
    const numFmt = format === '' ? 0 : this.numFmt(format)
    const key = `${style ?? ''}\n${numFmt}`
    let i = this.byKey.get(key)
    if (i !== undefined) return i
    const s = style === undefined ? {} : this.styles[style]
    const font = s.font ? this.fonts.of(fontXml(s.font, this.base)) : 0
    const fill = s.fill ? this.fills.of(fillXml(s.fill)) : 0
    const border = s.border ? this.borders.of(borderXml(s.border)) : 0
    i = this.xfs.of(xfXml(numFmt, font, fill, border, s))
    this.byKey.set(key, i)
    return i
  }

  /** The differential format of a conditional format's style: its index in `<dxfs>`. */
  dxf(s: ConditionalStyle): number {
    return this.dxfs.of(dxfXml(s, (c) => this.numFmt(c)))
  }

  /** The body of a dxf, as an Excel 2010 rule holds it (`x14:dxf`). */
  dxfBody(i: number): string {
    return this.dxfs.items[i].slice(5, -6)
  }

  xml(): string {
    const list = (tag: string, items: string[]) => `<${tag} count="${items.length}">${items.join('')}</${tag}>`
    const numFmts = [...this.formats].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${escapeAttr(code)}"/>`)
    return (
      `${HEAD}<styleSheet xmlns="${MAIN}">` +
      (numFmts.length ? list('numFmts', numFmts) : '') +
      list('fonts', this.fonts.items) +
      list('fills', this.fills.items) +
      list('borders', this.borders.items) +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      list('cellXfs', this.xfs.items) +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      (this.dxfs.items.length ? list('dxfs', this.dxfs.items) : '') +
      '</styleSheet>'
    )
  }
}
