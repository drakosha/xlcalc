/**
 * styles.xml from the book's styles: a cell format (`<xf>`) for each pair of a style and a number format that
 * cells, rows and columns use, fonts, fills and borders shared between them. Format 0 is the default (Calibri 11,
 * General). Number formats Excel shows by the regional settings are written as codes of their own, so the file
 * shows what `numberFormat` says in every locale.
 */
import type { BorderLine, CellStyle, ConditionalStyle, FontStyle, NumberFormat } from '../../contract.js'
import { BUILTIN, BUILTIN_ID } from '../styles/numfmt.js'
import { escapeAttr } from '../xml/text.js'

/** The book's default font, as the model assumes it. */
export const DEFAULT_FONT = { name: 'Calibri', size: 11 }

const rgb = (c: string) => `FF${c.slice(1).toUpperCase()}`

/** A font's underline, as a cell's font and a dxf's write it. */
const underlineXml = (u: FontStyle['underline']) => (u ? (u === 'double' ? '<u val="double"/>' : '<u/>') : '')

function fontXml(f: FontStyle): string {
  const out = [
    f.bold ? '<b/>' : '',
    f.italic ? '<i/>' : '',
    f.strike ? '<strike/>' : '',
    underlineXml(f.underline),
    `<sz val="${f.size ?? DEFAULT_FONT.size}"/>`,
    f.color ? `<color rgb="${rgb(f.color)}"/>` : '',
    `<name val="${escapeAttr(f.name ?? DEFAULT_FONT.name)}"/>`,
  ]
  return `<font>${out.join('')}</font>`
}

const line = (side: string, l: BorderLine | undefined) => (l ? `<${side} style="${l.style}">${l.color ? `<color rgb="${rgb(l.color)}"/>` : ''}</${side}>` : `<${side}/>`)

function borderXml(b: NonNullable<CellStyle['border']>): string {
  return `<border>${line('left', b.left)}${line('right', b.right)}${line('top', b.top)}${line('bottom', b.bottom)}<diagonal/></border>`
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

export class StyleWriter {
  private readonly fonts = new Table(fontXml({}))
  private readonly fills = new Table('<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>')
  private readonly borders = new Table('<border><left/><right/><top/><bottom/><diagonal/></border>')
  private readonly formats = new Map<string, number>()
  private readonly xfs = new Table('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')
  private readonly byKey = new Map<string, number>([['\n', 0]])
  private readonly dxfs = new Table()

  constructor(private readonly styles: CellStyle[] = []) {}

  private numFmt(code: NumberFormat | undefined): number {
    // a built-in format given by number is written as that number, as Excel stores it
    if (typeof code === 'number') return code
    if (code === undefined || code === 'General') return 0
    const builtin = BUILTIN_ID.get(code)
    if (builtin !== undefined) return builtin
    let id = this.formats.get(code)
    if (id === undefined) this.formats.set(code, (id = 164 + this.formats.size))
    return id
  }

  /** The cell format of a style (by index) and a number format; 0 is the default. */
  xf(style: number | undefined, format: NumberFormat | undefined): number {
    const key = `${style ?? ''}\n${format ?? ''}`
    let i = this.byKey.get(key)
    if (i !== undefined) return i
    const s = style === undefined ? {} : this.styles[style]
    const numFmt = this.numFmt(format)
    const font = s.font ? this.fonts.of(fontXml(s.font)) : 0
    const fill = s.fill ? this.fills.of(`<fill><patternFill patternType="solid"><fgColor rgb="${rgb(s.fill.color)}"/><bgColor indexed="64"/></patternFill></fill>`) : 0
    const border = s.border ? this.borders.of(borderXml(s.border)) : 0
    const a = s.alignment
    const align = a
      ? `<alignment${a.horizontal ? ` horizontal="${a.horizontal}"` : ''}${a.vertical ? ` vertical="${a.vertical}"` : ''}${a.wrap ? ' wrapText="1"' : ''}${a.indent ? ` indent="${a.indent}"` : ''}/>`
      : ''
    const applied = [numFmt ? ' applyNumberFormat="1"' : '', font ? ' applyFont="1"' : '', fill ? ' applyFill="1"' : '', border ? ' applyBorder="1"' : '', align ? ' applyAlignment="1"' : ''].join('')
    const open = `<xf numFmtId="${numFmt}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"${applied}`
    i = this.xfs.of(align ? `${open}>${align}</xf>` : `${open}/>`)
    this.byKey.set(key, i)
    return i
  }

  /** The differential format of a conditional format's style: its index in `<dxfs>`. */
  dxf(s: ConditionalStyle): number {
    const f = s.font
    const flag = (tag: string, v: boolean | undefined) => (v === undefined ? '' : v ? `<${tag}/>` : `<${tag} val="0"/>`)
    const font = f
      ? `<font>${flag('b', f.bold)}${flag('i', f.italic)}${flag('strike', f.strike)}${underlineXml(f.underline)}${f.color ? `<color rgb="${rgb(f.color)}"/>` : ''}</font>`
      : ''
    const code = typeof s.numberFormat === 'number' ? (s.numberFormat === 0 ? 'General' : BUILTIN[s.numberFormat]) : s.numberFormat
    const numFmt = s.numberFormat === undefined ? '' : `<numFmt numFmtId="${this.numFmt(s.numberFormat)}" formatCode="${escapeAttr(code!)}"/>`
    // a dxf's solid fill is its background colour
    const fill = s.fill ? `<fill><patternFill><bgColor rgb="${rgb(s.fill.color)}"/></patternFill></fill>` : ''
    const border = s.border ? borderXml(s.border) : ''
    return this.dxfs.of(`<dxf>${font}${numFmt}${fill}${border}</dxf>`)
  }

  /** The body of a dxf, as an Excel 2010 rule holds it (`x14:dxf`). */
  dxfBody(i: number): string {
    return this.dxfs.items[i].slice(5, -6)
  }

  xml(): string {
    const list = (tag: string, items: string[]) => `<${tag} count="${items.length}">${items.join('')}</${tag}>`
    const numFmts = [...this.formats].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${escapeAttr(code)}"/>`)
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
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
