/**
 * styles.xml from the book's styles: a cell format (`<xf>`) for each pair of a style and a number format that
 * cells, rows and columns use, fonts, fills and borders shared between them. Format 0 is the default (Calibri 11,
 * General). Number formats Excel shows by the regional settings are written as codes of their own, so the file
 * shows what `numberFormat` says in every locale.
 */
import { InputError } from '../../input-error.js'
import { formatRefused } from '../../numfmt/refused.js'
import type { BorderLine, CellStyle, ConditionalStyle, FontStyle, NumberFormat } from '../../contract.js'
import { BUILTIN, BUILTIN_ID } from '../styles/numfmt.js'
import type { Issues } from '../issues.js'
import { escapeAttr } from '../xml/text.js'
import { clip, counted, HEAD, MAIN, rgb } from './xml.js'

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

/** A written font's name: an empty one falls back to the workbook's, a long one is clipped to Excel's limit. */
export function writtenFont(f: FontStyle, issues?: Issues): FontStyle {
  if (f.name === undefined) return f
  if (f.name !== '') return { ...f, name: issues ? clip(f.name, MAX_FONT_NAME, issues, { key: 'fontNameCut' }, {}) : f.name }
  issues?.add('style', { key: 'fontNameEmpty' })
  const { name: _, ...rest } = f
  return rest
}

/** A cell's solid fill. */
export const fillXml = (f: NonNullable<CellStyle['fill']>) => `<fill><patternFill patternType="solid"><fgColor rgb="${rgb(f.color)}"/><bgColor indexed="64"/></patternFill></fill>`

/** A cell format's alignment, empty without one. */
function alignmentXml(a: CellStyle['alignment']): string {
  return a ? `<alignment${a.horizontal ? ` horizontal="${a.horizontal}"` : ''}${a.vertical ? ` vertical="${a.vertical}"` : ''}${a.wrap ? ' wrapText="1"' : ''}${a.indent ? ` indent="${a.indent}"` : ''}${a.rotation ? ` textRotation="${a.rotation === 'vertical' ? 255 : a.rotation > 0 ? a.rotation : 90 - a.rotation}"` : ''}/>` : ''
}

/** A cell format's protection, empty without one. */
function cellProtectionXml(p: CellStyle['protection']): string {
  return p && (p.locked === false || p.hidden) ? `<protection${p.locked === false ? ' locked="0"' : ''}${p.hidden ? ' hidden="1"' : ''}/>` : ''
}

/**
 * A cell format (`<xf>`) of a number format and a font, fill and border by index, with the style's alignment and protection;
 * `button`, Excel's filter button of a PivotTable's header (`pivotButton`), which Excel draws only from the file until a refresh.
 */
export function xfXml(numFmt: number, font: number, fill: number, border: number, s: CellStyle | undefined, button = false): string {
  const align = alignmentXml(s?.alignment)
  const protect = cellProtectionXml(s?.protection)
  const applied = [numFmt ? ' applyNumberFormat="1"' : '', font ? ' applyFont="1"' : '', fill ? ' applyFill="1"' : '', border ? ' applyBorder="1"' : '', align ? ' applyAlignment="1"' : '', protect ? ' applyProtection="1"' : ''].join('')
  const open = `<xf numFmtId="${numFmt}" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"${button ? ' pivotButton="1"' : ''}${applied}`
  return align || protect ? `${open}>${align}${protect}</xf>` : `${open}/>`
}

/** A conditional format's differential format, its number format by `numFmt`'s id. */
export function dxfXml(s: ConditionalStyle, numFmt: (code: NumberFormat) => number, filter = false): string {
  if (filter) return `<dxf>${s.fill ? `<fill><patternFill patternType="solid"><fgColor rgb="${rgb(s.fill.color)}"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>` : ''}</dxf>`
  const f = s.font
  const flag = (tag: string, v: boolean | undefined) => (v === undefined ? '' : v ? `<${tag}/>` : `<${tag} val="0"/>`)
  const font = f
    ? `<font>${flag('b', f.bold)}${flag('i', f.italic)}${flag('strike', f.strike)}${underlineXml(f.underline)}${f.color ? `<color rgb="${rgb(f.color)}"/>` : ''}</font>`
    : ''
  const code = typeof s.numberFormat === 'number' ? (s.numberFormat === 0 ? 'General' : BUILTIN[s.numberFormat]) : s.numberFormat
  const id = s.numberFormat === undefined ? undefined : numFmt(s.numberFormat)
  // no id for a custom format over the limit (`StyleWriter.numFmt`): its code in the file would count too
  const numFmtXml = id === undefined || (id === 0 && code !== 'General') ? '' : `<numFmt numFmtId="${id}" formatCode="${escapeAttr(code!)}"/>`
  // a dxf's solid fill is its background colour
  const fill = s.fill ? `<fill><patternFill><bgColor rgb="${rgb(s.fill.color)}"/></patternFill></fill>` : ''
  const border = s.border ? borderXml(s.border) : ''
  return `<dxf>${font}${numFmtXml}${fill}${border}</dxf>`
}

/** A code a file is to take: refused where Excel does not open a file with it, a code no cell uses too. */
export function checkWritten(code: string): void {
  if (formatRefused(code)) throw new InputError('numberFormat', `Excel does not open a file with the number format ${JSON.stringify(code)}: text in a format goes in quotes (0 "kg", 0 "шт/ч")`)
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
  /** The cell format of a style (by index) and a number format, with a PivotTable's filter button; 0 is the default. */
  xf(style: number | undefined, format: NumberFormat | undefined, button?: boolean): number
  /** A whole run font, or its cell's font, with a written name. */
  fontOf(style?: number, override?: FontStyle, issues?: Issues): FontStyle
  /** The differential format of a conditional format's style: its index in `<dxfs>`. */
  dxf(s: ConditionalStyle, filter?: boolean): number
  /** The body of a dxf, as an Excel 2010 rule holds it (`x14:dxf`). */
  dxfBody(i: number): string
  /** The id of a number format, a custom one added (a pivot cache's field, a value field). */
  numFmt(code: NumberFormat | undefined): number
}

/**
 * What Excel for Windows and for Mac load (`bench/coverage/links/file-refusals*.ts`): different codes of custom number formats,
 * those of `numFmts` and of conditional styles counted together, cell formats, a font's name; over any, it refuses the file.
 */
const MAX_FORMAT_CODES = 206
const MAX_CELL_FORMATS = 65_430
const MAX_FONT_NAME = 31

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
    base: BaseFont = DEFAULT_FONT,
    private readonly issues?: Issues,
  ) {
    this.base = { ...base, name: this.fontName(base.name) }
    this.fonts = new Table(fontXml({}, this.base))
  }

  private readonly base: BaseFont

  private fontName(name: string): string {
    return this.issues ? clip(name, MAX_FONT_NAME, this.issues, { key: 'fontNameCut' }, {}) : name
  }


  numFmt(code: NumberFormat | undefined): number {
    return numFmtId(code, (c) => {
      let id = this.formats.get(c)
      if (id === undefined) {
        checkWritten(c)
        if (this.formats.size >= MAX_FORMAT_CODES) {
          this.issues?.add('numberFormat', { key: 'formatCodesOver' })
          return 0
        }
        this.formats.set(c, (id = 164 + this.formats.size))
      }
      return id
    })
  }

  /** The cell format of a style (by index) and a number format, with a PivotTable's filter button; 0 is the default. */
  xf(style: number | undefined, format: NumberFormat | undefined, button = false): number {
    // by the format's id: 0 (General) and '0' (whole numbers) are two formats, 14 and '14' too; an empty code stays
    // General, as it was written before (whether Excel opens an empty formatCode is not measured)
    const numFmt = format === '' ? 0 : this.numFmt(format)
    const key = `${style ?? ''}\n${numFmt}${button ? '\nb' : ''}`
    let i = this.byKey.get(key)
    if (i !== undefined) return i
    if (this.xfs.items.length >= MAX_CELL_FORMATS) {
      this.issues?.add('style', { key: 'cellFormatsOver' })
      this.byKey.set(key, 0)
      return 0
    }
    const s = style === undefined ? {} : this.styles[style]
    const font = s.font ? this.fonts.of(fontXml(writtenFont(s.font, this.issues), this.base)) : 0
    const fill = s.fill ? this.fills.of(fillXml(s.fill)) : 0
    const border = s.border ? this.borders.of(borderXml(s.border)) : 0
    i = this.xfs.of(xfXml(numFmt, font, fill, border, s, button))
    this.byKey.set(key, i)
    return i
  }

  /** A run's whole font, or its cell's font when no override was given. */
  fontOf(style?: number, override?: FontStyle, issues: Issues | undefined = this.issues): FontStyle {
    return { name: this.base.name, size: this.base.size, ...writtenFont(override ?? this.styles[style ?? -1]?.font ?? {}, issues) }
  }

  /** The differential format of a conditional format's style: its index in `<dxfs>`. */
  dxf(s: ConditionalStyle, filter = false): number {
    return this.dxfs.of(dxfXml(s, (c) => this.numFmt(c), filter))
  }

  /** The body of a dxf, as an Excel 2010 rule holds it (`x14:dxf`). */
  dxfBody(i: number): string {
    return this.dxfs.items[i].slice(5, -6)
  }

  xml(): string {
    const numFmts = [...this.formats].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${escapeAttr(code)}"/>`)
    return (
      `${HEAD}<styleSheet xmlns="${MAIN}">` +
      (numFmts.length ? counted('numFmts', numFmts) : '') +
      counted('fonts', this.fonts.items) +
      counted('fills', this.fills.items) +
      counted('borders', this.borders.items) +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      counted('cellXfs', this.xfs.items) +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      (this.dxfs.items.length ? counted('dxfs', this.dxfs.items) : '') +
      '</styleSheet>'
    )
  }
}
