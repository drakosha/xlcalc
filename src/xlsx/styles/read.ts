/**
 * styles.xml as the model's style table: each cell format (`<xf>` of `<cellXfs>`) becomes a style of the book
 * (font, solid fill, borders, alignment; equal ones shared) and a number format, a code or a built-in format's number.
 * What a style cannot hold (pattern and gradient fills, diagonal borders, rotation, shrinking, sub- and superscript,
 * protection, a font's outline or shadow) is named per cell format, for the reader to report on the cells that use it.
 */
import type { AlignmentStyle, BorderLine, CellStyle, FontStyle, NumberFormat } from '../../contract.js'
import { HORIZONTAL, LINES, VERTICAL } from '../../input.js'
import { isBuiltin } from '../../numfmt/builtin.js'
import { attributes, element, elements, flag, type Element } from '../xml/scan.js'
import { decodeXml } from '../xml/text.js'
import { colorOf, type Palette } from './color.js'
import { BUILTIN } from './numfmt.js'

/** A cell format as the model sees it. */
export interface Xf {
  /** The book style, by index; none for a format that looks as the default. */
  style?: number
  /** The number format, a code or a built-in format's number; none for General. */
  format?: NumberFormat
  /** A built-in format the model does not have, read as the code Excel for Mac en-US shows (50 to 58). */
  eastAsian?: number
  /** A built-in number format this reader does not know. */
  unknownFormat?: number
  /** What the style table cannot hold. */
  lost: string[]
}

export interface StyleTable {
  styles: CellStyle[]
  xfs: Xf[]
  /** The default font (the first one), when there are fonts. */
  font?: { name?: string; size?: number }
}

const on = (body: string, local: string): boolean => {
  const e = element(body, local)
  return e !== undefined && flag(attributes(e.attrs).val, true)
}

const val = (body: string, local: string): string | undefined => {
  const e = element(body, local)
  return e && attributes(e.attrs).val
}

interface Font {
  style: FontStyle
  lost: string[]
}

/**
 * A font's underline (`<u>`, single without a `val`); an accounting one is read as its plain kind and named in `lost`,
 * and so is a dxf's `none`, which takes an underline away (a style cannot say so; a cell's font has none then).
 */
export function underlineOf(body: string, lost: string[], dxf = false): FontStyle['underline'] {
  const u = element(body, 'u')
  const under = u && (attributes(u.attrs).val ?? 'single')
  if (under === 'single' || under === 'double') return under
  if (under === 'singleAccounting' || under === 'doubleAccounting') {
    lost.push('accounting underline')
    return under === 'singleAccounting' ? 'single' : 'double'
  }
  if (dxf && under === 'none') lost.push('an underline taken away')
  return undefined
}

export function font(e: Element, p: Palette): Font {
  const b = e.body ?? ''
  const style: FontStyle = {}
  const lost: string[] = []
  const name = val(b, 'name')
  if (name !== undefined) style.name = decodeXml(name)
  const size = Number(val(b, 'sz'))
  if (size > 0) style.size = size
  if (on(b, 'b')) style.bold = true
  if (on(b, 'i')) style.italic = true
  const underline = underlineOf(b, lost)
  if (underline) style.underline = underline
  if (on(b, 'strike')) style.strike = true
  const color = colorOf(element(b, 'color'), p)
  if (color) style.color = color
  const va = val(b, 'vertAlign')
  if (va === 'superscript' || va === 'subscript') lost.push(va)
  for (const x of ['outline', 'shadow']) if (on(b, x)) lost.push(`font ${x}`)
  return { style, lost }
}

export function border(e: Element, p: Palette): { border?: CellStyle['border']; lost: string[] } {
  const b = e.body ?? ''
  const out: NonNullable<CellStyle['border']> = {}
  const lost: string[] = []
  const side = (locals: string[]): BorderLine | undefined => {
    for (const local of locals) {
      const s = element(b, local)
      const kind = s && attributes(s.attrs).style
      if (!s || !kind || kind === 'none') continue
      if (!LINES.includes(kind as BorderLine['style'])) return void lost.push(`border ${kind}`)
      const color = s.body ? colorOf(element(s.body, 'color'), p) : undefined
      return color ? { style: kind as BorderLine['style'], color } : { style: kind as BorderLine['style'] }
    }
    return undefined
  }
  const sides = { top: side(['top']), right: side(['right', 'end']), bottom: side(['bottom']), left: side(['left', 'start']) }
  for (const [k, v] of Object.entries(sides)) if (v) out[k as keyof typeof sides] = v
  if (side(['diagonal'])) lost.push('diagonal border')
  return Object.keys(out).length ? { border: out, lost } : { lost }
}

/**
 * A fill (`<fill>`) of a cell format or, with `dxf`, of a conditional format, as Excel for Mac and for Windows show them
 * (the set `cond/fills-held`). Both read a pattern beside a gradient alone (Excel shows a cell so, and opens no file with
 * a dxf so), and name a gradient or a pattern other than solid lost. They read a pattern fill otherwise, as Excel does:
 * without a pattern type a cell's is none (Excel saves it so) and a dxf's solid (Excel writes its dxfs so); a cell's solid
 * fill shows its `fgColor`, white without one (Excel's automatic colour), and a dxf's its `bgColor` alone, keeping the
 * cell's own colour without one (white on a cell without a fill); a dxf's none takes the cell's fill away. The last two
 * the model cannot say: they are named lost.
 */
export function readFill(e: Element, p: Palette, dxf = false): { fill?: CellStyle['fill']; lost: string[] } {
  const b = e.body ?? ''
  const pf = element(b, 'patternFill')
  const lost = element(b, 'gradientFill') ? [pf ? 'gradient fill beside its pattern' : 'gradient fill'] : []
  if (!pf) return { lost }
  const type = attributes(pf.attrs).patternType ?? (dxf ? 'solid' : 'none')
  if (type === 'none') return { lost: dxf ? [...lost, 'fill taken away'] : lost }
  if (type !== 'solid') return { lost: [...lost, 'pattern fill'] }
  const color = colorOf(element(pf.body ?? '', dxf ? 'bgColor' : 'fgColor'), p)
  if (color) return { fill: { color }, lost }
  return dxf ? { lost: [...lost, 'fill without its colour'] } : { fill: { color: '#FFFFFF' }, lost }
}

function alignment(e: Element | undefined): { alignment?: AlignmentStyle; lost: string[] } {
  if (!e) return { lost: [] }
  const a = attributes(e.attrs)
  const out: AlignmentStyle = {}
  const lost: string[] = []
  const horizontal = a.horizontal as AlignmentStyle['horizontal']
  const vertical = a.vertical as AlignmentStyle['vertical']
  // the defaults, general and bottom, are left out
  if (horizontal && HORIZONTAL.includes(horizontal) && horizontal !== 'general') out.horizontal = horizontal
  if (vertical && VERTICAL.includes(vertical) && vertical !== 'bottom') out.vertical = vertical
  if (flag(a.wrapText)) out.wrap = true
  const indent = Number(a.indent ?? 0)
  if (indent > 0) out.indent = Math.min(250, Math.round(indent))
  if (Number(a.textRotation ?? 0)) lost.push('text rotation')
  if (flag(a.shrinkToFit)) lost.push('shrink to fit')
  if (Number(a.readingOrder ?? 0)) lost.push('reading order')
  if (flag(a.justifyLastLine)) lost.push('justify last line')
  return Object.keys(out).length ? { alignment: out, lost } : { lost }
}

/** A font's fields that differ from the default font: what a style leaves out is the default. */
function relative(f: FontStyle, base: FontStyle): FontStyle | undefined {
  const out: FontStyle = { ...f }
  if (out.name === base.name) delete out.name
  if (out.size === base.size) delete out.size
  if (out.color === base.color) delete out.color
  return Object.keys(out).length ? out : undefined
}

export function readStyles(xml: string | undefined, p: Palette): StyleTable {
  if (!xml) return { styles: [], xfs: [] }
  const formats = new Map<number, string>()
  for (const e of elements(element(xml, 'numFmts')?.body ?? '', 'numFmt')) {
    const a = attributes(e.attrs)
    if (a.numFmtId !== undefined && a.formatCode !== undefined) formats.set(Number(a.numFmtId), decodeXml(a.formatCode))
  }
  const fonts = elements(element(xml, 'fonts')?.body ?? '', 'font').map((e) => font(e, p))
  const base = fonts[0]?.style ?? {}
  const fills = elements(element(xml, 'fills')?.body ?? '', 'fill').map((e) => readFill(e, p))
  const borders = elements(element(xml, 'borders')?.body ?? '', 'border').map((e) => border(e, p))
  const styles: CellStyle[] = []
  const byKey = new Map<string, number>()
  const xfs = elements(element(xml, 'cellXfs')?.body ?? '', 'xf').map((e): Xf => {
    const a = attributes(e.attrs)
    const f = fonts[Number(a.fontId ?? 0)]
    const fi = fills[Number(a.fillId ?? 0)]
    const bo = borders[Number(a.borderId ?? 0)]
    const al = alignment(e.body ? element(e.body, 'alignment') : undefined)
    const pr = e.body ? element(e.body, 'protection') : undefined
    const pa = pr ? attributes(pr.attrs) : {}
    const lost = [...(f?.lost ?? []), ...(fi?.lost ?? []), ...(bo?.lost ?? []), ...al.lost, ...(!flag(pa.locked, true) || flag(pa.hidden) ? ['protection'] : [])]
    const style: CellStyle = {}
    const fs = f && relative(f.style, base)
    if (fs) style.font = fs
    if (fi?.fill) style.fill = fi.fill
    if (bo?.border) style.border = bo.border
    if (al.alignment) style.alignment = al.alignment
    const xf: Xf = { lost }
    if (Object.keys(style).length) {
      const key = JSON.stringify(style)
      let i = byKey.get(key)
      if (i === undefined) byKey.set(key, (i = styles.push(style) - 1))
      xf.style = i
    }
    const id = Number(a.numFmtId ?? 0)
    const own = formats.get(id)
    // a format of the file's own by its code, a built-in one by its number, as the model keeps it
    if (own !== undefined) {
      if (own !== 'General') xf.format = own
    } else if (isBuiltin(id)) {
      if (id !== 0) xf.format = id
    } else if (BUILTIN[id] !== undefined) {
      xf.format = BUILTIN[id]
      xf.eastAsian = id
    } else xf.unknownFormat = id
    return xf
  })
  return { styles, xfs, ...(fonts.length ? { font: { name: base.name, size: base.size } } : {}) }
}
