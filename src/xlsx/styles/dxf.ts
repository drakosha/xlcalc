/**
 * The differential formats of styles.xml (`<dxfs>`), which conditional formats point at, as the model's
 * `ConditionalStyle`: a font's bold, italic, underline, strike and colour (`<b val="0"/>` takes bold away), a solid
 * fill (`readFill`: a dxf keeps its colour in `bgColor`), borders and a number format. What it cannot hold is named for the reader
 * to report on the rules that use it.
 */
import type { CellStyle, Color, ConditionalStyle } from '../../contract.js'
import type { StyleLoss } from '../../locale/messages.js'
import { isBuiltin } from '../../numfmt/builtin.js'
import { attributes, element, flag, within, type Element } from '../xml/scan.js'
import { decodeXml } from '../xml/text.js'
import { colorOf, type Palette } from './color.js'
import { BUILTIN } from './numfmt.js'
import { alignment, border, protection, readFill, underlineOf } from './read.js'

export interface Dxf {
  style: ConditionalStyle
  lost: StyleLoss[]
  /** The color of a solid cell pattern's foreground, which a filter's or sort's format holds where a rule's holds its background. */
  pattern?: Color
  /** What a cell's style holds beyond a conditional one: the font's name and size, the alignment and protection (a PivotTable's formats take them). */
  cell?: Pick<CellStyle, 'alignment' | 'protection'> & { name?: string; size?: number }
}

/** The elements of a dxf's font that a conditional style has no place for. */
const DXF_FONT: [local: string, part: StyleLoss][] = [
  ['name', 'fontName'],
  ['sz', 'fontSize'],
  ['vertAlign', 'fontVertAlign'],
  ['outline', 'fontOutline'],
  ['shadow', 'fontShadow'],
  ['condense', 'fontCondense'],
  ['extend', 'fontExtend'],
]

/** A flag element of a font: true, false (`val="0"`) or absent. */
function on(body: string, local: string): boolean | undefined {
  const e = element(body, local)
  return e === undefined ? undefined : flag(attributes(e.attrs).val, true)
}

/** A dxf as the model's style. */
export function readDxf(e: Element, p: Palette): Dxf {
  const b = e.body ?? ''
  const style: ConditionalStyle = {}
  const lost: StyleLoss[] = []
  let pattern: Color | undefined
  const f = element(b, 'font')
  if (f?.body) {
    const fb = f.body
    const font: NonNullable<ConditionalStyle['font']> = {}
    for (const [local, key] of [['b', 'bold'], ['i', 'italic'], ['strike', 'strike']] as const) {
      const v = on(fb, local)
      if (v !== undefined) font[key] = v
    }
    const underline = underlineOf(fb, lost, true)
    if (underline) font.underline = underline
    const color = colorOf(element(fb, 'color'), p)
    if (color) font.color = color
    for (const [x, part] of DXF_FONT) if (element(fb, x)) lost.push(part)
    if (Object.keys(font).length) style.font = font
  }
  const fill = element(b, 'fill')
  if (fill) {
    const x = readFill(fill, p, true)
    if (x.fill) style.fill = x.fill
    lost.push(...x.lost)
    const pf = element(fill.body ?? '', 'patternFill')
    if (pf && attributes(pf.attrs).patternType === 'solid') pattern = colorOf(element(pf.body ?? '', 'fgColor'), p)
  }
  const bo = element(b, 'border')
  if (bo) {
    const x = border(bo, p)
    if (x.border) style.border = x.border
    lost.push(...x.lost)
  }
  const nf = element(b, 'numFmt')
  if (nf) {
    // a dxf's code is what it shows, in every locale: Excel saves `0.00%` with the id 14 (the set `cond/files`) and
    // shows 14 with `m/d/yyyy` as that code in ru-RU (`cond/formats-ru-held`); the id counts only where the code is that
    // built-in format's own, which a rule's look shows by its en-US code (`lookShown`), so a rule written by number reads back so
    const a = attributes(nf.attrs)
    const id = Number(a.numFmtId)
    const code = a.formatCode === undefined ? undefined : decodeXml(a.formatCode)
    if (code === undefined) {
      if (isBuiltin(id)) style.numberFormat = id
    } else style.numberFormat = isBuiltin(id) && id !== 0 && BUILTIN[id] === code ? id : code
  }
  const cell: NonNullable<Dxf['cell']> = {}
  const name = f?.body && element(f.body, 'name')
  const size = f?.body && element(f.body, 'sz')
  if (name) cell.name = decodeXml(attributes(name.attrs).val ?? '')
  if (size) cell.size = Number(attributes(size.attrs).val)
  const al = alignment(element(b, 'alignment'))
  if (al.alignment) cell.alignment = al.alignment
  const pr = protection(element(b, 'protection'))
  if (pr) cell.protection = pr
  if (element(b, 'alignment')) lost.push('alignment')
  if (element(b, 'protection')) lost.push('protection')
  return { style, lost, ...(pattern ? { pattern } : {}), ...(Object.keys(cell).length ? { cell } : {}) }
}

/** The dxfs of styles.xml in order. */
export const readDxfs = (xml: string | undefined, p: Palette): Dxf[] => (xml ? within(xml, 'dxfs', 'dxf').map((e) => readDxf(e, p)) : [])
