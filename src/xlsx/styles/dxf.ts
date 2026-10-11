/**
 * The differential formats of styles.xml (`<dxfs>`), which conditional formats point at, as the model's
 * `ConditionalStyle`: a font's bold, italic, underline, strike and colour (`<b val="0"/>` takes bold away), a solid
 * fill (a dxf keeps its colour in `bgColor`), borders and a number format. What it cannot hold is named for the reader
 * to report on the rules that use it.
 */
import type { ConditionalStyle } from '../../contract.js'
import { isBuiltin } from '../../numfmt/builtin.js'
import { attributes, element, elements, flag, type Element } from '../xml/scan.js'
import { decodeXml } from '../xml/text.js'
import { colorOf, type Palette } from './color.js'
import { BUILTIN } from './numfmt.js'
import { border } from './read.js'

export interface Dxf {
  style: ConditionalStyle
  lost: string[]
}

/** A flag element of a font: true, false (`val="0"`) or absent. */
function on(body: string, local: string): boolean | undefined {
  const e = element(body, local)
  return e === undefined ? undefined : flag(attributes(e.attrs).val, true)
}

/** A dxf as the model's style. */
export function readDxf(e: Element, p: Palette): Dxf {
  const b = e.body ?? ''
  const style: ConditionalStyle = {}
  const lost: string[] = []
  const f = element(b, 'font')
  if (f?.body) {
    const fb = f.body
    const font: NonNullable<ConditionalStyle['font']> = {}
    for (const [local, key] of [['b', 'bold'], ['i', 'italic'], ['strike', 'strike']] as const) {
      const v = on(fb, local)
      if (v !== undefined) font[key] = v
    }
    const u = element(fb, 'u')
    const under = u && (attributes(u.attrs).val ?? 'single')
    if (under === 'single' || under === 'double') font.underline = under
    else if (under === 'singleAccounting' || under === 'doubleAccounting') {
      font.underline = under === 'singleAccounting' ? 'single' : 'double'
      lost.push('accounting underline')
    } else if (under === 'none') lost.push('an underline taken away')
    const color = colorOf(element(fb, 'color'), p)
    if (color) font.color = color
    for (const x of ['name', 'sz', 'vertAlign', 'outline', 'shadow', 'condense', 'extend']) if (element(fb, x)) lost.push(`font ${x}`)
    if (Object.keys(font).length) style.font = font
  }
  const fill = element(b, 'fill')
  if (fill?.body) {
    if (element(fill.body, 'gradientFill')) lost.push('gradient fill')
    const pf = element(fill.body, 'patternFill')
    const type = pf && attributes(pf.attrs).patternType
    if (pf && type !== undefined && type !== 'solid' && type !== 'none') lost.push('pattern fill')
    else if (pf?.body) {
      // a dxf's solid fill is its background colour
      const color = colorOf(element(pf.body, 'bgColor'), p) ?? colorOf(element(pf.body, 'fgColor'), p)
      if (color) style.fill = { color }
    }
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
      if (isBuiltin(id)) style.format = id
    } else style.format = isBuiltin(id) && id !== 0 && BUILTIN[id] === code ? id : code
  }
  if (element(b, 'alignment')) lost.push('alignment')
  if (element(b, 'protection')) lost.push('protection')
  return { style, lost }
}

/** The dxfs of styles.xml in order. */
export const readDxfs = (xml: string | undefined, p: Palette): Dxf[] => (xml ? elements(element(xml, 'dxfs')?.body ?? '', 'dxf').map((e) => readDxf(e, p)) : [])
