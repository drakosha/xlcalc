/**
 * A sheet's conditional formatting as its `conditionalFormats`: the rules of `<conditionalFormatting>` with their
 * differential formats (`dxfs.ts`), and Excel 2010's rules in the sheet's extension list (`x14:conditionalFormatting`):
 * a data bar's 2010 part joined to its 2007 rule by their id, and the rules only 2010 has (its icon sets, custom icons,
 * formulas over other sheets) with their own priority and format (`x14:dxf`). Rules come in priority order. Formulas
 * come in the engine's spelling; what the model cannot hold is reported.
 */
import type { Color, ConditionalFormat, ConditionalStyle, ConditionalValue, IconSetName } from '../../contract.js'
import { CELL_IS, ICON_SETS, PERIODS } from '../../cond/check.js'
import { defaultFormula, originOf, rangeText } from '../../cond/rules.js'
import { formatAddress } from '../../model.js'
import { parseRange } from '../../structure/area.js'
import type { Issues } from '../issues.js'
import { colorOf, type Palette } from '../styles/color.js'
import { readDxf, type Dxf } from '../styles/dxf.js'
import { attributes, element, elements, flag, type Element } from '../xml/scan.js'
import { readText } from '../xml/text.js'
import { readFormula } from './formulas.js'

/** The extension of a sheet that holds Excel 2010's conditional formatting. */
export const X14_CF = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}'

/** A rule read with its priority and the id that joins a data bar's two parts. */
interface Read {
  rule: ConditionalFormat
  priority: number
  order: number
  id?: string
}

const SIMPLE = new Set(['expression', 'containsText', 'notContainsText', 'beginsWith', 'endsWith', 'containsBlanks', 'notContainsBlanks', 'containsErrors', 'notContainsErrors', 'timePeriod', 'top10', 'aboveAverage', 'duplicateValues', 'uniqueValues', 'cellIs'])

interface Context {
  sheet: string
  dxfs: Dxf[]
  palette: Palette
  issues: Issues
}

/** The worksheet's own `<conditionalFormatting>` blocks (not 2010's), and the rest of the tail. */
function split(tail: string): { blocks: Element[]; rest: string } {
  const blocks: Element[] = []
  const rest = tail.replace(/<conditionalFormatting(\s(?:"[^"]*"|'[^']*'|[^>"'])*?)?(?:\/>|>([\s\S]*?)<\/conditionalFormatting\s*>)/g, (_, attrs: string | undefined, body: string | undefined) => {
    blocks.push({ attrs: attrs ?? '', body: body ?? null })
    return ''
  })
  return { blocks, rest }
}

/** The ranges of a `sqref`, those the model takes, as the model writes them (`A1:A1048576`, Excel's whole column, is `A:A`). */
function ranges(sqref: string, ctx: Context): string[] {
  const out: string[] = []
  for (const r of sqref.replaceAll('$', '').trim().split(/\s+/)) {
    const x = r ? parseRange(r) : null
    if (x) out.push(rangeText(x))
    else if (r) ctx.issues.add('conditionalFormat', 'a range of a conditional format outside the sheet is left out', { sheet: ctx.sheet, cell: r })
  }
  return out
}

/** A formula in the engine's spelling, reported where it does not parse. */
function formula(text: string, where: string, ctx: Context): string {
  return readFormula(readText(text), ctx.issues, { sheet: ctx.sheet, cell: where }).formula
}

/** A threshold: `val` (2007) or `<xm:f>` (2010) as its formula. */
function value(e: Element, where: string, ctx: Context): ConditionalValue {
  const a = attributes(e.attrs)
  const type = (a.type ?? 'min') as ConditionalValue['type']
  const f = e.body ? element(e.body, 'f') : undefined
  const text = f ? (f.body ?? '') : a.val
  return text === undefined || type === 'min' || type === 'max' || type === 'autoMin' || type === 'autoMax' ? { type } : { type, value: formula(text, where, ctx) }
}

/** The colours of a colour scale or data bar, in order. */
const colors = (body: string, local: string, ctx: Context): (Color | undefined)[] => elements(body, local).map((e) => colorOf(e, ctx.palette))

/** A style of a rule: its dxf by id (2007), or its own `x14:dxf` (2010). */
function styleOf(a: Record<string, string>, body: string, where: string, ctx: Context): ConditionalStyle | undefined {
  const own = element(body, 'dxf')
  const d = own ? readDxf(own, ctx.palette) : a.dxfId !== undefined ? ctx.dxfs[Number(a.dxfId)] : undefined
  if (!d) return undefined
  for (const x of d.lost) ctx.issues.add('conditionalFormat', `a conditional format's ${x} is not in the model`, { sheet: ctx.sheet, cell: where })
  return Object.keys(d.style).length ? d.style : undefined
}

/** One `cfRule`, 2007's or 2010's (`x14`: formulas in `<xm:f>`); null for a rule the model does not have. */
function rule(e: Element, sqref: string[], x14: boolean, ctx: Context): Read | null {
  const a = attributes(e.attrs)
  const body = e.body ?? ''
  const where = sqref[0]
  const base = { ranges: sqref, ...(flag(a.stopIfTrue) ? { stopIfTrue: true } : {}) }
  const fs = elements(body, x14 ? 'f' : 'formula').map((f) => formula(f.body ?? '', where, ctx))
  const style = () => styleOf(a, body, where, ctx)
  const withStyle = <T extends object>(r: T): T => {
    const s = style()
    return s ? { ...r, style: s } : r
  }
  const type = a.type ?? ''
  const read = (r: ConditionalFormat): Read => ({ rule: r, priority: Number(a.priority ?? 0), order: 0, ...(a.id ? { id: a.id } : {}) })
  const lost = (why: string) => {
    ctx.issues.add('conditionalFormat', why, { sheet: ctx.sheet, cell: where })
    return null
  }
  if (SIMPLE.has(type)) {
    const own = fs[0] !== undefined ? { formula: fs[0] } : {}
    switch (type) {
      case 'cellIs': {
        const op = a.operator ?? 'equal'
        if (!(CELL_IS as readonly string[]).includes(op)) return lost(`a cell value rule with the operator ${op} is not in the model`)
        const n = op === 'between' || op === 'notBetween' ? 2 : 1
        if (fs.length < n) return lost('a cell value rule without its formulas is left out')
        return read(withStyle({ ...base, type: 'cellIs', operator: op as (typeof CELL_IS)[number], formulas: fs.slice(0, n) }))
      }
      case 'expression':
        return fs.length ? read(withStyle({ ...base, type: 'expression', formula: fs[0] })) : lost('a formula rule without its formula is left out')
      case 'containsText':
      case 'notContainsText':
      case 'beginsWith':
      case 'endsWith':
        return read(withStyle({ ...base, type, text: readText(a.text ?? ''), ...own }))
      case 'timePeriod': {
        const period = a.timePeriod ?? 'today'
        if (!(PERIODS as readonly string[]).includes(period)) return lost(`a time period ${period} is not in the model`)
        return read(withStyle({ ...base, type, period: period as (typeof PERIODS)[number], ...own }))
      }
      case 'top10': {
        const rank = Number(a.rank ?? 10)
        return read(withStyle({ ...base, type, rank: Math.min(Math.max(Math.trunc(rank) || 1, 1), flag(a.percent) ? 100 : 1000), ...(flag(a.percent) ? { percent: true } : {}), ...(flag(a.bottom) ? { bottom: true } : {}) }))
      }
      case 'aboveAverage': {
        const sd = Number(a.stdDev ?? 0)
        const equal = flag(a.equalAverage)
        // Excel keeps equal and drops the deviations of a rule with both (the set `cond/files`)
        if (equal && sd) ctx.issues.add('conditionalFormat', 'an average rule with both equal and standard deviations is read with equal alone, as Excel keeps it', { sheet: ctx.sheet, cell: where })
        return read(withStyle({ ...base, type, ...(!flag(a.aboveAverage, true) ? { below: true } : {}), ...(equal ? { equal: true } : sd >= 1 && sd <= 3 ? { stdDev: Math.trunc(sd) } : {}) }))
      }
      default:
        return read(withStyle({ ...base, type, ...own } as ConditionalFormat))
    }
  }
  if (type === 'colorScale') {
    const cs = element(body, 'colorScale')?.body ?? ''
    const vs = elements(cs, 'cfvo').map((v) => value(v, where, ctx))
    const cl = colors(cs, 'color', ctx)
    if (vs.length < 2 || vs.length > 3 || cl.length !== vs.length || cl.some((c) => !c)) return lost('a colour scale without its points or colours is left out')
    return read({ ...base, type, points: vs.map((v, i) => ({ value: v, color: cl[i]! })) })
  }
  if (type === 'dataBar') {
    const db = element(body, 'dataBar')
    const d = db?.body ?? ''
    const vs = elements(d, 'cfvo').map((v) => value(v, where, ctx))
    const color = x14 ? colorOf(element(d, 'fillColor'), ctx.palette) : colors(d, 'color', ctx)[0]
    if (vs.length !== 2) return lost('a data bar without its thresholds is left out')
    const da = attributes(db?.attrs ?? '')
    const r: Extract<ConditionalFormat, { type: 'dataBar' }> = { ...base, type, min: vs[0], max: vs[1], color: color ?? '#638EC6', ...lengths(da, 10, 90) }
    if (!flag(da.showValue, true)) r.showValue = false
    // the 2010 part's id, in the 2007 rule's extension
    const id = x14 ? a.id : element(body, 'id')?.body?.trim()
    return { rule: x14 ? r : { ...r, excel2007: true }, priority: Number(a.priority ?? 0), order: 0, ...(id ? { id } : {}) }
  }
  if (type === 'iconSet') {
    const is = element(body, 'iconSet')
    const ia = attributes(is?.attrs ?? '')
    const set = (ia.iconSet ?? '3TrafficLights1') as IconSetName
    if (!(set in ICON_SETS) && set !== ('NoIcons' as string)) return lost(`the icon set ${set} is not in the model`)
    const ib = is?.body ?? ''
    const vs = elements(ib, 'cfvo')
    const icons = elements(ib, 'cfIcon').map((c) => {
      const ca = attributes(c.attrs)
      return ca.iconSet === 'NoIcons' ? null : { set: ca.iconSet as IconSetName, index: Number(ca.iconId ?? 0) }
    })
    const n = ICON_SETS[set]
    if (vs.length !== n) return lost('an icon set without a threshold for each icon is left out')
    const reverse = flag(ia.reverse)
    const points = vs.map((v, i) => {
      const gte = flag(attributes(v.attrs).gte, true)
      // a custom icon that is the set's own at its place is no custom icon
      const mine = icons[i]?.set === set && icons[i]?.index === (reverse ? n - 1 - i : i)
      const icon = flag(ia.custom) && !mine ? icons[i] : undefined
      return { value: value(v, where, ctx), ...(gte ? {} : { gte: false }), ...(icon !== undefined ? { icon } : {}) }
    })
    const r: ConditionalFormat = { ...base, type, icons: set, points }
    if (reverse) r.reverse = true
    if (!flag(ia.showValue, true)) r.showValue = false
    return read(r)
  }
  return lost(`a conditional format of type ${type || '(none)'} is not in the model`)
}

/** A bar's shortest and longest lengths (10 and 90 when the file has none, as xlsx's defaults), left out where they are the model's default (`lo`, `hi`). */
function lengths(a: Record<string, string>, lo: number, hi: number): { minLength?: number; maxLength?: number } {
  const [min, max] = [Number(a.minLength ?? 10), Number(a.maxLength ?? 90)]
  const ok = (x: number) => Number.isInteger(x) && x >= 0 && x <= 100
  return { ...(ok(min) && min !== lo ? { minLength: min } : {}), ...(ok(max) && max !== hi ? { maxLength: max } : {}) }
}

/** Excel 2010's part of a data bar laid over its 2007 rule: the 2010 thresholds, fill, border, direction, negative bars and axis. */
function joinBar(base: Extract<ConditionalFormat, { type: 'dataBar' }>, x: Element, ctx: Context): ConditionalFormat {
  const d = element(x.body ?? '', 'dataBar')
  if (!d) return base
  const a = attributes(d.attrs)
  const b = d.body ?? ''
  const vs = elements(b, 'cfvo').map((v) => value(v, base.ranges[0], ctx))
  const { excel2007: _, ...out } = base
  if (vs.length === 2) [out.min, out.max] = vs
  delete out.minLength
  delete out.maxLength
  Object.assign(out, lengths(a, 0, 100))
  if (!flag(a.gradient, true)) out.gradient = false
  const color = (local: string) => colorOf(element(b, local), ctx.palette)
  if (flag(a.border)) out.border = color('borderColor') ?? base.color
  if (a.direction === 'leftToRight' || a.direction === 'rightToLeft') out.direction = a.direction
  if (!flag(a.negativeBarColorSameAsPositive)) out.negativeColor = color('negativeFillColor') ?? '#FF0000'
  if (out.border && !flag(a.negativeBarBorderColorSameAsPositive, true)) out.negativeBorder = color('negativeBorderColor') ?? '#FF0000'
  if (a.axisPosition === 'middle' || a.axisPosition === 'none') out.axis = a.axisPosition
  const axis = color('axisColor')
  if (axis && axis !== '#000000') out.axisColor = axis
  return out
}

/** The sheet's conditional formats in priority order, from its tail (the worksheet after `<sheetData>`). */
export function readConditional(tail: string, sheet: string, dxfs: Dxf[], palette: Palette, issues: Issues): ConditionalFormat[] | undefined {
  const ctx: Context = { sheet, dxfs, palette, issues }
  const { blocks, rest } = split(tail)
  const found: Read[] = []
  for (const b of blocks) {
    const a = attributes(b.attrs)
    const sqref = ranges(a.sqref ?? '', ctx)
    if (flag(a.pivot)) {
      issues.add('conditionalFormat', "a pivot table's conditional formatting is not in the model", { sheet, ...(sqref[0] ? { cell: sqref[0] } : {}) })
      continue
    }
    if (!sqref.length) continue
    for (const e of elements(b.body ?? '', 'cfRule')) {
      const r = rule(e, sqref, false, ctx)
      if (r) found.push(r)
    }
  }
  // Excel 2010's rules: in the extension of the sheet's own extension list
  const ext = elements(element(rest, 'extLst')?.body ?? '', 'ext').find((e) => attributes(e.attrs).uri?.toUpperCase() === X14_CF.toUpperCase())
  for (const b of elements(ext?.body ?? '', 'conditionalFormatting')) {
    const body = b.body ?? ''
    const sqref = ranges(readText(element(body, 'sqref')?.body ?? ''), ctx)
    if (!sqref.length) continue
    for (const e of elements(body, 'cfRule')) {
      const id = attributes(e.attrs).id
      const bar = id === undefined ? undefined : found.find((x) => x.id === id && x.rule.type === 'dataBar')
      if (bar) {
        bar.rule = joinBar(bar.rule as Extract<ConditionalFormat, { type: 'dataBar' }>, e, ctx)
        continue
      }
      const r = rule(e, sqref, true, ctx)
      if (r) found.push(r)
    }
  }
  found.forEach((x, i) => (x.order = i))
  found.sort((a, b) => a.priority - b.priority || a.order - b.order)
  return found.length ? found.map((x) => plain(x.rule)) : undefined
}

/** A rule without a formula that is the one Excel writes for it (`defaultFormula`), which the model makes itself. */
function plain(rule: ConditionalFormat): ConditionalFormat {
  if (rule.type === 'expression' || !('formula' in rule) || rule.formula === undefined) return rule
  const o = originOf(rule)
  if (rule.formula !== defaultFormula(rule, formatAddress(o.r, o.c))) return rule
  const { formula: _, ...rest } = rule
  return rest as ConditionalFormat
}
