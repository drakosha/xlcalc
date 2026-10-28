/**
 * A sheet's conditional formatting as its `conditionalFormats`: the rules of `<conditionalFormatting>` with their
 * differential formats (`dxfs.ts`), and Excel 2010's rules in the sheet's extension list (`x14:conditionalFormatting`):
 * a data bar's 2010 part joined to its 2007 rule by their id, and the rules only 2010 has (its icon sets, custom icons,
 * formulas over other sheets) with their own priority and format (`x14:dxf`). Rules come in priority order. Formulas
 * come in the engine's spelling; what the model cannot hold is reported.
 */
import type { Color, ConditionalFormat, ConditionalStyle, ConditionalValue, IconSetName } from '../../contract.js'
import { CELL_IS, ICON_SETS, PERIODS, STYLED, VALUE_TYPES, WITH_VALUE } from '../../cond/check.js'
import { defaultFormula, originOf } from '../../cond/rules.js'
import { formatAddress } from '../../model.js'
import { areaText } from '../../structure/area.js'
import type { XlsxMessage } from '../issues.js'
import type { XlsxIssueCode } from '../types.js'
import { colorOf, type Palette } from '../styles/color.js'
import { readDxf, type Dxf } from '../styles/dxf.js'
import { Issues, type Where } from '../issues.js'
import { attributes, element, elements, flag, ownEntry, placed, type Element, type Placed } from '../xml/scan.js'
import { readText } from '../xml/text.js'
import { readFormula, type FormulaOptions } from './formulas.js'
import { sqrefRanges } from './refs.js'

/** The extension of a sheet that holds Excel 2010's conditional formatting. */
export const X14_CF = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}'

/** A rule read with its priority and the id that joins a data bar's two parts. */
interface Read {
  rule: ConditionalFormat
  priority: number
  order: number
  id?: string
}

interface Context {
  sheet: string
  dxfs: Dxf[]
  palette: Palette
  issues: Issues
  formulas: FormulaOptions
}

/** The ranges of a `sqref`, those the model takes, as the model writes them (`A1:A1048576`, Excel's whole column, is `A:A`). */
function ranges(sqref: string, ctx: Context): string[] {
  const out: string[] = []
  for (const { text, rect } of sqrefRanges(sqref)) {
    if (rect) out.push(areaText(rect))
    else ctx.issues.add('conditionalFormat', { key: 'condRangeOutside' }, { sheet: ctx.sheet, cell: text })
  }
  return out
}

/** A formula in the engine's spelling, reported where it does not parse. */
function formula(text: string, where: string, ctx: Context): string {
  return readFormula(readText(text), ctx.issues, { sheet: ctx.sheet, cell: where }, ctx.formulas).formula
}

/** A threshold: `val` (2007) or `<xm:f>` (2010) as its formula; null for a type the model does not have, or one without the value it needs. */
function value(e: Element, where: string, ctx: Context): ConditionalValue | null {
  const a = attributes(e.attrs)
  const type = (a.type ?? 'min') as ConditionalValue['type']
  if (!VALUE_TYPES.includes(type)) return null
  const f = e.body ? element(e.body, 'f') : undefined
  const text = f ? (f.body ?? '') : a.val
  if (!WITH_VALUE.has(type)) return { type }
  return text === undefined ? null : { type, value: formula(text, where, ctx) }
}

/** The thresholds of `<cfvo>` elements, or null when one of them is none (`value`). */
function values(body: string, where: string, ctx: Context): ConditionalValue[] | null {
  const out = elements(body, 'cfvo').map((v) => value(v, where, ctx))
  return out.includes(null) ? null : (out as ConditionalValue[])
}

/** The colours of a colour scale or data bar, in order. */
const colors = (body: string, local: string, ctx: Context): (Color | undefined)[] => elements(body, local).map((e) => colorOf(e, ctx.palette))

/** A style of a rule: its dxf by id (2007), or its own `x14:dxf` (2010). */
function styleOf(a: Record<string, string>, body: string, where: string, ctx: Context): ConditionalStyle | undefined {
  const own = element(body, 'dxf')
  const d = own ? readDxf(own, ctx.palette) : a.dxfId !== undefined ? ctx.dxfs[Number(a.dxfId)] : undefined
  if (!d) return undefined
  for (const x of d.lost) ctx.issues.add('conditionalFormat', { key: 'condStyleLost', part: x }, { sheet: ctx.sheet, cell: where })
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
  const lost = (why: XlsxMessage) => {
    ctx.issues.add('conditionalFormat', why, { sheet: ctx.sheet, cell: where })
    return null
  }
  // a rule with a style is read from its attributes and formulas
  if (STYLED.has(type)) {
    const own = fs[0] !== undefined ? { formula: fs[0] } : {}
    switch (type) {
      case 'cellIs': {
        const op = a.operator ?? 'equal'
        if (!(CELL_IS as readonly string[]).includes(op)) return lost({ key: 'condOperator', operator: op })
        const n = op === 'between' || op === 'notBetween' ? 2 : 1
        if (fs.length < n) return lost({ key: 'condNoFormulas' })
        return read(withStyle({ ...base, type: 'cellIs', operator: op as (typeof CELL_IS)[number], formulas: fs.slice(0, n) }))
      }
      case 'expression':
        return fs.length ? read(withStyle({ ...base, type: 'expression', formula: fs[0] })) : lost({ key: 'condNoFormula' })
      case 'containsText':
      case 'notContainsText':
      case 'beginsWith':
      case 'endsWith':
        return read(withStyle({ ...base, type, text: readText(a.text ?? ''), ...own }))
      case 'timePeriod': {
        const period = a.timePeriod ?? 'today'
        if (!(PERIODS as readonly string[]).includes(period)) return lost({ key: 'condPeriod', period })
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
        if (equal && sd) ctx.issues.add('conditionalFormat', { key: 'condAverage' }, { sheet: ctx.sheet, cell: where })
        return read(withStyle({ ...base, type, ...(!flag(a.aboveAverage, true) ? { below: true } : {}), ...(equal ? { equal: true } : sd >= 1 && sd <= 3 ? { stdDev: Math.trunc(sd) } : {}) }))
      }
      default:
        return read(withStyle({ ...base, type, ...own } as ConditionalFormat))
    }
  }
  if (type === 'colorScale') {
    const cs = element(body, 'colorScale')?.body ?? ''
    const vs = values(cs, where, ctx)
    const cl = colors(cs, 'color', ctx)
    if (!vs || vs.length < 2 || vs.length > 3 || cl.length !== vs.length || cl.some((c) => !c)) return lost({ key: 'condColorScale' })
    return read({ ...base, type, points: vs.map((v, i) => ({ value: v, color: cl[i]! })) })
  }
  if (type === 'dataBar') {
    const db = element(body, 'dataBar')
    const d = db?.body ?? ''
    const vs = values(d, where, ctx)
    const color = x14 ? colorOf(element(d, 'fillColor'), ctx.palette) : colors(d, 'color', ctx)[0]
    if (vs?.length !== 2) return lost({ key: 'condDataBar' })
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
    if (ownEntry(ICON_SETS, set) === undefined && set !== ('NoIcons' as string)) return lost({ key: 'condIconSet', set })
    const ib = is?.body ?? ''
    const vs = elements(ib, 'cfvo')
    const icons = elements(ib, 'cfIcon').map((c) => {
      const ca = attributes(c.attrs)
      return ca.iconSet === 'NoIcons' ? null : { set: ca.iconSet as IconSetName, index: Number(ca.iconId ?? 0) }
    })
    const n = ICON_SETS[set]
    const thresholds = values(ib, where, ctx)
    if (thresholds?.length !== n) return lost({ key: 'condIcons' })
    const reverse = flag(ia.reverse)
    const points = vs.map((v, i) => {
      const gte = flag(attributes(v.attrs).gte, true)
      // a custom icon that is the set's own at its place is no custom icon
      const mine = icons[i]?.set === set && icons[i]?.index === (reverse ? n - 1 - i : i)
      const icon = flag(ia.custom) && !mine ? icons[i] : undefined
      return { value: thresholds[i], ...(gte ? {} : { gte: false }), ...(icon !== undefined ? { icon } : {}) }
    })
    const r: ConditionalFormat = { ...base, type, icons: set, points }
    if (reverse) r.reverse = true
    if (!flag(ia.showValue, true)) r.showValue = false
    return read(r)
  }
  return lost({ key: 'condType', type })
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
  // thresholds the model does not have leave the 2007 ones
  const vs = values(b, base.ranges[0], ctx)
  const { excel2007: _, ...out } = base
  if (vs?.length === 2) [out.min, out.max] = vs
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

/**
 * A rule of the file as the reader met it (`readConditional`'s `sources`): where its `cfRule` and its block lie in the
 * text read (`x14`: the extension list's, else the worksheet's tail), its priority, and what became of it.
 */
export interface RuleSource {
  x14: boolean
  start: number
  end: number
  block: { start: number; end: number; sqref: string }
  priority: number
  /** Its place in the list of rules read, undefined for a rule the model does not have; a 2010 part of a data bar, its 2007 rule's. */
  index?: number
  /** A data bar's id that joins its 2007 rule and its 2010 part. */
  id?: string
  /** The 2010 part of a data bar, laid over its 2007 rule, or beside it where neither is read. */
  joined?: boolean
  /** A pivot table's block (`pivot="1"`). */
  pivot?: boolean
  /** What the reader reported of it (`conditionalFormat` issues by key): a rule read with some lost, or not read. */
  lost: string[]
}

/** Issues passed on, the keys of the conditional formats' noted: what each rule loses. */
class Noted extends Issues {
  keys: string[] = []
  constructor(private readonly to: Issues) {
    super()
  }
  override add(code: XlsxIssueCode, message: XlsxMessage, where?: Where): void {
    if (code === 'conditionalFormat') this.keys.push(message.key)
    this.to.add(code, message, where)
  }
}

/**
 * The sheet's conditional formats in priority order, from its tail (the worksheet after `<sheetData>`, without its own
 * extension list, whose content is `ext`: Excel 2010's rules are there). `sources` gets every rule of the file, read or
 * not, in the order of the rules read.
 */
export function readConditional(tail: string, ext: string, sheet: string, dxfs: Dxf[], palette: Palette, issues: Issues, formulas: FormulaOptions, sources?: RuleSource[]): ConditionalFormat[] | undefined {
  const noted = sources ? new Noted(issues) : undefined
  const ctx: Context = { sheet, dxfs, palette, issues: noted ?? issues, formulas }
  const met: (RuleSource & { read?: Read; order: number })[] = []
  /** The keys noted since `from`. */
  const since = (from: number) => noted?.keys.slice(from) ?? []
  const at = () => noted?.keys.length ?? 0
  const note = (s: Omit<RuleSource, 'lost'>, lost: string[], read: Read | null) => {
    if (sources) met.push({ ...s, lost, order: met.length, ...(read ? { read } : {}) })
  }
  // 2007's blocks, with whatever prefix the file gives SpreadsheetML (2010's are in the extension list)
  const found: Read[] = []
  /** The ids of the data bars of 2007 the model does not have. */
  const lostBars = new Set<string>()
  for (const b of placed(tail, 'conditionalFormatting')) {
    const a = attributes(b.attrs)
    const k0 = at()
    const sqref = ranges(a.sqref ?? '', ctx)
    const own = since(k0)
    const pivot = flag(a.pivot)
    if (pivot) issues.add('conditionalFormat', { key: 'condPivot' }, { sheet, ...(sqref[0] ? { cell: sqref[0] } : {}) })
    const body = b.body ?? ''
    const inner = bodyAt(tail, b)
    const block = { start: b.start, end: b.end, sqref: a.sqref ?? '' }
    for (const e of placed(body, 'cfRule')) {
      const k = at()
      const r = pivot || !sqref.length ? null : rule(e, sqref, false, ctx)
      if (r) found.push(r)
      const id = element(e.body ?? '', 'id')?.body?.trim()
      if (!r && id) lostBars.add(id)
      note({ x14: false, start: inner + e.start, end: inner + e.end, block, priority: Number(attributes(e.attrs).priority ?? 0), ...(id ? { id } : {}), ...(pivot ? { pivot } : {}) }, [...own, ...since(k)], r)
    }
  }
  // Excel 2010's rules: in their extension of the sheet's own extension list
  const x14 = placed(ext, 'ext').find((e) => attributes(e.attrs).uri?.toUpperCase() === X14_CF.toUpperCase())
  const x14Body = x14?.body ?? ''
  const x14At = x14 ? bodyAt(ext, x14) : 0
  for (const b of placed(x14Body, 'conditionalFormatting')) {
    const body = b.body ?? ''
    const raw = readText(element(body, 'sqref')?.body ?? '')
    const k0 = at()
    const sqref = ranges(raw, ctx)
    const own = since(k0)
    const inner = x14At + bodyAt(x14Body, b)
    const block = { start: x14At + b.start, end: x14At + b.end, sqref: raw }
    for (const e of placed(body, 'cfRule')) {
      const k = at()
      const id = attributes(e.attrs).id
      const where = { x14: true, start: inner + e.start, end: inner + e.end, block, priority: Number(attributes(e.attrs).priority ?? 0), ...(id ? { id } : {}) }
      const bar = id === undefined || !sqref.length ? undefined : found.find((x) => x.id === id && x.rule.type === 'dataBar')
      if (bar) {
        bar.rule = joinBar(bar.rule as Extract<ConditionalFormat, { type: 'dataBar' }>, e, ctx)
        note({ ...where, joined: true }, [...own, ...since(k)], bar)
        continue
      }
      const r = sqref.length ? rule(e, sqref, true, ctx) : null
      if (r) found.push(r)
      note({ ...where, ...(!r && id !== undefined && lostBars.has(id) ? { joined: true } : {}) }, [...own, ...since(k)], r)
    }
  }
  found.forEach((x, i) => (x.order = i))
  found.sort((a, b) => a.priority - b.priority || a.order - b.order)
  if (sources) {
    // the rules in the order of their priority, the 2010 parts of data bars after them
    const index = new Map(found.map((x, i) => [x, i]))
    met.sort((a, b) => Number(!!a.joined) - Number(!!b.joined) || a.priority - b.priority || a.order - b.order)
    for (const { read, order: _, ...s } of met) sources.push(read ? { ...s, index: index.get(read)! } : s)
  }
  return found.length ? found.map((x) => plain(x.rule)) : undefined
}

/** Where the content of an element `placed` gave lies in its text. */
function bodyAt(xml: string, p: Placed): number {
  return xml.lastIndexOf('</', p.end - 1) - (p.body ?? '').length
}

/** A rule without a formula that is the one Excel writes for it (`defaultFormula`), which the model makes itself. */
function plain(rule: ConditionalFormat): ConditionalFormat {
  if (rule.type === 'expression' || !('formula' in rule) || rule.formula === undefined) return rule
  const o = originOf(rule)
  if (rule.formula !== defaultFormula(rule, formatAddress(o.r, o.c))) return rule
  const { formula: _, ...rest } = rule
  return rest as ConditionalFormat
}
