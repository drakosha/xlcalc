/** Conditional formats in CondFmt/CF and their Excel 2007 extensions. */
import { quoteSheetIfNeeded } from '../a1.js'
import { parseRange } from '../structure/area.js'
import { escapeAttr, escapeText } from '../xlsx/xml/text.js'
import { ICONS } from '../xlsb/filter.js'
import { formulaText } from '../xlsb/formula.js'
import { rangeOrigin } from '../xlsb/future.js'
import type { Payload } from '../xlsb/records.js'
import { HORIZONTAL, LINES, PATTERNS, UNDERLINE, VERTICAL } from '../xlsb/styles.js'
import type { Xml } from '../xlsb/xml.js'
import { parsedFormula } from './formula.js'
import { reference, unicode, type RecordInfo } from './records.js'
import { color, extensions, fullColor } from './styles.js'
import type { Sheet, Workbook } from './workbook.js'

const TESTS = ['', 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual']
const VALUES = ['', 'num', 'min', 'max', 'percent', 'percentile', '', 'formula']
const TEMPLATES: Readonly<Partial<Record<number, string>>> = {
  7: 'uniqueValues', 9: 'containsBlanks', 10: 'notContainsBlanks', 11: 'containsErrors', 12: 'notContainsErrors', 27: 'duplicateValues',
}

const TEXT_RULES = ['containsText', 'notContainsText', 'beginsWith', 'endsWith']
const DATE_RULES = ['today', 'tomorrow', 'yesterday', 'last7Days', 'lastMonth', 'nextMonth', 'thisWeek', 'nextWeek', 'lastWeek', 'thisMonth']

interface Rule {
  type: string
  formula: string
  sourceFormula?: string
  attrs: string
  priority?: number
  style?: string
  stop?: boolean
}

interface Group {
  refs: string[]
  rules: (Rule | undefined)[]
}

function dxfAlignment(payload: Payload, flags: number, options: number): string {
  const alignment = payload.u8()
  const rotation = payload.u8()
  const indent = payload.u8()
  payload.skip(1)
  const relative = payload.i32()
  let attrs = ''
  if (!(flags & 1)) attrs += ` horizontal="${HORIZONTAL[alignment & 7]}"`
  if (!(flags & 2)) attrs += ` vertical="${VERTICAL[(alignment >>> 4) & 7] ?? 'bottom'}"`
  if (!(flags & 4)) attrs += ` wrapText="${(alignment >>> 3) & 1}"`
  if (!(flags & 8)) attrs += ` textRotation="${rotation}"`
  if (!(flags & 16)) attrs += ` justifyLastLine="${alignment >>> 7}"`
  if (!(flags & 32)) attrs += relative === 255 ? ` indent="${indent & 15}"` : ` relativeIndent="${relative}"`
  if (!(flags & 64)) attrs += ` shrinkToFit="${(indent >>> 4) & 1}"`
  if (options & 0x8000 && !(flags & 0x80000000)) attrs += ` readingOrder="${indent >>> 6}"`
  return attrs ? `<alignment${attrs}/>` : ''
}

export function dxfn(payload: Payload, extended: boolean, use: 'rule' | 'filter' = 'rule'): string {
  const flags = payload.u32()
  const options = payload.u16()
  let number = '', font = '', borders = '', pattern = ''
  let foreground: string | undefined, background: string | undefined
  let fontColor: string | undefined
  if (flags & (1 << 25)) {
    if (options & 1) {
      payload.skip(2)
      number = `<numFmt numFmtId="164" formatCode="${escapeAttr(unicode(payload))}"/>`
    } else {
      payload.skip(1)
      number = `<numFmt numFmtId="${payload.u8()}"/>`
    }
  }
  if (flags & (1 << 26)) {
    const block = payload.section(118)
    const fontName = block.section(64)
    const length = fontName.u8()
    if (length) font += `<name val="${escapeAttr(unicode(fontName, length))}"/>`
    const height = block.i32()
    const bits = block.u32()
    const weight = block.u16()
    const vert = block.u16()
    const underline = block.u8()
    block.skip(3)
    const index = block.i32()
    block.skip(4)
    const ignored = block.u32()
    const ignoreVert = block.u32()
    const ignoreUnderline = block.u32()
    const ignoreWeight = block.u32()
    if (!(ignored & 2)) font += `<i val="${(bits >>> 1) & 1}"/>`
    if (!(ignored & 128)) font += `<strike val="${(bits >>> 7) & 1}"/>`
    if (!ignoreWeight) font += `<b val="${+(weight >= 700)}"/>`
    if (!ignoreUnderline) font += `<u val="${UNDERLINE[underline] ?? 'none'}"/>`
    if (!ignoreVert && vert) font += `<vertAlign val="${vert === 1 ? 'superscript' : 'subscript'}"/>`
    if (height > 0) font += `<sz val="${height / 20}"/>`
    if (index >= 0 && index !== 0x7fff) fontColor = color(index)
  }
  const aligned = flags & (1 << 27) ? dxfAlignment(payload, flags, options) : ''
  let format: { sides: number; colors: number } | undefined
  if (flags & (1 << 28)) {
    format = { sides: payload.u16(), colors: payload.u32() }
    payload.skip(2)
  }
  if (flags & (1 << 29)) {
    const pat = payload.u16() >>> 10
    const colors = payload.u16()
    if (!(flags & 0x10000) && PATTERNS[pat] && (pat !== 1 || use === 'filter')) pattern = ` patternType="${PATTERNS[pat]}"`
    if (!(flags & 0x20000)) foreground = color(colors & 127)
    if (!(flags & 0x40000)) background = color((colors >>> 7) & 127)
  }
  if (flags & (1 << 30)) payload.skip(2)
  let ext = new Map<number, string | number>()
  if (extended && payload.remaining) {
    payload.skip(6)
    ext = extensions(payload, payload.u16())
  }
  const fg = ext.get(4) ?? foreground
  const bg = ext.get(5) ?? background
  const paint = ext.get(13) ?? fontColor
  if (paint) font += `<color ${paint}/>`
  if (format) {
    for (const [i, side, property] of [[0, 'left', 9], [1, 'right', 10], [2, 'top', 7], [3, 'bottom', 8]] as const) {
      if (flags & (1 << (i + 10))) continue
      const line = LINES[(format.sides >>> (i * 4)) & 15]
      const index = (format.colors >>> (i * 7)) & 127
      if (line) borders += `<${side} style="${line}"><color ${ext.get(property) ?? color(index)}/></${side}>`
    }
  }
  const fill = pattern || fg || bg
    ? `<fill><patternFill${pattern}>${fg ? `<fgColor ${fg}/>` : ''}${bg ? `<bgColor ${bg}/>` : ''}</patternFill></fill>` : ''
  return number + (font ? `<font>${font}</font>` : '') + aligned + (borders ? `<border>${borders}</border>` : '') + fill
}

function sizedDxf(payload: Payload): string {
  const size = payload.u32()
  if (size === 0) {
    payload.skip(2)
    return ''
  }
  return dxfn(payload.section(size), true)
}

function convert(payload: Payload, size: number, book: Workbook, sheet: Sheet, refs: string[], relative = true): string | undefined {
  if (!size) return ''
  const { row, col } = rangeOrigin(refs)
  const result = formulaText(parsedFormula(payload.section(size), size), book.context, row, col, relative, sheet.name)
  return result && 'text' in result ? result.text : undefined
}

function threshold(payload: Payload, book: Workbook, sheet: Sheet, refs: string[]): { type?: string; value?: string } {
  const type = VALUES[payload.u8()]
  const size = payload.u16()
  let formula = convert(payload, size, book, sheet, refs)
  if (formula) {
    const local = quoteSheetIfNeeded(sheet.name) + '!'
    if (formula.startsWith(local)) {
      const localReference = formula.slice(local.length)
      if (parseRange(localReference.replaceAll('$', ''))) formula = localReference
    }
  }
  const value = size ? formula : type === 'min' || type === 'max' ? undefined : String(payload.f64())
  return { type, value }
}

function valueXml({ type, value }: ReturnType<typeof threshold>, equal = true): string {
  return `<cfvo type="${type ?? 'unknown'}"${value === undefined ? '' : ` val="${escapeAttr(value)}"`}${equal ? '' : ' gte="0"'}/>`
}

function visual(type: number, payload: Payload, book: Workbook, sheet: Sheet, refs: string[]): string | undefined {
  payload.skip(3)
  if (type === 3) {
    const thresholds = payload.u8()
    const colors = payload.u8()
    payload.skip(1)
    if (thresholds < 2 || thresholds > 3 || colors !== thresholds) return undefined
    let values = '', stops = ''
    for (let i = 0; i < thresholds; i++) {
      values += valueXml(threshold(payload, book, sheet, refs))
      payload.skip(8)
    }
    for (let i = 0; i < colors; i++) {
      payload.skip(8)
      stops += `<color ${fullColor(payload)}/>`
    }
    return `<colorScale>${values}${stops}</colorScale>`
  }
  if (type === 4) {
    const flags = payload.u8()
    const min = payload.u8()
    const max = payload.u8()
    const attrs = fullColor(payload)
    const a = valueXml(threshold(payload, book, sheet, refs))
    const b = valueXml(threshold(payload, book, sheet, refs))
    return `<dataBar minLength="${min}" maxLength="${max}" showValue="${(flags >>> 1) & 1}">${a}${b}<color ${attrs}/></dataBar>`
  }
  const count = payload.u8()
  const icon = ICONS[payload.u8()]
  const flags = payload.u8()
  if (!icon || count < 3 || count > 5) return undefined
  let values = ''
  for (let i = 0; i < count; i++) {
    const value = threshold(payload, book, sheet, refs)
    const equal = Boolean(payload.u8())
    payload.skip(4)
    values += valueXml(value, i === 0 || equal)
  }
  return `<iconSet iconSet="${icon}" showValue="${flags & 1 ? 0 : 1}" reverse="${(flags >>> 2) & 1}">${values}</iconSet>`
}

function applyTemplate(rule: Rule, template: number, params: Payload): Rule {
  const simple = TEMPLATES[template]
  if (simple) return { ...rule, type: simple, attrs: '', formula: '' }
  if (template === 5 && params.remaining >= 3) {
    const top = params.u8()
    const count = params.u16()
    return { ...rule, type: 'top10', attrs: ` rank="${count}" bottom="${top & 1 ? 0 : 1}" percent="${(top >>> 1) & 1}"`, formula: '' }
  }
  if (template === 8 && params.remaining >= 2) {
    const comparison = params.u16()
    const type = TEXT_RULES[comparison]
    if (!type) return rule
    const pattern = comparison < 2 && /SEARCH\(\s*"((?:[^"]|"")*)"/i.exec(rule.sourceFormula ?? '')
    const text = pattern ? pattern[1].replaceAll('""', '"') : ''
    return { ...rule, type, attrs: ` text="${escapeAttr(text)}"` }
  }
  const period = DATE_RULES[template - 15]
  if (period) return { ...rule, type: 'timePeriod', attrs: ` timePeriod="${period}"` }
  if (template === 25 || template === 26 || template === 29 || template === 30) {
    const below = template === 26 || template === 30
    const equal = template === 29 || template === 30
    const deviation = params.remaining >= 2 ? params.u16() : 0
    const attrs = (below ? ' aboveAverage="0"' : '') + (equal ? ' equalAverage="1"' : '')
      + (deviation ? ` stdDev="${deviation}"` : '')
    return { ...rule, type: 'aboveAverage', attrs, formula: '' }
  }
  return rule
}

function extendedRule(payload: Payload, book: Workbook, sheet: Sheet, refs: string[]): Rule | undefined {
  payload.skip(12)
  const kind = payload.u8()
  const compare = payload.u8()
  const first = payload.u16()
  const second = payload.u16()
  const style = sizedDxf(payload)
  const a = convert(payload, first, book, sheet, refs)
  const b = convert(payload, second, book, sheet, refs)
  if (a === undefined || b === undefined) return undefined
  const active = payload.u16()
  payload.skip(active)
  const flags = payload.u8()
  const priority = payload.u16()
  const template = payload.u16()
  const params = payload.section(payload.u8())
  const common = { style, priority, stop: Boolean(flags & 2) }
  if (kind === 3 || kind === 4 || kind === 6) {
    if (active) return undefined
    const body = visual(kind, payload, book, sheet, refs)
    return body ? { ...common, type: kind === 3 ? 'colorScale' : kind === 4 ? 'dataBar' : 'iconSet', attrs: '', formula: body } : undefined
  }
  if (TEMPLATES[template]) return { ...common, type: TEMPLATES[template]!, attrs: '', formula: '' }
  if (kind === 5) {
    payload.skip(3)
    const top = payload.u8()
    const count = payload.u16()
    return { ...common, type: 'top10', attrs: ` rank="${count}" bottom="${top & 1 ? 0 : 1}" percent="${(top >>> 1) & 1}"`, formula: '' }
  }
  const rule = withFormula(kind, compare, a, b, common)
  return rule ? applyTemplate(rule, template, params) : undefined
}

function withFormula(kind: number, operator: number, a: string, b: string, common: Partial<Rule>): Rule | undefined {
  if ((kind !== 1 && kind !== 2) || !a || (kind === 1 && !TESTS[operator])) return undefined
  const type = kind === 1 ? 'cellIs' : 'expression'
  const formula = [a, b].filter(Boolean).map((text) => `<formula>${escapeText(text)}</formula>`).join('')
  const attrs = kind === 1 ? ` operator="${TESTS[operator]}"` : ''
  return { ...common, type, attrs, formula, sourceFormula: a }
}

export function conditionalRules(
  source: Generator<RecordInfo>, book: Workbook, sheet: Sheet, xml: Xml, pivotPriorities: ReadonlySet<number>,
): void {
  const indexed = new Map<number, Group>()
  const groups: Group[] = []
  let group: Group | undefined
  let priority = 0
  let declined = false
  for (const { type, payload } of source) {
    if (type === 0x01b0 || type === 0x0879) {
      if (type === 0x0879) payload.skip(12)
      payload.skip(2)
      const id = payload.u16() >>> 1
      payload.skip(8)
      const count = payload.u16()
      const refs: string[] = []
      for (let i = 0; i < count; i++) refs.push(reference(payload))
      group = { refs, rules: [] }
      groups.push(group)
      indexed.set(id, group)
    } else if ((type === 0x01b1 || type === 0x087a) && group?.refs.length) {
      let rule: Rule | undefined
      if (type === 0x087a) rule = extendedRule(payload, book, sheet, group.refs)
      else {
        const kind = payload.u8(), compare = payload.u8()
        const first = payload.u16(), second = payload.u16()
        const style = dxfn(payload.section(payload.remaining - first - second), false)
        const a = convert(payload, first, book, sheet, group.refs)
        const b = convert(payload, second, book, sheet, group.refs)
        if (a !== undefined && b !== undefined) rule = withFormula(kind, compare, a, b, { style, stop: true })
      }
      group.rules.push(rule)
      declined ||= !rule
    } else if (type === 0x087b) {
      payload.skip(12)
      const extended = payload.u32()
      const target = indexed.get(payload.u16())
      if (extended) {
        group = target
        continue
      }
      const index = payload.u16()
      const rule = target?.rules[index]
      if (!target || !rule) { declined = true; continue }
      payload.skip(1)
      const template = payload.u8()
      rule.priority = payload.u16()
      const flags = payload.u8()
      rule.stop = Boolean(flags & 2)
      if (payload.u8()) rule.style = sizedDxf(payload)
      const params = payload.section(payload.u8())
      target.rules[index] = applyTemplate(rule, template, params)
    }
  }
  for (const { refs, rules } of groups) {
    let open: boolean | undefined
    for (const rule of rules) {
      if (!rule) continue
      priority++
      const pivot = rule.priority !== undefined && pivotPriorities.has(rule.priority)
      if (open !== pivot) {
        if (open !== undefined) xml.add('</conditionalFormatting>')
        xml.add(`<conditionalFormatting sqref="${refs.join(' ')}"${pivot ? ' pivot="1"' : ''}>`)
        open = pivot
      }
      const dxf = rule.style ? ` dxfId="${book.styles.dxf(rule.style)}"` : ''
      xml.add(`<cfRule type="${rule.type}" priority="${rule.priority ?? priority}"${rule.stop ? ' stopIfTrue="1"' : ''}${dxf}${rule.attrs}>${rule.formula}</cfRule>`)
    }
    if (open !== undefined) xml.add('</conditionalFormatting>')
  }
  if (declined) book.losses.add('conditionalFormat', 'some BIFF conditional formatting rules were left out', 'Часть правил условного форматирования BIFF пропущена', sheet.name)
}
