import { parseAddress } from '../model.js'
import { escapeAttr, escapeText } from '../xlsx/xml/text.js'
import { colorAttrs } from './fields.js'
import { ICONS } from './filter.js'
import { formula, formulaText, type Formula } from './formula.js'
import { future, guid, ranges } from './future.js'
import type { Part } from './parts.js'
import type { Payload } from './records.js'
import type { Xml } from './xml.js'

const TYPES = ['', 'cellIs', 'expression', 'colorScale', 'dataBar', 'top10', 'iconSet']
const TEMPLATES: Readonly<Partial<Record<number, string>>> = {
  0x07: 'uniqueValues', 0x09: 'containsBlanks', 0x0a: 'notContainsBlanks', 0x0b: 'containsErrors',
  0x0c: 'notContainsErrors', 0x19: 'aboveAverage', 0x1a: 'aboveAverage', 0x1b: 'duplicateValues', 0x1d: 'aboveAverage', 0x1e: 'aboveAverage',
}
const TEXT_RULES = ['containsText', 'notContainsText', 'beginsWith', 'endsWith']
const PERIODS = ['today', 'yesterday', 'last7Days', 'thisWeek', 'lastWeek', 'lastMonth', 'tomorrow', 'nextWeek', 'nextMonth', 'thisMonth']
const OPERATORS = ['', 'between', 'notBetween', 'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual']
const VALUES = ['', 'num', 'min', 'max', 'percent', 'percentile', '', 'formula', 'autoMin', 'autoMax']

export class Conditional {
  private row = 0
  private col = 0
  private visual = false
  private futureRefs = ''
  private barColors: string[] = []

  constructor(private readonly part: Part, private readonly extensions: Xml) {}

  take(type: number, payload: Payload): boolean {
    const { xml } = this.part
    let text: string
    switch (type) {
      case 461: {
        payload.u32()
        const pivot = payload.u32()
        const refs = ranges(payload)
        if (!refs.length) payload.fail('conditional formatting has no ranges')
        const corner = parseAddress(refs[0].split(':', 1)[0])
        this.row = corner.row
        this.col = corner.col
        text = `<conditionalFormatting sqref="${refs.join(' ')}"${pivot ? ' pivot="1"' : ''}>`
        break
      }
      case 462: text = '</conditionalFormatting>'; break
      case 463: text = this.rule(payload); break
      case 464: text = '</cfRule>'; break
      case 465: {
        this.visual = true
        const set = payload.u32()
        const flags = payload.u16()
        text = `<iconSet iconSet="${ICONS[set] ?? 'NoIcons'}" showValue="${flags & 2 ? 0 : 1}" reverse="${(flags >>> 2) & 1}">`
        break
      }
      case 466: this.visual = false; text = '</iconSet>'; break
      case 467: {
        this.visual = true
        const min = payload.u8()
        const max = payload.u8()
        const value = payload.u8()
        text = `<dataBar minLength="${min}" maxLength="${max}" showValue="${value}">`
        break
      }
      case 468: this.visual = false; text = '</dataBar>'; break
      case 469: this.visual = true; text = '<colorScale>'; break
      case 470: this.visual = false; text = '</colorScale>'; break
      case 471: text = this.threshold(payload); break
      case 564:
        if (!this.visual) return false
        text = `<color ${colorAttrs(payload)}/>`
        break
      case 1146:
        payload.u32()
        text = `<extLst><ext><x14:id>${guid(payload)}</x14:id></ext></extLst>`
        break
      default: return this.takeFuture(type, payload)
    }
    payload.done()
    xml.add(text)
    return true
  }

  private takeFuture(type: number, payload: Payload): boolean {
    let text: string
    switch (type) {
      case 1046: {
        const header = future(payload)
        payload.u32()
        const pivot = payload.u32()
        if (!header.ranges.length) payload.fail('extended conditional formatting has no range')
        const corner = parseAddress(header.ranges[0].split(':', 1)[0])
        this.row = corner.row
        this.col = corner.col
        this.futureRefs = header.ranges.join(' ')
        text = `<ext><x14:conditionalFormattings><x14:conditionalFormatting${pivot ? ' pivot="1"' : ''}>`
        break
      }
      case 1047:
        text = `<xm:sqref>${this.futureRefs}</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext>`
        break
      case 1048: text = this.rule(payload, future(payload).formulas); break
      case 1049: text = '</x14:cfRule>'; break
      case 1050: text = this.threshold(payload, future(payload).formulas[0], true); break
      case 1051: {
        payload.u32()
        const min = payload.u8()
        const max = payload.u8()
        const show = payload.u8()
        const direction = ['context', 'leftToRight', 'rightToLeft'][payload.u8()]
        const axis = ['automatic', 'middle', 'none'][payload.u8()]
        const flags = payload.u16()
        this.barColors = [
          'fillColor',
          ...(flags & 1 ? ['borderColor'] : []),
          ...(flags & 4 ? ['negativeFillColor'] : []),
          ...(flags & 8 ? ['negativeBorderColor'] : []),
          ...(axis === 'none' ? [] : ['axisColor']),
        ]
        text = `<x14:dataBar minLength="${min}" maxLength="${max}" showValue="${show}" direction="${direction}" axisPosition="${axis}"`
          + ` border="${flags & 1}" gradient="${(flags >>> 1) & 1}" negativeBarColorSameAsPositive="${flags & 4 ? 0 : 1}" negativeBarBorderColorSameAsPositive="${flags & 8 ? 0 : 1}">`
        break
      }
      case 1052: {
        payload.u32()
        const set = payload.u32()
        const flags = payload.u16()
        text = `<x14:iconSet iconSet="${ICONS[set] ?? 'NoIcons'}" showValue="${flags & 2 ? 0 : 1}" reverse="${(flags >>> 2) & 1}" custom="${flags & 1}">`
        break
      }
      case 1055: {
        payload.u32()
        const tag = this.barColors.shift() ?? 'color'
        text = `<x14:${tag} ${colorAttrs(payload)}/>`
        break
      }
      case 1112: {
        payload.u32()
        const set = payload.u32()
        const icon = payload.u32()
        text = `<x14:cfIcon iconSet="${ICONS[set] ?? 'NoIcons'}"${icon === 0xffffffff ? '' : ` iconId="${icon}"`}/>`
        break
      }
      case 1155: text = '</x14:iconSet>'; break
      case 1156: this.barColors = []; text = '</x14:dataBar>'; break
      case 1157: payload.u32(); text = '<x14:colorScale>'; break
      case 1158: text = '</x14:colorScale>'; break
      default: return false
    }
    payload.done()
    this.extensions.add(text)
    return true
  }

  private rule(payload: Payload, extended?: Formula[]): string {
    const { context, pending } = this.part
    const type = payload.u32()
    const template = payload.u32()
    const dxf = payload.u32()
    const priority = payload.i32()
    const parameter = payload.u32()
    payload.skip(8)
    const flags = payload.u16()
    const counts = [payload.u32(), payload.u32(), payload.u32()]
    let id: string | undefined
    if (extended) {
      payload.u32()
      const key = guid(payload)
      if (payload.u32()) id = key
    }
    const parameterText = payload.remaining ? payload.nullableWide() : undefined
    let name = TEMPLATES[template] ?? TYPES[type] ?? 'unknown'
    let attrs = dxf === 0xffffffff || extended ? '' : ` dxfId="${dxf}"`
    if (id) attrs += ` id="${id}"`
    if (flags & 2) attrs += ' stopIfTrue="1"'
    if (template === 8) {
      name = TEXT_RULES[parameter] ?? name
      attrs += ` text="${escapeAttr(parameterText ?? '')}"`
    }
    if (template >= 15 && template <= 24) {
      name = 'timePeriod'
      attrs += ` timePeriod="${PERIODS[parameter] ?? 'today'}"`
    }
    if (name === 'cellIs') attrs += ` operator="${OPERATORS[parameter] || 'unknown'}"`
    if (name === 'top10') attrs += ` rank="${parameter}" percent="${(flags >>> 4) & 1}" bottom="${(flags >>> 3) & 1}"`
    if (name === 'aboveAverage') attrs += ` aboveAverage="${(flags >>> 2) & 1}" equalAverage="${template >= 29 ? 1 : 0}" stdDev="${parameter}"`
    let text = `<${extended ? 'x14:' : ''}cfRule type="${name}" priority="${priority}"${attrs}>`
    const ownDxf = extended && dxf !== 0xffffffff ? context.dxfs14[dxf] : undefined
    if (ownDxf !== undefined) text += `<x14:dxf>${ownDxf}</x14:dxf>`
    else if (extended && dxf !== 0xffffffff) pending()
    for (const parsed of extended ?? counts.flatMap((count) => count ? [formula(payload)] : [])) {
      const value = formulaText(parsed, context, this.row, this.col, true)
      if (!value || !('text' in value)) pending()
      else {
        const tag = extended ? 'xm:f' : 'formula'
        text += `<${tag}>${escapeText(value.text)}</${tag}>`
      }
    }
    return text
  }

  private threshold(payload: Payload, extendedFormula?: Formula, extended = false): string {
    const kind = VALUES[payload.u32()] || 'unknown'
    const number = payload.f64()
    const hasGte = payload.u32()
    const gte = payload.u32()
    const count = payload.u32()
    let value: string | number | undefined
    if (kind === 'num' || kind === 'percent' || kind === 'percentile') value = number
    const parsed = extended ? extendedFormula : count ? formula(payload) : undefined
    if (parsed) {
      const decoded = formulaText(parsed, this.part.context, this.row, this.col, true)
      if (decoded && 'text' in decoded) value = decoded.text
      else this.part.pending()
    }
    const attrs = `type="${kind}"${hasGte ? ` gte="${gte}"` : ''}`
    if (extended) return `<x14:cfvo ${attrs}>${value === undefined ? '' : `<xm:f>${escapeText(String(value))}</xm:f>`}</x14:cfvo>`
    return `<cfvo ${attrs}${value === undefined ? '' : ` val="${escapeAttr(String(value))}"`}/>`
  }
}
