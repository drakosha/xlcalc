import { DATE_LEVELS, DYNAMIC_KINDS, SORT_BY } from '../filter/check.js'
import { escapeAttr } from '../xlsx/xml/text.js'
import { range } from './fields.js'
import type { Payload } from './records.js'
import type { Xml } from './xml.js'

export const ICONS = ['3Arrows', '3ArrowsGray', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols',
  '3Symbols2', '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray',
  '5Rating', '5Quarters', '3Stars', '3Triangles', '5Boxes']

const OPERATORS = ['', 'lessThan', 'equal', 'lessThanOrEqual', 'greaterThan', 'notEqual', 'greaterThanOrEqual']
const DATE_FIELDS = ['year', 'month', 'day', 'hour', 'minute', 'second']
const dx = (id: number) => id === 0xffffffff ? '' : ` dxfId="${id}"`

/** Filter records occur in worksheets, tables, and PivotTable filter definitions. */
export class Filters {
  private custom?: { prefix: string; and: boolean; conditions: string[] }
  private inFilter = false
  private skippedAlternate = 0

  constructor(private readonly xml: Xml) {}

  take(type: number, payload: Payload): boolean {
    if (this.inFilter && this.alternate(type, payload)) return true
    let text: string
    switch (type) {
      case 161: this.inFilter = true; text = `<autoFilter ref="${range(payload)}">`; break
      case 162: this.inFilter = false; text = '</autoFilter>'; break
      case 163: {
        const column = payload.u32()
        const flags = payload.u16()
        text = `<filterColumn colId="${column}"${flags & 1 ? ' hiddenButton="1"' : ''}${flags & 2 ? ' showButton="0"' : ''}>`
        break
      }
      case 164: text = '</filterColumn>'; break
      case 165:
        text = `<filters${payload.u32() ? ' blank="1"' : ''}>`
        payload.u32()
        break
      case 166: text = '</filters>'; break
      case 167: case 1177: {
        const prefix = type === 1177 ? 'x14:' : ''
        if (prefix) payload.skip(4)
        text = `<${prefix}filter val="${escapeAttr(payload.wide())}"/>`
        break
      }
      case 168: {
        const id = payload.u32()
        text = `<colorFilter${dx(id)} cellColor="${payload.u32() ? 1 : 0}"/>`
        break
      }
      case 169: case 1181: {
        const prefix = type === 1181 ? 'x14:' : ''
        if (prefix && payload.remaining === 12) payload.skip(4)
        const icon = iconAttrs(payload)
        text = `<${prefix}iconFilter${icon}/>`
        break
      }
      case 170: {
        const flags = payload.u8()
        const value = payload.f64()
        const applied = payload.f64()
        text = `<top10 top="${flags & 1}" percent="${(flags >>> 1) & 1}" val="${value}"${flags & 4 ? ` filterVal="${applied}"` : ''}/>`
        break
      }
      case 171: {
        const type = payload.u32()
        const flags = payload.u8()
        const minimum = payload.f64()
        const maximum = payload.f64()
        const kind = DYNAMIC_KINDS[type < 3 ? type : type - 5]
        if (!kind) payload.fail(`unknown dynamic filter type ${type}`)
        const min = flags & 1 && type > 0 && type < 24 ? ` val="${minimum}"` : ''
        const max = flags & 1 && type >= 8 && type < 24 ? ` maxVal="${maximum}"` : ''
        text = `<dynamicFilter type="${kind}"${min}${max}/>`
        break
      }
      case 172: case 1178:
        // Excel uses 1 for AND and 0 for OR in both records. MS-XLSB lists the values in reverse.
        this.custom = { prefix: type === 1178 ? 'x14:' : '', and: payload.u32() === 1, conditions: [] }
        text = ''
        break
      case 173: {
        const filter = this.custom
        if (!filter) return payload.fail('custom filter end without a start')
        const and = filter.and && filter.conditions.length === 2 ? ' and="1"' : ''
        const tag = filter.prefix + 'customFilters'
        text = `<${tag}${and}>${filter.conditions.join('')}</${tag}>`
        this.custom = undefined
        break
      }
      case 174: case 1180:
        if (!this.custom) return payload.fail('custom filter criteria without a start')
        if (this.custom.conditions.length === 2) return payload.fail('too many custom filter criteria')
        this.custom.conditions.push(custom(payload, this.custom.prefix))
        text = ''
        break
      case 175: text = dateGroup(payload); break
      case 530: {
        const flags = payload.u16()
        const ref = range(payload)
        text = `<sortState ref="${ref}"${flags & 1 && !this.inFilter ? ' columnSort="1"' : ''}${flags & 2 ? ' caseSensitive="1"' : ''}${flags & 4 ? ' sortMethod="stroke"' : ''}>`
        break
      }
      case 531: text = '</sortState>'; break
      case 532: case 1152: text = sortCondition(payload, type === 1152); break
      case 533: case 1153: text = ''; break
      default: return false
    }
    payload.done()
    if (text) this.xml.add(text)
    return true
  }

  private alternate(type: number, payload: Payload): boolean {
    if (this.skippedAlternate) {
      if (type === 37) this.skippedAlternate++
      else if (type === 38) this.skippedAlternate--
      return true
    }
    if (type === 38) return true
    if (type !== 37) return false
    const count = payload.u16()
    if (!count || count * 4 !== payload.remaining) return payload.fail(`invalid filter alternate version count ${count}`)
    let supported = false
    for (let i = 0; i < count; i++) {
      const version = payload.u16()
      const product = payload.u16()
      // The XLSX filter reader takes Excel 2010's Choice. Exact Excel 2007 records are its Fallback.
      const matchesVersion = version === 0x0e02 || Boolean(product & 0x8000) && version <= 0x0e02
      if ((product & 0x7fff) === 0 && matchesVersion) supported = true
    }
    if (!supported) this.skippedAlternate = 1
    return true
  }
}

function iconAttrs(payload: Payload): string {
  const set = payload.u32()
  const icon = payload.u32()
  if (set === 0xffffffff) return ' iconSet="NoIcons"'
  if (!ICONS[set]) return payload.fail(`unknown icon set ${set}`)
  return ` iconSet="${ICONS[set]}"${icon === 0xffffffff ? '' : ` iconId="${icon}"`}`
}

function custom(payload: Payload, prefix: string): string {
  const type = payload.u8()
  const rule = payload.u8()
  let value = ''
  switch (type) {
    case 4: value = String(payload.f64()); break
    case 6: payload.skip(8); value = payload.wide(); break
    case 8: value = payload.u8() ? 'TRUE' : 'FALSE'; payload.skip(7); break
    case 12: case 14: payload.skip(8); break
    default: return payload.fail(`unknown filter value type ${type}`)
  }
  const operator = OPERATORS[rule]
  if (!operator) return payload.fail(`unknown filter operator ${rule}`)
  return `<${prefix}customFilter${operator === 'equal' ? '' : ` operator="${operator}"`} val="${escapeAttr(value)}"/>`
}

function dateGroup(payload: Payload): string {
  const values = [payload.u16(), payload.u16(), payload.u32(), payload.u16(), payload.u16(), payload.u16()]
  payload.skip(6)
  const level = payload.u32()
  if (level >= DATE_LEVELS.length) return payload.fail(`unknown date grouping ${level}`)
  const attrs = DATE_FIELDS.slice(0, level + 1).map((name, i) => ` ${name}="${values[i]}"`).join('')
  return `<dateGroupItem${attrs} dateTimeGrouping="${DATE_LEVELS[level]}"/>`
}

function sortCondition(payload: Payload, future: boolean): string {
  if (future) payload.skip(4)
  const flags = payload.u16()
  const by = (flags >>> 1) & 15
  const ref = range(payload)
  let extras = ''
  if (by === 3) extras = iconAttrs(payload)
  else {
    if (by === 1 || by === 2) extras = dx(payload.u32())
    else payload.u32()
    payload.u32()
  }
  const list = payload.nullableWide()
  if (list !== undefined) extras += ` customList="${escapeAttr(list)}"`
  if (!SORT_BY[by]) return payload.fail(`unknown sort field ${by}`)
  return `<${future ? 'x14:' : ''}sortCondition ref="${ref}"${by ? ` sortBy="${SORT_BY[by]}"` : ''}${flags & 1 ? ' descending="1"' : ''}${extras}/>`
}
