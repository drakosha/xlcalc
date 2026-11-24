import { momentOf, ymd } from '../dates.js'
import { escapeAttr } from '../xlsx/xml/text.js'
import { errorText, range } from './fields.js'
import { formula, formulaText, type Formula } from './formula.js'
import { NS, type Part } from './parts.js'
import { records, type Payload } from './records.js'

interface Field {
  flags: number
  count: number
  source: boolean
}

interface CalculatedItem {
  formula: Formula
  names: (string | undefined)[]
  field?: number
  item?: number
}

const GROUP = ['range', 'seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years']
const itemSize: Readonly<Partial<Record<number, number>>> = { 20: 0, 21: 8, 22: 1, 23: 1, 24: 4, 25: 8, 26: 4 }
const runTypes: Readonly<Partial<Record<number, number>>> = { 1: 21, 2: 24, 16: 23, 32: 25 }
const attr = (key: string, v: string | number) => ` ${key}="${escapeAttr(String(v))}"`

const formulaName = (name: string): string => /^[\p{L}_][\p{L}\p{N}_.]*$/u.test(name) ? name : `'${name.replace(/'/g, "''")}'`

/** PtgSxName indexes the following BrtBeginPName collection, which in turn indexes the cache fields. */
function fieldFormulaNames(bytes: Uint8Array, path: string): { fields: string[]; references: number[][]; calculatedItems: boolean } {
  const fields: string[] = []
  const references: number[][] = []
  let current: number[] | undefined
  let calculatedItems = false
  for (const { type, payload } of records(bytes, path)) {
    if (type === 183) {
      payload.skip(20)
      fields.push(formulaName(payload.wide()))
      current = []
      references.push(current)
    } else if (type === 184) current = undefined
    else if (type === 255 && current) {
      current.push(payload.u32())
    } else if (type === 243) calculatedItems = true
  }
  return { fields, references, calculatedItems }
}

/** The definition supplies the types used by compact BrtPCRRecord rows. */
export function cacheDefinition(part: Part): Field[] {
  const { bytes, path, xml, pending, context } = part
  let fields: Field[] = []
  let field: Field | undefined
  const formulaNames = fieldFormulaNames(bytes, path)
  const itemsByField: (string | undefined)[][] = []
  let items: (string | undefined)[] | undefined
  let calculated: CalculatedItem | undefined
  for (const { type, payload } of records(bytes, path)) {
    switch (type) {
      case 179: {
        const refreshed = payload.u8()
        const minimum = payload.u8()
        const created = payload.u8()
        const flags = payload.u8()
        const retain = payload.i32()
        payload.f64()
        const extra = payload.u8()
        const count = payload.u32()
        const who = extra & 1 ? payload.wide() : undefined
        const id = extra & 2 ? payload.nullableWide() : undefined
        // BrtBeginPivotCacheDef ends with four unused bytes when the user is absent.
        if (who === undefined) payload.skip(4)
        xml.add(`<pivotCacheDefinition ${NS} refreshedVersion="${refreshed}" minRefreshableVersion="${minimum}" createdVersion="${created}"`
          + ` saveData="${flags & 1}" refreshOnLoad="${(flags >>> 2) & 1}" enableRefresh="${(flags >>> 4) & 1}" recordCount="${count}"`
          + (retain >= 0 ? attr('missingItemsLimit', retain) : '') + (who ? attr('refreshedBy', who) : '') + (id ? attr('r:id', id) : '') + '>')
        break
      }
      case 180: xml.add('</pivotCacheDefinition>'); break
      case 185: {
        const kind = ['worksheet', 'external', 'consolidation', 'scenario'][payload.u32()] ?? 'external'
        const connection = payload.u32()
        xml.add(`<cacheSource type="${kind}"${connection ? attr('connectionId', connection) : ''}>`)
        break
      }
      case 186: xml.add('</cacheSource>'); break
      case 187: xml.add(`<worksheetSource${sourceRange(payload)}/>`); break
      case 181:
        fields = []
        xml.add(`<cacheFields count="${payload.u32()}">`)
        break
      case 182: xml.add('</cacheFields>'); break
      case 183: {
        const flags = payload.u16()
        const format = payload.u32()
        const sql = payload.u16()
        payload.skip(12)
        const name = payload.wide()
        const caption = flags & 8 ? payload.wide() : undefined
        let formulaAttr = ''
        if (flags & 0x100) {
          const pivotNames = formulaNames.references[fields.length].map((index) => {
            const fieldName = formulaNames.fields[index]
            if (fieldName === undefined) return payload.fail(`pivot formula field index ${index} out of bounds`)
            return fieldName
          })
          const contextOfField = { ...context, pivotNames }
          const decoded = formulaText(formula(payload), contextOfField)
          if (!decoded || !('text' in decoded)) pending()
          else formulaAttr = attr('formula', decoded.text)
        }
        if (payload.remaining) payload.skip(payload.remaining)
        field = { flags: 0, count: 0, source: Boolean(flags & 4) }
        fields.push(field)
        if (formulaNames.calculatedItems) itemsByField.push([])
        xml.add(`<cacheField name="${escapeAttr(name)}" numFmtId="${format}" sqlType="${sql}" serverField="${flags & 1}"${caption ? attr('caption', caption) : ''}${formulaAttr}>`)
        break
      }
      case 184: xml.add('</cacheField>'); field = undefined; break
      case 189: {
        if (!field) return payload.fail('shared items outside a pivot cache field')
        field.flags = payload.u16()
        field.count = payload.u32()
        xml.add(shared(payload, field, context.date1904))
        items = itemsByField.at(-1)
        break
      }
      case 190: xml.add('</sharedItems>'); items = undefined; break
      case 191: cacheRun(part, payload, items); break
      case 195: xml.add('<cacheHierarchies/>'); continue
      case 207: payload.skip(2); xml.add('<consolidation>'); break
      case 208: xml.add('</consolidation>'); break
      case 215: xml.add(`<rangeSets count="${payload.u32()}">`); break
      case 216: xml.add('</rangeSets>'); break
      case 217:
        payload.skip(16)
        xml.add(`<rangeSet${sourceRange(payload)}/>`)
        break
      case 219: {
        const parent = payload.i32()
        const base = payload.i32()
        xml.add(`<fieldGroup${parent >= 0 ? attr('par', parent) : ''}${base >= 0 ? attr('base', base) : ''}>`)
        break
      }
      case 220: xml.add('</fieldGroup>'); break
      case 221:
        xml.add(`<groupItems count="${payload.u32()}">`)
        items = itemsByField.at(-1)
        if (items) items.length = 0
        break
      case 222: xml.add('</groupItems>'); items = undefined; break
      case 223: xml.add(groupRange(payload, context.date1904)); break
      case 225: xml.add(`<discretePr count="${payload.u32()}">`); break
      case 226: xml.add('</discretePr>'); break
      case 243: xml.add(`<calculatedItems count="${payload.u32()}">`); break
      case 244: xml.add('</calculatedItems>'); break
      case 245:
        payload.skip(4)
        calculated = { formula: formula(payload), names: [] }
        break
      case 246:
        if (!calculated) return payload.fail('calculated item end without a formula')
        writeCalculatedItem(part, calculated, itemsByField)
        calculated = undefined
        break
      case 251:
        if (!calculated) continue
        calculated.field ??= payload.u32()
        payload.skip(payload.remaining)
        break
      case 382:
        if (!calculated) continue
        if (calculated.item === undefined) calculated.item = payload.u32()
        else payload.skip(4)
        break
      case 255:
        if (!calculated) continue
        calculated.names.push(undefined)
        payload.skip(6)
        break
      case 259: {
        if (!calculated) continue
        const flags = payload.u8()
        const fieldIndex = payload.u32()
        const index = payload.i32()
        const name = formulaNames.fields[fieldIndex]
        const last = calculated.names.length - 1
        if (last < 0 || name === undefined) return payload.fail(`invalid calculated item field index ${fieldIndex}`)
        if (!(flags & 1)) {
          const values = itemsByField[fieldIndex]
          if (index < 0 || index >= values.length) return payload.fail(`calculated item index ${index} out of bounds`)
          const value = values[index]
          if (value !== undefined) calculated.names[last] = `${name}[${formulaName(value)}]`
        }
        break
      }
      case 1066: {
        payload.u32()
        payload.u8()
        const id = payload.u32()
        if (id) xml.add(`<extLst><ext><x14:pivotCacheDefinition pivotCacheId="${id}"/></ext></extLst>`)
        break
      }
      default:
        if (!cacheItem(part, type, payload, items)) continue
    }
    payload.done()
  }
  return fields
}

function sourceRange(payload: Payload): string {
  const named = payload.u8()
  const builtin = payload.u8()
  const flags = payload.u8()
  const sheet = flags & 2 ? payload.wide() : undefined
  const book = flags & 1 ? payload.nullableWide() : undefined
  const ref = named ? attr('name', (builtin ? '_xlnm.' : '') + payload.wide()) : attr('ref', range(payload))
  return ref + (sheet ? attr('sheet', sheet) : '') + (book ? attr('r:id', book) : '')
}

function writeCalculatedItem(part: Part, calculated: CalculatedItem, items: (string | undefined)[][]): void {
  const { field, item, names } = calculated
  if (field === undefined || item === undefined || names.some((name) => name === undefined)) {
    part.pending()
    return
  }
  if (field >= items.length || item >= items[field].length) return calculated.formula.tokens.fail('calculated item is outside the cache')
  const pivotNames = names as string[]
  const decoded = formulaText(calculated.formula, { ...part.context, pivotNames })
  if (!decoded || !('text' in decoded)) {
    part.pending()
    return
  }
  part.xml.add(`<calculatedItem${attr('formula', decoded.text)}><pivotArea><references><reference field="${field}"><x v="${item}"/>`
    + '</reference></references></pivotArea></calculatedItem>')
}

export function cacheRecords(part: Part, fields: Field[]): void {
  const { bytes, path, xml } = part
  const sourceFields = fields.filter((field) => field.source)
  let open = false
  let index = 0
  const finishRow = (payload: Payload) => {
    if (!open) return
    if (index !== sourceFields.length) payload.fail(`pivot record has ${index} values for ${sourceFields.length} source fields`)
    xml.add('</r>')
    open = false
  }
  for (const { type, payload } of records(bytes, path)) {
    switch (type) {
      case 193: xml.add(`<pivotCacheRecords count="${payload.u32()}">`); break
      case 194:
        finishRow(payload)
        xml.add('</pivotCacheRecords>')
        break
      case 34:
        finishRow(payload)
        xml.add('<r>')
        open = true
        index = 0
        break
      case 33:
        finishRow(payload)
        xml.add('<r>')
        for (const field of sourceFields) {
          const type = field.count ? 26 : field.flags & 0x40 ? 21 : (field.flags & 12) === 4 ? 25 : 24
          if (type === 26) xml.add(indexItem(payload, field))
          else xml.add(item(payload, type))
        }
        xml.add('</r>')
        break
      default:
        if (type < 20 || type > 26) continue
        if (!open || index >= sourceFields.length) return payload.fail('pivot record value outside its source fields')
        if (type === 26) xml.add(indexItem(payload, sourceFields[index]))
        else xml.add(item(payload, type))
        index++
    }
    payload.done()
  }
}

function indexItem(payload: Payload, field: Field | undefined): string {
  const index = payload.u32()
  if (field?.count && index >= field.count) payload.fail(`pivot shared item index ${index} out of bounds`)
  return `<x v="${index}"/>`
}

function cacheItem({ xml }: Part, type: number, payload: Payload, capture?: (string | undefined)[]): boolean {
  if (type >= 20 && type <= 26) {
    xml.add(item(payload, type, capture))
    return true
  }
  if (type < 27 || type > 32) return false
  const tag = item(payload, type - 7, capture)
  const flags = payload.u16()
  payload.skip(payload.remaining)
  xml.add(flags & 1 ? tag.replace('/>', ' u="1"/>') : tag)
  return true
}

function item(payload: Payload, type: number, capture?: (string | undefined)[]): string {
  let tag: string
  let value: string | number | undefined
  switch (type) {
    case 20: tag = 'm'; break
    case 21: tag = 'n'; value = payload.f64(); break
    case 22: tag = 'b'; value = payload.u8(); break
    case 23: tag = 'e'; value = errorText(payload.u8()); break
    case 24: tag = 's'; value = payload.wide(); break
    case 25: tag = 'd'; value = date(payload.u16(), payload.u16(), payload.u8(), payload.u8(), payload.u8(), payload.u8()); break
    case 26: tag = 'x'; value = payload.u32(); break
    default: return payload.fail(`unknown pivot cache item type ${type}`)
  }
  // A date or a positional reference needs the displayed label in its PivotTable.
  capture?.push(type === 24 || type === 21 || type === 23 ? String(value) : type === 22 ? (value ? 'TRUE' : 'FALSE') : undefined)
  return `<${tag}${value === undefined ? '' : attr('v', value)}/>`
}

function cacheRun({ xml }: Part, payload: Payload, capture?: (string | undefined)[]): void {
  const op = payload.u16()
  const type = runTypes[op]
  const count = payload.u32()
  if (type === undefined || count * itemSize[type]! > payload.remaining) return payload.fail(`invalid pivot cache run ${op}, length ${count}`)
  for (let i = 0; i < count; i++) xml.add(item(payload, type, capture))
}

function shared(payload: Payload, field: Field, date1904: boolean | undefined): string {
  const { flags, count } = field
  const names = ['containsSemiMixedTypes', 'containsNonDate', 'containsDate', 'containsString', 'containsBlank',
    'containsMixedTypes', 'containsNumber', 'containsInteger']
  const attrs = names.map((name, bit) => attr(name, (flags >>> bit) & 1)).join('')
  const format = flags & 4 && !(flags & 64) ? (n: number) => serialDate(n, date1904) : (n: number) => n
  const suffix = flags & 4 && !(flags & 64) ? 'Date' : 'Value'
  const bounds = flags & 0x100 ? attr(`min${suffix}`, format(payload.f64())) + attr(`max${suffix}`, format(payload.f64())) : ''
  return `<sharedItems count="${count}"${attrs} longText="${(flags >>> 9) & 1}"${bounds}>`
}

function groupRange(payload: Payload, date1904: boolean | undefined): string {
  const kind = GROUP[payload.u8()] ?? 'range'
  const flags = payload.u8()
  const start = payload.f64()
  const end = payload.f64()
  const interval = payload.f64()
  const bounds = flags & 4 ? attr('startDate', serialDate(start, date1904)) + attr('endDate', serialDate(end, date1904)) : attr('startNum', start) + attr('endNum', end)
  return `<rangePr groupBy="${kind}" autoStart="${flags & 1}" autoEnd="${(flags >>> 1) & 1}" groupInterval="${interval}"${bounds}/>`
}

function date(y: number, m: number, d: number, hour: number, minute: number, second: number): string {
  const pad = (n: number) => String(n).padStart(2, '0')
  return `${String(y).padStart(4, '0')}-${pad(m)}-${pad(d)}T${pad(hour)}:${pad(minute)}:${pad(second)}`
}

function serialDate(serial: number, date1904: boolean | undefined): string {
  const { day, second } = momentOf(serial)
  // Excel writes pivot cache bounds on the Gregorian calendar, including serials before March 1900.
  const stamp = new Date(Date.UTC(date1904 ? 1904 : 1899, date1904 ? 0 : 11, (date1904 ? 1 : 30) + day))
  const { y, m, d } = date1904 || day < 61
    ? { y: stamp.getUTCFullYear(), m: stamp.getUTCMonth() + 1, d: stamp.getUTCDate() }
    : ymd(day)
  return date(y, m, d, Math.floor(second / 3600), Math.floor(second / 60) % 60, second % 60)
}
