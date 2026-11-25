/** The BIFF8 cache's dictionaries and records, stored in _SX_DB_CUR. */
import { escapeAttr } from '../xlsx/xml/text.js'
import { errorText } from '../xlsb/fields.js'
import type { Context } from '../xlsb/parts.js'
import type { Payload } from '../xlsb/records.js'
import type { Xml } from '../xlsb/xml.js'
import { virtualPath } from './external.js'
import { CacheFormulas } from './pivot-formulas.js'
import { BUILTIN_NAMES, CONTINUE, EOF, records, reference, unicode } from './records.js'

export interface CacheSource {
  id: number
  type: number
  connection?: number
  olap?: boolean
  areas: SourceArea[]
  consolidationCount?: number
  enableRefresh?: boolean
  minRefreshVersion?: number
  retainItems?: number
}

interface SourceLocation {
  sheet?: string
  book?: string
}

type SourceArea = ({ range: string } | { name: string }) & SourceLocation

/** The optional file after a DConName or the required file and sheet after a DConRef. */
function sourceLocation(payload: Payload): SourceLocation | undefined {
  const length = payload.u16()
  if (!length) return {}
  const target = unicode(payload, length)
  if (target.charCodeAt(0) === 2) return { sheet: target.slice(1) }
  const location = virtualPath(target)
  return location && { book: location.target, sheet: location.sheet }
}

export function cacheSource(source: CacheSource, type: number, payload: Payload): boolean {
  if (type === 0x00e3) source.type = payload.u16()
  else if (type === 0x00d0) source.consolidationCount = payload.u16()
  else if (type === 0x0051) {
    const range = reference(payload, true)
    const location = sourceLocation(payload)
    if (location?.sheet) source.areas.push({ range, ...location })
  } else if (type === 0x0052) {
    const name = unicode(payload)
    const location = sourceLocation(payload)
    if (location && name) source.areas.push({ name, ...location })
  } else if (type === 0x01b5) {
    const name = BUILTIN_NAMES[payload.u8()]
    payload.skip(3)
    const location = sourceLocation(payload)
    if (location && name) source.areas.push({ name: '_xlnm.' + name, ...location })
  } else if (type === 0x0864 && payload.remaining >= 6) {
    const info = payload.copy()
    info.skip(4)
    if (info.u8() !== 3) return false
    const property = info.u8()
    if (property === 0x02 && info.remaining >= 12) {
      info.skip(6)
      const retain = info.i32()
      if (retain >= 0) source.retainItems = retain
      info.skip(1)
      source.minRefreshVersion = info.u8()
    } else if (property === 0x34 && info.remaining >= 4) {
      source.enableRefresh = Boolean(info.u32() & 1)
    } else return false
  } else return false
  return true
}

type Value = { type: string; value?: string | number }

function pivotValue(type: number, payload: Payload): Value | undefined {
  switch (type) {
    case 0x00c9: return { type: 'n', value: payload.f64() }
    case 0x00ca: return { type: 'b', value: payload.u16() ? 1 : 0 }
    case 0x00cb: return { type: 'e', value: errorText(payload.u16()) }
    case 0x00cc: return { type: 'n', value: payload.i16() }
    case 0x00cd: return { type: 's', value: unicode(payload) }
    case 0x00ce: {
      const year = payload.u16()
      const month = payload.u16()
      const day = payload.u8(), hour = payload.u8(), minute = payload.u8(), second = payload.u8()
      const two = (n: number) => String(n).padStart(2, '0')
      return { type: 'd', value: `${String(year).padStart(4, '0')}-${two(month)}-${two(day)}T${two(hour)}:${two(minute)}:${two(second)}` }
    }
    case 0x00cf: return { type: 'm' }
    default: return undefined
  }
}

function valueXml(value: Value, unused = false): string {
  const content = value.value === undefined ? '' : ` v="${escapeAttr(String(value.value))}"`
  return `<${value.type}${content}${unused ? ' u="1"' : ''}/>`
}

const GROUP_BY = ['range', 'seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years']

interface Group {
  parent?: number
  base?: number
  itemCount: number
  items: Value[]
  mappingCount: number
  mapping: number[]
  rangeFlags?: number
  bounds: Value[]
}

interface Field {
  name: string
  flags: number
  indexed: boolean
  calculated: boolean
  wide: boolean
  count: number
  group?: Group
  items: Value[]
  used?: Uint8Array
  types: Set<string>
  minNum: number
  maxNum: number
  minDate?: string
  maxDate?: string
  long?: boolean
  integer: boolean
}

function field(payload: Payload, index: number, sourceCount: number): Field {
  const flags = payload.u16()
  const parent = payload.u16()
  const base = payload.u16()
  payload.skip(2)
  const itemCount = payload.u16()
  const mappingCount = payload.u16()
  const count = payload.u16()
  const group = flags & 0x18 || itemCount || mappingCount ? {
    parent: flags & 8 ? parent : undefined,
    base: flags & 16 || index >= sourceCount ? base : undefined,
    itemCount, items: [], mappingCount, mapping: [], bounds: [],
  } : undefined
  return {
    name: unicode(payload), flags, indexed: Boolean(flags & 1), calculated: Boolean(flags & 0x8000), wide: Boolean(flags & 0x200),
    count, group, items: [], types: new Set(), minNum: Infinity, maxNum: -Infinity, integer: true,
  }
}

function groupValue(group: Group | undefined, value: Value): boolean {
  if (!group) return false
  if (group.items.length < group.itemCount) {
    group.items.push(value)
    return true
  }
  if (group.rangeFlags !== undefined && group.bounds.length < 3) {
    group.bounds.push(value)
    return true
  }
  return false
}

function groupXml(group: Group, fieldCount: number): string | undefined {
  const { parent, base, rangeFlags, bounds, mapping, items } = group
  if (items.length !== group.itemCount || mapping.length !== group.mappingCount || mapping.some((i) => i >= items.length)) return undefined
  if (parent !== undefined && parent >= fieldCount || base !== undefined && base >= fieldCount) return undefined
  const attrs = (parent === undefined ? '' : ` par="${parent}"`) + (base === undefined ? '' : ` base="${base}"`)
  let xml = `<fieldGroup${attrs}>`
  if (rangeFlags !== undefined) {
    const by = (rangeFlags >>> 2) & 7
    const kind = by ? 'd' : 'n'
    if (bounds.length !== 3 || bounds[0].type !== kind || bounds[1].type !== kind || bounds[2].type !== 'n') return undefined
    const unit = by ? 'Date' : 'Num'
    xml += `<rangePr groupBy="${GROUP_BY[by]}" autoStart="${rangeFlags & 1}" autoEnd="${(rangeFlags >>> 1) & 1}"`
      + ` start${unit}="${bounds[0].value}" end${unit}="${bounds[1].value}" groupInterval="${bounds[2].value}"/>`
  }
  if (mapping.length) xml += '<discretePr>' + mapping.map((i) => `<x v="${i}"/>`).join('') + '</discretePr>'
  if (items.length) xml += '<groupItems>' + items.map((item) => valueXml(item)).join('') + '</groupItems>'
  return xml + '</fieldGroup>'
}

function track(field: Field, value: Value): void {
  field.types.add(value.type)
  if (value.type === 'n' && typeof value.value === 'number') {
    field.minNum = Math.min(field.minNum, value.value)
    field.maxNum = Math.max(field.maxNum, value.value)
    field.integer &&= Number.isInteger(value.value)
  } else if (value.type === 'd' && typeof value.value === 'string') {
    if (field.minDate === undefined || value.value < field.minDate) field.minDate = value.value
    if (field.maxDate === undefined || value.value > field.maxDate) field.maxDate = value.value
  } else if (value.type === 's' && String(value.value).length > 255) field.long = true
}

function facts(field: Field, saved: boolean): string {
  // A cache without records can still describe the type of a field whose values had no dictionary.
  const unobserved = !saved && field.types.size === 0
  const number = field.types.has('n') || (unobserved && Boolean(field.flags & 0x20))
  const date = field.types.has('d') || (unobserved && Boolean(field.flags & 0x800))
  // Excel sets the bit after fNumField only when such a numeric field holds whole numbers.
  const integer = number && (unobserved ? Boolean(field.flags & 0x40) : field.integer)
  const attrs = [
    ['containsBlank', field.types.has('m')],
    // Error and boolean types are represented by their items in SpreadsheetML.
    ['containsString', field.types.has('s') || field.types.has('b') || field.types.has('e') || (!field.types.size && !number && !date)],
    ['containsDate', date],
    ['containsNumber', number],
    ['containsInteger', integer],
    ['longText', Boolean(field.long)],
  ] as const
  return attrs.map(([key, value]) => ` ${key}="${+value}"`).join('')
    + (field.minNum === Infinity ? '' : ` minValue="${field.minNum}" maxValue="${field.maxNum}"`)
    + (field.minDate === undefined ? '' : ` minDate="${field.minDate}" maxDate="${field.maxDate}"`)
}

function sourcePart(source: CacheSource): { xml: string; books: string[] } | undefined {
  const books: string[] = []
  const attrs = (area: SourceArea) => {
    const what = 'name' in area ? ` name="${escapeAttr(area.name)}"` : ` ref="${area.range}"`
    const sheet = area.sheet === undefined ? '' : ` sheet="${escapeAttr(area.sheet)}"`
    if (area.book === undefined) return what + sheet
    const link = ` r:id="source${books.length}"`
    books.push(area.book)
    return what + sheet + link
  }
  if (source.type === 1 && source.areas.length === 1) {
    const area = source.areas[0]
    return { xml: `<cacheSource type="worksheet"><worksheetSource${attrs(area)}/></cacheSource>`, books }
  }
  if (source.type === 2) {
    const connection = source.connection === undefined ? '' : ` connectionId="${source.connection}"`
    return { xml: `<cacheSource type="external"${connection}/>`, books }
  }
  if (source.type === 0x10) return { xml: '<cacheSource type="scenario"/>', books }
  if (source.type === 4 && source.areas.length && source.areas.length === source.consolidationCount
    && source.areas.every((area) => 'range' in area)) {
    const sets = source.areas.map((area) => `<rangeSet${attrs(area)}/>`).join('')
    return { xml: `<cacheSource type="consolidation"><consolidation><rangeSets>${sets}</rangeSets></consolidation></cacheSource>`, books }
  }
  return undefined
}

export function cacheParts(source: CacheSource, bytes: Uint8Array, def: Xml, recordsXml: Xml, context: Context): { saved: boolean; names: string[]; books: string[] } | undefined {
  const input = sourcePart(source)
  if (!input) return undefined
  let save = false
  let sourceCount = 0
  let fieldCount: number | undefined
  let recordCount: number | undefined
  let expectedId: number | undefined
  let flags = 0
  const fields: Field[] = []
  const formulas = new CacheFormulas()
  const stored: Value[] = []
  let row: Value[] = []
  let rowCount = 0
  let terminated = false
  const addIndexes = (): boolean => {
    while (row.length < sourceCount && fields[row.length].indexed) {
      const next = stored[row.length]
      if (!next) return false
      row.push(next)
    }
    if (row.length === sourceCount) {
      recordsXml.add('<r>')
      for (const value of row) recordsXml.add(valueXml(value))
      recordsXml.add('</r>')
      rowCount++
      row = []
      stored.length = 0
    }
    return true
  }
  recordsXml.add('<pivotCacheRecords>')
  for (const { type, payload } of records(bytes, 0, bytes.length)) {
    if (type === EOF) { terminated = true; break }
    if (type === 0x00c6) {
      recordCount = payload.u32()
      expectedId = payload.u16()
      flags = payload.u16()
      save = Boolean(flags & 1)
      payload.skip(2)
      sourceCount = payload.u16()
      fieldCount = payload.u16()
      continue
    }
    if (type === 0x00c7) {
      fields.push(field(payload, fields.length, sourceCount))
      continue
    }
    const last = fields.at(-1)
    if (formulas.take(type, payload, last?.calculated ? fields.length - 1 : undefined)) continue
    const group = last?.group
    if (type === 0x00d8) {
      if (!group || group.rangeFlags !== undefined) return undefined
      group.rangeFlags = payload.u16()
      continue
    }
    if (type === 0x00d9 || (type === CONTINUE && group && group.mapping.length < group.mappingCount)) {
      if (!group || payload.remaining % 2 || group.mapping.length + payload.remaining / 2 > group.mappingCount) return undefined
      while (payload.remaining) group.mapping.push(payload.u16())
      continue
    }
    const value = pivotValue(type, payload)
    if (value && last) {
      if (groupValue(group, value)) continue
      if (last.items.length < last.count) {
        last.items.push(value)
        track(last, value)
        continue
      }
    }
    if (type === 0x00c8) {
      if (row.length || stored.length || !sourceCount || fields.length < sourceCount) return undefined
      for (let i = 0; i < sourceCount; i++) {
        const field = fields[i]
        if (!field.indexed) continue
        const index = field.wide ? payload.u16() : payload.u8()
        if (index >= field.items.length) return undefined
        stored[i] = { type: 'x', value: index }
        const used = (field.used ??= new Uint8Array(field.items.length))
        used[index] = 1
      }
      if (!addIndexes()) return undefined
    } else if (value) {
      const field = fields[row.length]
      if (!sourceCount || row.length >= sourceCount || !field || field.indexed) return undefined
      row.push(value)
      track(field, value)
      if (!addIndexes()) return undefined
    } else if ([0x00c1, 0x00c2, 0x00c3].includes(type)) return undefined
  }
  recordsXml.add('</pivotCacheRecords>')
  if (!terminated || expectedId !== source.id || !fields.length || fields.length !== fieldCount || sourceCount > fields.length
    || row.length || stored.length || (save && rowCount !== recordCount)) return undefined
  const calculated = formulas.compile(fields, context)
  if (!calculated || fields.some((field, i) => field.calculated && !calculated.fields.has(i))) return undefined
  for (const { field, item } of calculated.items) {
    // A calculated item cannot occur in the source rows. It is still a live member of the field.
    const used = fields[field].used
    if (used) used[item] = 1
  }
  const refresh = flags & 4 ? ' refreshOnLoad="1"' : ''
  // For level 3 and later, SXDB requires the legacy fEnableRefresh bit to be zero.
  const version = source.minRefreshVersion
  const enabled = source.enableRefresh ?? (Boolean(flags & 32) || (version !== undefined && version >= 3 && version !== 255))
  const retention = source.retainItems === undefined ? '' : ` missingItemsLimit="${source.retainItems}"`
  def.add(`<pivotCacheDefinition saveData="${+save}" enableRefresh="${+enabled}"${refresh}${retention}>`)
  def.add(input.xml + '<cacheFields>')
  for (const [i, field] of fields.entries()) {
    const formula = calculated.fields.get(i)
    const expression = formula === undefined ? '' : ` formula="${escapeAttr(formula)}"`
    def.add(`<cacheField name="${escapeAttr(field.name)}"${expression}>`)
    if (field.items.length || (i < sourceCount && !source.olap)) {
      def.add(`<sharedItems${facts(field, save)}>`)
      for (const [index, value] of field.items.entries()) def.add(valueXml(value, field.used?.[index] === 0))
      def.add('</sharedItems>')
    }
    if (field.group !== undefined) {
      const group = groupXml(field.group, fields.length)
      if (!group) return undefined
      def.add(group)
    }
    def.add('</cacheField>')
  }
  def.add('</cacheFields>')
  if (calculated.items.length) {
    def.add('<calculatedItems>')
    for (const { field, item, formula } of calculated.items) {
      def.add(`<calculatedItem formula="${escapeAttr(formula)}"><reference field="${field}"><x v="${item}"/></reference></calculatedItem>`)
    }
    def.add('</calculatedItems>')
  }
  // This synthetic part lets the SpreadsheetML reader report the OLAP metadata outside BookInput.
  if (source.olap) def.add('<cacheHierarchies/>')
  def.add('</pivotCacheDefinition>')
  return { saved: save, names: fields.map((field) => field.name), books: input.books }
}
