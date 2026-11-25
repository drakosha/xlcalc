import { DYNAMIC_KINDS, SORT_BY } from '../filter/check.js'
import { parseRange, rectText } from '../structure/area.js'
import { escapeAttr } from '../xlsx/xml/text.js'
import { errorText, range, rkNumber } from '../xlsb/fields.js'
import { dateGroup, iconAttrs } from '../xlsb/filter.js'
import type { Payload } from '../xlsb/records.js'
import { dxfn } from './conditional.js'
import { hiddenDropdown } from './drawing.js'
import { CONTINUE, unicode, type RecordInfo } from './records.js'
import type { Workbook, Sheet } from './workbook.js'

const OPERATORS = ['', 'lessThan', 'equal', 'lessThanOrEqual', 'greaterThan', 'notEqual', 'greaterThanOrEqual']
const CONTINUE_FRT12 = 0x087f

interface Condition {
  type: number
  operator: string
  value?: string
  textLength?: number
  literal?: boolean
}

function readCondition(payload: Payload, extended = false): Condition {
  const type = payload.u8()
  const operator = OPERATORS[payload.u8()] ?? ''
  const data = payload.section(8)
  switch (type) {
    case 2: return { type, operator, value: String(rkNumber(data.u32())) }
    case 4: return { type, operator, value: String(data.f64()) }
    case 6:
      if (!extended) data.skip(4)
      return { type, operator, textLength: data.u8(), literal: Boolean(data.u8()) }
    case 8: {
      const value = data.u8()
      return { type, operator, value: data.u8() ? errorText(value) : value ? 'TRUE' : 'FALSE' }
    }
    case 0x0c: return { type, operator: 'equal', value: '' }
    case 0x0e: return { type, operator: 'notEqual', value: ' ' }
    default: return { type, operator }
  }
}

function readTextCondition(item: Condition, payload: Payload): void {
  if (item.textLength !== undefined) item.value = unicode(payload, item.textLength)
}

function usable(item: Condition): boolean {
  return Boolean(item.operator) && item.value !== undefined
}

function values(items: Condition[], dates: string[] = []): string {
  const blank = items.some((item) => item.type === 0x0c)
  const entries = items.filter((item) => item.type !== 0x0c).map((item) => `<filter val="${escapeAttr(item.value!)}"/>`)
  return '<filters' + (blank ? ' blank="1"' : '') + '>' + entries.join('') + dates.join('') + '</filters>'
}

function custom(items: Condition[], and: boolean): string {
  const tests = items.map((item) => {
    const operator = item.type === 0x0c ? '' : ` operator="${item.operator}"`
    return `<customFilter${operator} val="${escapeAttr(item.value!)}"/>`
  })
  return '<customFilters' + (and ? ' and="1"' : '') + '>' + tests.join('') + '</customFilters>'
}

export interface SavedColumn {
  column: number
  body: string
  hidden?: boolean
}

export function filterColumn(payload: Payload, tableColumn?: number): SavedColumn | undefined {
  const storedColumn = payload.u16()
  const column = tableColumn ?? storedColumn
  const flags = payload.u16()
  const items = [readCondition(payload), readCondition(payload)]
  if (flags & 16) {
    const top = Boolean(flags & 32)
    const cutoff = items[0]
    // Excel saves the resolved Top N boundary in doper1, a field earlier readers could ignore.
    const saved = cutoff.type === 4 && cutoff.operator === (top ? 'greaterThanOrEqual' : 'lessThanOrEqual') && Number.isFinite(Number(cutoff.value))
    const threshold = saved ? ` filterVal="${cutoff.value}"` : ''
    return { column, body: `<top10 top="${+top}" percent="${(flags >>> 6) & 1}" val="${flags >>> 7}"${threshold}/>` }
  }
  const present = items.filter((item) => item.type !== 0)
  for (const item of present) readTextCondition(item, payload)
  if (!present.length) return { column, body: '' }
  if (!present.every(usable)) return undefined
  const simple = present.every((item) =>
    item.type === 0x0c || item.type === 6 && item.operator === 'equal' && !/[~?*]/.test(item.value!))
  const valueList = simple && Boolean(flags & 4) && (present.length === 1 || (flags & 9) === 9)
  return { column, body: valueList ? values(present) : custom(present, present.length === 2 && !(flags & 1)) }
}

type Continuations = Iterable<Payload>

function dynamicFilter(kind: number, items: Condition[]): string | undefined {
  const name = DYNAMIC_KINDS[kind < 3 ? kind : kind - 5]
  if (!name || kind >= 3 && kind < 8) return undefined
  const bounds = items.filter((item) => item.type === 4 && Number.isFinite(Number(item.value)))
  const min = bounds.find((item) => item.operator === 'greaterThanOrEqual')
  const max = bounds.find((item) => item.operator === 'lessThan')
  const value = kind <= 2 ? bounds[0] : min
  return `<dynamicFilter type="${name}"${value ? ` val="${value.value}"` : ''}${max && kind > 2 ? ` maxVal="${max.value}"` : ''}/>`
}

function extendedFilter(
  payload: Payload, more: Continuations, book: Workbook, previous: SavedColumn | undefined,
): SavedColumn | undefined {
  const column = payload.u16()
  const hidden = Boolean(payload.u32())
  const type = payload.u32()
  const dynamic = payload.u32()
  const count = payload.u32()
  const dateCount = payload.u32()
  payload.skip(26)
  let body: string | undefined
  if (type === 1 || type === 2) {
    const id = book.styles.dxf(dxfn(payload, true, 'filter'))
    body = `<colorFilter dxfId="${id}"${type === 2 ? ' cellColor="0"' : ''}/>`
  } else if (type === 3) body = `<iconFilter${iconAttrs(payload, '3Arrows')}/>`
  else if (type === 0) {
    const items: Condition[] = []
    const dates: string[] = []
    for (const data of more) {
      if (items.length < count) {
        const item = readCondition(data, true)
        readTextCondition(item, data)
        items.push(item)
      } else if (dates.length < dateCount) dates.push(dateGroup(data))
      else return undefined
    }
    if (items.length !== count || dates.length !== dateCount) return undefined
    if (dynamic) body = dynamicFilter(dynamic, items)
    else if (items.every(usable)) {
      if (items.every((item) => item.operator === 'equal' && item.literal !== false)) body = values(items, dates)
      else if (!dates.length && items.length <= 2) body = custom(items, previous?.body.startsWith('<customFilters and="1"') === true)
    }
  }
  if (body === undefined) return undefined
  return { column, body, ...(hidden ? { hidden } : {}) }
}

function sortCondition(payload: Payload, book: Workbook): { ref: string; xml: string } | undefined {
  const flags = payload.u16()
  const by = (flags >>> 1) & 15
  const ref = range(payload)
  let extra = ''
  if (by === 3) extra = iconAttrs(payload)
  else {
    const dxf = payload.u32()
    payload.skip(4)
    if (by === 1 || by === 2) {
      const index = book.styles.savedDxf(dxf)
      if (index === undefined) return undefined
      extra = ` dxfId="${index}"`
    }
  }
  const length = payload.u32()
  if (!by && length) extra = ` customList="${escapeAttr(unicode(payload, length))}"`
  if (by && !SORT_BY[by]) return undefined
  const xml = `<sortCondition ref="${ref}"${flags & 1 ? ' descending="1"' : ''}${by ? ` sortBy="${SORT_BY[by]}"` : ''}${extra}/>`
  return { ref, xml }
}

export interface SavedFilter {
  columns: Map<number, SavedColumn>
  sort?: string
}

export interface SheetFilters extends SavedFilter {
  enabled: boolean
  active: boolean
  rangeSort?: string
  tables: Map<number, SavedFilter>
}

export function readFilters(source: Generator<RecordInfo>, book: Workbook, sheet: Sheet): SheetFilters {
  const filter: SheetFilters = { enabled: false, active: false, columns: new Map(), tables: new Map() }
  const table = (id: number) => {
    let saved = filter.tables.get(id)
    if (!saved) {
      saved = { columns: new Map() }
      filter.tables.set(id, saved)
    }
    return saved
  }
  const lost = (text: string, ru: string) => book.losses.add('autoFilter', text, ru, sheet.name)
  const rect = sheet.filter ? parseRange(sheet.filter) : undefined
  let drawing: Uint8Array[] = []
  let next = source.next()
  function* continuations(): Generator<Payload> {
    while (!next.done && next.value.type === CONTINUE_FRT12) {
      const more = next.value.payload
      next = source.next()
      more.skip(12)
      yield more
    }
  }
  while (!next.done) {
    const { type, payload } = next.value
    next = source.next()
    if (type === 0x009b) filter.active = true
    else if (type === 0x009d) filter.enabled = true
    else if (type === 0x009e) {
      const column = filterColumn(payload)
      if (column) filter.columns.set(column.column, column)
      else lost('an AutoFilter criterion was left out', 'Условие автофильтра пропущено')
    } else if (type === 0x087e) {
      const header = payload.copy()
      header.skip(40)
      const owner = header.u32()
      const customView = header.bytes(16).some((byte) => byte !== 0)
      if (customView) continue
      const target = owner === 0xffffffff ? filter : table(owner)
      const columnHeader = payload.copy()
      columnHeader.skip(12)
      const index = columnHeader.u16()
      payload.skip(12)
      const previous = target.columns.get(index)
      const column = extendedFilter(payload, continuations(), book, previous)
      if (column) target.columns.set(index, { ...column, hidden: column.hidden || previous?.hidden })
      else lost('an extended AutoFilter criterion was left out', 'Расширенное условие автофильтра пропущено')
    } else if (type === 0x0895) {
      payload.skip(12)
      const flags = payload.u16()
      const ref = range(payload)
      const count = payload.u32()
      const id = payload.u32()
      const place = (flags >>> 3) & 7
      const conditions: { ref: string; xml: string }[] = []
      let found = 0
      for (const data of continuations()) {
        const condition = sortCondition(data, book)
        if (condition) conditions.push(condition)
        found++
      }
      if (found !== count || found !== conditions.length || place > 2) {
        lost('a saved sort condition was left out', 'Сохранённое условие сортировки пропущено')
      }
      if (!conditions.length || place > 2) continue
      const attrs = (flags & 1 ? ' columnSort="1"' : '') + (flags & 2 ? ' caseSensitive="1"' : '') + (flags & 4 ? ' sortMethod="stroke"' : '')
      const area = parseRange(ref)!
      const key = parseRange(conditions[0].ref)!
      // BIFF's table and filter sort ranges include the header. The sort key covers the data rows.
      const hasHeader = place !== 0 && !(flags & 1) && key.r1 === area.r1 + 1 && key.r2 <= area.r2
      const sorted = hasHeader ? rectText({ ...area, r1: key.r1, r2: key.r2 }) : ref
      const sort = `<sortState ref="${sorted}"${attrs}>${conditions.map((condition) => condition.xml).join('')}</sortState>`
      if (place === 0) filter.rangeSort = sort
      else if (place === 1) table(id).sort = sort
      else filter.sort = sort
    }
    if (type === 0x00ec) drawing = rect ? [payload.bytes()] : []
    else if (type === CONTINUE && drawing.length) drawing.push(payload.bytes())
    else {
      if (type === 0x005d && rect && drawing.length) {
        const position = hiddenDropdown(drawing, payload)
        if (position && position.row === rect.r1 && position.col >= rect.c1 && position.col <= rect.c2) {
          const column = position.col - rect.c1
          filter.columns.set(column, { column, body: filter.columns.get(column)?.body ?? '', hidden: true })
        }
      }
      drawing = []
    }
  }
  return filter
}

export function filterMarkup(columns: ReadonlyMap<number, SavedColumn>, overrides?: ReadonlyMap<number, SavedColumn>): string {
  const merged = new Map(columns)
  for (const [index, column] of overrides ?? []) {
    merged.set(index, { ...column, hidden: column.hidden || columns.get(index)?.hidden })
  }
  return [...merged.values()].sort((a, b) => a.column - b.column).map(({ column, body, hidden }) => {
    if (!body && !hidden) return ''
    return `<filterColumn colId="${column}"${hidden ? ' hiddenButton="1"' : ''}>${body}</filterColumn>`
  }).join('')
}
