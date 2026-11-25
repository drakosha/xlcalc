/** Excel 2003 lists and later BIFF8 table records. */
import type { SheetInput } from '../contract.js'
import { formatAddress } from '../model.js'
import { parseRange } from '../structure/area.js'
import { copyFormula } from '../structure/copy.js'
import { escapeAttr } from '../xlsx/xml/text.js'
import type { Payload } from '../xlsb/records.js'
import type { Xml } from '../xlsb/xml.js'
import { filterColumn, filterMarkup, type SavedColumn, type SavedFilter } from './filter.js'
import { combinedPayload, reference, unicode, type RecordInfo } from './records.js'

const TOTALS = ['', 'average', 'count', 'countNums', 'max', 'min', 'sum', 'stdDev', 'var', 'custom']

export interface Table {
  name: string
  id: number
  range: string
  headers: number
  totals: number
  filter: boolean
  columns: string[]
  calculated: number[]
  filters: Map<number, SavedColumn>
  style?: string
}

function features(payload: Payload): Table | undefined {
  payload.skip(4)
  const range = reference(payload)
  const kind = payload.u16()
  if (kind !== 5) return undefined
  payload.skip(5)
  const refs = payload.u16()
  payload.skip(6 + refs * 8)
  const source = payload.u32()
  const id = payload.u32()
  const headers = payload.u32()
  const totals = payload.u32()
  payload.skip(4)
  const size = payload.u32()
  payload.skip(4)
  const flags = payload.u32()
  payload.skip(size - 32)
  const name = unicode(payload)
  const count = payload.u16()
  const filter = Boolean(flags & 2)
  const singleCell = Boolean(flags & 512)
  if (source === 1 || !count || count > 256 || !name) return undefined
  if (flags & 0x4000) unicode(payload)
  if (flags & 0x100000) unicode(payload)
  const table: Table = { id, name, range, headers, totals, filter, columns: [], calculated: [], filters: new Map() }
  for (let i = 0; i < count; i++) {
    const id = payload.u32()
    payload.skip(8)
    const total = TOTALS[payload.u32()]
    const styleTotal = payload.u32()
    payload.skip(4)
    const fieldFlags = payload.u32()
    const styleInsert = payload.u32()
    payload.skip(4)
    const sourceName = unicode(payload)
    const caption = singleCell ? sourceName : unicode(payload)
    payload.skip(styleTotal + styleInsert)
    if (filter) {
      const bytes = payload.u32()
      payload.skip(2)
      // The AutoFilter within a field uses a placeholder column. Its enclosing field gives the column.
      const criterion = bytes ? filterColumn(payload.section(bytes), i) : undefined
      const hidden = Boolean(fieldFlags & 2)
      if (criterion || hidden) table.filters.set(i, { column: i, body: criterion?.body ?? '', hidden })
    }
    if (fieldFlags & (4 | 8 | 256)) return undefined
    if (fieldFlags & 128) payload.skip(payload.u16())
    const label = fieldFlags & 1024 ? unicode(payload) : undefined
    if (fieldFlags & 2048) table.calculated.push(i)
    if (source === 3) payload.skip(4)
    if (!headers && !singleCell) {
      payload.skip(payload.u32()) // CachedDiskHeader's differential format
      if (fieldFlags & 512) unicode(payload)
    }
    table.columns.push(`<tableColumn id="${id}" name="${escapeAttr(caption)}"`
      + (total ? ` totalsRowFunction="${total}"` : '') + (label ? ` totalsRowLabel="${escapeAttr(label)}"` : '') + '/>')
  }
  return table
}

export function tables(source: Generator<RecordInfo>): { tables: Table[]; missed: boolean } {
  const found = new Map<number, Table>()
  let missed = false
  let next = source.next()
  while (!next.done) {
    const { type, payload } = next.value
    next = source.next()
    if (type === 0x0872 || type === 0x0878) {
      let data = payload
      if (!next.done && next.value.type === 0x0875) {
        const parts = [payload.bytes()]
        do {
          const more = next.value.payload
          more.skip(12)
          parts.push(more.bytes())
          next = source.next()
        } while (!next.done && next.value.type === 0x0875)
        data = combinedPayload(parts, 'Workbook table feature', type)
      }
      const table = features(data)
      if (table) found.set(table.id, table)
      else missed = true
    } else if (type === 0x0877) {
      payload.skip(12)
      const kind = payload.u16()
      const table = found.get(payload.u32())
      if (!table) continue
      if (kind === 1) {
        const flags = payload.u16()
        const style = unicode(payload)
        const name = style ? ` name="${escapeAttr(style)}"` : ''
        table.style = `<tableStyleInfo${name} showFirstColumn="${flags & 1}" showLastColumn="${(flags >>> 1) & 1}"`
          + ` showRowStripes="${(flags >>> 2) & 1}" showColumnStripes="${(flags >>> 3) & 1}"/>`
      } else if (kind === 2) {
        const name = unicode(payload)
        if (name) table.name = name
      }
    }
  }
  return { tables: [...found.values()], missed }
}

export function tablePart(table: Table, xml: Xml, saved: SavedFilter | undefined): string {
  const name = escapeAttr(table.name)
  xml.add(`<table id="${table.id}" name="${name}" displayName="${name}" ref="${table.range}" headerRowCount="${table.headers}" totalsRowCount="${table.totals}">`)
  if (table.filter && table.headers) {
    xml.add(`<autoFilter ref="${table.range}">${filterMarkup(table.filters, saved?.columns)}</autoFilter>`)
  }
  if (saved?.sort) xml.add(saved.sort)
  xml.add('<tableColumns>')
  for (const column of table.columns) xml.add(column)
  xml.add('</tableColumns>' + (table.style ?? '') + '</table>')
  return xml.text()
}

/** BIFF stores the calculated-column flag. Excel recovers the template from the bottom of the data column. */
export function restoreCalculatedColumns(saved: readonly Table[], sheet: SheetInput): void {
  if (!sheet.tables?.length) return
  const byName = new Map(sheet.tables.map((table) => [table.name, table]))
  for (const original of saved) {
    const table = byName.get(original.name)
    const rect = table && parseRange(table.range)
    if (!table || !rect || !original.calculated.length) continue
    const first = rect.r1 + original.headers
    const last = rect.r2 - original.totals
    if (first > last) continue
    for (const column of original.calculated) {
      const col = rect.c1 + column
      const bottom = formatAddress(last, col)
      const formula = sheet.cells[bottom]?.formula
      if (formula === undefined) continue
      const from = { sheet: sheet.name, cell: bottom }
      // Excel accepts one exception among the bottom five data cells, provided the bottom cell has the template.
      let exceptions = 0
      for (let row = Math.max(first, last - 4); row < last && exceptions < 2; row++) {
        const cell = formatAddress(row, col)
        if (sheet.cells[cell]?.formula !== copyFormula(formula, from, { sheet: sheet.name, cell })) exceptions++
      }
      if (exceptions < 2) {
        const cell = formatAddress(first, col)
        table.columns[column].formula = copyFormula(formula, from, { sheet: sheet.name, cell })
      }
    }
  }
}
