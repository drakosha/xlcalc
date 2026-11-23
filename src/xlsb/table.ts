import { parseAddress } from '../model.js'
import { escapeAttr, escapeText } from '../xlsx/xml/text.js'
import { range } from './fields.js'
import { formula, formulaText } from './formula.js'
import type { Part } from './parts.js'
import { records, type Payload } from './records.js'

export interface TableIndex {
  id: number
  name: string
  columns: string[]
}

const TOTALS = ['none', 'average', 'count', 'countNums', 'max', 'min', 'sum', 'stdDev', 'var', 'custom']
const TABLE_TYPE: Readonly<Partial<Record<number, string>>> = { 0: 'worksheet', 2: 'xml', 3: 'queryTable' }
const attr = (key: string, value: string | number | undefined) => value === undefined ? '' : ` ${key}="${escapeAttr(String(value))}"`

function heading(payload: Payload) {
  const ref = range(payload)
  const type = payload.u32()
  const id = payload.u32()
  const headers = payload.u32()
  const totals = payload.u32()
  const flags = payload.u32()
  const dxfs = ['headerRowDxfId', 'dataDxfId', 'totalsRowDxfId', 'tableBorderDxfId', 'headerRowBorderDxfId', 'totalsRowBorderDxfId']
    .map((key) => { const value = payload.u32(); return value === 0xffffffff ? '' : attr(key, value) }).join('')
  const connection = payload.u32()
  const name = payload.nullableWide()
  const display = payload.nullableWide()
  const comment = payload.nullableWide()
  const headerStyle = payload.nullableWide()
  const dataStyle = payload.nullableWide()
  const totalsStyle = payload.nullableWide()
  payload.done()
  const attrs = attr('id', id) + attr('ref', ref) + attr('tableType', TABLE_TYPE[type]) + attr('name', name) + attr('displayName', display)
    + attr('comment', comment || undefined) + attr('headerRowCount', headers) + attr('totalsRowCount', totals)
    + attr('totalsRowShown', flags & 1) + attr('connectionId', connection || undefined) + dxfs
    + attr('headerRowCellStyle', headerStyle) + attr('dataCellStyle', dataStyle) + attr('totalsRowCellStyle', totalsStyle)
  return { id, ref, headers, name: display ?? name ?? '', attrs }
}

function column(payload: Payload) {
  const id = payload.u32()
  const totals = payload.u32()
  const dxfs = ['headerRowDxfId', 'dataDxfId', 'totalsRowDxfId']
    .map((key) => { const value = payload.u32(); return value === 0xffffffff ? '' : attr(key, value) }).join('')
  const field = payload.u32()
  payload.nullableWide()
  const name = payload.nullableWide() ?? ''
  const total = payload.nullableWide()
  const headerStyle = payload.nullableWide()
  const dataStyle = payload.nullableWide()
  const totalStyle = payload.nullableWide()
  payload.done()
  const attrs = attr('id', id) + attr('name', name) + attr('totalsRowFunction', TOTALS[totals])
    + attr('totalsRowLabel', total) + attr('queryTableFieldId', field || undefined) + dxfs
    + attr('headerRowCellStyle', headerStyle) + attr('dataCellStyle', dataStyle) + attr('totalsRowCellStyle', totalStyle)
  return { name, attrs }
}

export function tableIndex(bytes: Uint8Array, path: string): TableIndex | undefined {
  let index: TableIndex | undefined
  for (const { type, payload } of records(bytes, path)) {
    if (type === 343) {
      const { id, name } = heading(payload)
      index = { id, name, columns: [] }
    } else if (type === 347 && index) index.columns.push(column(payload).name)
  }
  return index
}

export function table({ bytes, path, xml, context, pending }: Part): void {
  let row = 0
  let firstCol = 0
  let col = 0
  for (const { type, payload } of records(bytes, path)) {
    switch (type) {
      case 343: {
        const read = heading(payload)
        const start = parseAddress(read.ref.split(':', 1)[0])
        row = start.row + read.headers
        firstCol = start.col
        xml.add(`<table xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"${read.attrs}>`)
        break
      }
      case 344: xml.add('</table>'); break
      case 345: xml.add(`<tableColumns count="${payload.u32()}">`); payload.done(); col = firstCol; break
      case 346: xml.add('</tableColumns>'); break
      case 347: xml.add(`<tableColumn${column(payload).attrs}>`); break
      case 348: xml.add('</tableColumn>'); col++; break
      case 351: case 352: {
        const flags = type === 351 ? payload.u8() : 0
        const formulaData = formula(payload)
        const decoded = formulaText(formulaData, context, row, col, true)
        payload.done()
        if (!decoded || !('text' in decoded)) pending()
        else {
          const tag = type === 351 ? 'calculatedColumnFormula' : 'totalsRowFormula'
          xml.add(`<${tag}${flags & 2 ? ' array="1"' : ''}>${escapeText(decoded.text)}</${tag}>`)
        }
        break
      }
      case 513: {
        const flags = payload.u16()
        const name = payload.nullableWide()
        payload.done()
        const options = ['showFirstColumn', 'showLastColumn', 'showRowStripes', 'showColumnStripes']
          .map((key, i) => attr(key, (flags >>> i) & 1)).join('')
        xml.add(`<tableStyleInfo${attr('name', name)}${options}/>`)
        break
      }
      case 161: xml.add(`<autoFilter ref="${range(payload)}">`); payload.done(); break
      case 162: xml.add('</autoFilter>'); break
      case 163: pending(); break
    }
  }
}
