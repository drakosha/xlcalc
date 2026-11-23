import { escapeAttr, escapeText } from '../xlsx/xml/text.js'
import { address, errorText } from './fields.js'
import { NS, type Part } from './parts.js'
import { records, type Payload } from './records.js'

export interface ExternalSource {
  sheets: string[]
  names: string[]
  scopes: number[]
}

export function externalIndex(bytes: Uint8Array, path: string): ExternalSource {
  const index: ExternalSource = { sheets: [], names: [], scopes: [] }
  for (const { type, payload } of records(bytes, path)) {
    if (type === 359) index.sheets = sheetNames(payload)
    if (type === 577) {
      index.names.push(payload.wide())
      index.scopes.push(0)
    }
    if (type === 586 && index.scopes.length) {
      payload.u16()
      index.scopes[index.scopes.length - 1] = payload.u32()
    }
  }
  return index
}

export function externalLink({ bytes, path, xml }: Part, invalidDefinition: (name: string) => void): void {
  let sheets: string[] = []
  let namesOpen = false
  let tableOpen = false
  let book = false
  let row = -1
  let vm = 0
  let name = ''
  let scope = 0
  let ref = ''
  xml.add(`<externalLink ${NS}>`)
  for (const { type, payload } of records(bytes, path)) {
    if (type === 360) {
      const kind = payload.u16()
      const id = payload.nullableWide()
      const detail = payload.nullableWide()
      payload.done()
      if (kind !== 0) {
        if (kind === 1) xml.add(`<ddeLink ddeService="${escapeAttr(id ?? '')}" ddeTopic="${escapeAttr(detail ?? '')}"/>`)
        if (kind === 2) xml.add(`<oleLink r:id="${escapeAttr(id ?? '')}" progId="${escapeAttr(detail ?? '')}"/>`)
        break
      }
      book = true
      xml.add(`<externalBook${id ? ` r:id="${escapeAttr(id)}"` : ''}>`)
      continue
    }
    if (!book) continue
    switch (type) {
      case 359:
        sheets = sheetNames(payload)
        xml.add('<sheetNames>')
        for (const name of sheets) xml.add(`<sheetName val="${escapeAttr(name)}"/>`)
        xml.add('</sheetNames>')
        break
      case 577:
        if (!namesOpen) xml.add('<definedNames>')
        namesOpen = true
        name = payload.wide()
        scope = 0
        ref = ''
        break
      case 585: {
        const decoded = definition(payload.section(payload.u32()), sheets)
        ref = decoded.ref
        if (decoded.invalidReference) invalidDefinition(name)
        break
      }
      case 586:
        payload.u16()
        scope = payload.u32()
        payload.u8()
        break
      case 587:
        xml.add(`<definedName name="${escapeAttr(name)}"${ref ? ` refersTo="${escapeAttr(ref)}"` : ''}${scope ? ` sheetId="${scope - 1}"` : ''}/>`)
        break
      case 363: {
        if (namesOpen) xml.add('</definedNames>')
        namesOpen = false
        if (!tableOpen) xml.add('<sheetDataSet>')
        tableOpen = true
        const sheet = payload.u32()
        const flags = payload.u8()
        if (sheet >= sheets.length) payload.fail(`external sheet index ${sheet} out of bounds`)
        xml.add(`<sheetData sheetId="${sheet}"${flags & 1 ? ' refreshError="1"' : ''}>`)
        row = -1
        break
      }
      case 364: xml.add('</sheetData>'); break
      case 366: row = payload.u32(); break
      case 472: vm = payload.u32(); break
      case 367: case 368: case 369: case 370: case 371:
        xml.add(cell(type, payload, row, vm))
        vm = 0
        break
      default: continue
    }
    payload.done()
  }
  if (namesOpen) xml.add('</definedNames>')
  if (tableOpen) xml.add('</sheetDataSet>')
  if (book) xml.add('</externalBook>')
  xml.add('</externalLink>')
}

function sheetNames(payload: Payload): string[] {
  const count = payload.u32()
  if (count * 4 > payload.remaining) payload.fail(`invalid external sheet count ${count}`)
  const names: string[] = []
  for (let i = 0; i < count; i++) names.push(payload.wide())
  return names
}

function cell(type: number, payload: Payload, row: number, vm: number): string {
  const ref = address(payload, row, payload.u32())
  const metadata = vm ? ` vm="${vm}"` : ''
  if (type === 367) return `<cell r="${ref}"${metadata}/>`
  const kind = type === 369 ? 'b' : type === 370 ? 'e' : type === 371 ? 'str' : 'n'
  const value = type === 368 ? payload.f64() : type === 369 ? payload.u8()
    : type === 370 ? errorText(payload.u8()) : escapeText(payload.nullableWide() ?? '')
  return `<cell r="${ref}"${metadata} t="${kind}"><v>${value}</v></cell>`
}

function definition(payload: Payload, sheets: readonly string[]): { ref: string; invalidReference?: true } {
  if (!payload.remaining) return { ref: '' }
  const type = payload.u8()
  if (type === 0x1c) {
    const ref = errorText(payload.u8())
    payload.done()
    return { ref }
  }
  if (type === 0x3c || type === 0x3d) {
    // ExtPtgRefErr3D and ExtPtgAreaErr3D leave the coordinate bytes undefined.
    payload.skip(type === 0x3c ? 8 : 12)
    payload.done()
    return { ref: '#REF!', invalidReference: true }
  }
  const first = payload.i16()
  const last = payload.i16()
  const start = sheets[first]
  const end = sheets[last]
  const prefix = start === undefined || end === undefined ? '#REF!' : `'${(first === last ? start : `${start}:${end}`).replace(/'/g, "''")}'!`
  const cell = (row: number, col: number) => {
    const ref = address(payload, row, col & 0x3fff)
    const digits = String(row + 1)
    return (col & 0x4000 ? '' : '$') + ref.slice(0, -digits.length) + (col & 0x8000 ? '' : '$') + digits
  }
  if (type === 0x3a) {
    const ref = '=' + prefix + cell(payload.u16(), payload.u16())
    payload.done()
    return { ref }
  }
  if (type === 0x3b) {
    const row1 = payload.u16()
    const row2 = payload.u16()
    const ref = '=' + prefix + cell(row1, payload.u16()) + ':' + cell(row2, payload.u16())
    payload.done()
    return { ref }
  }
  return { ref: '#REF!' }
}
