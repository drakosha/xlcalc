/**
 * The cell formats of the `keep` step: a file Excel makes with formats the workbook does not tell apart (`MADE_FORMATS`:
 * a quote prefix, a theme colour beside its colour, a cell style of its own, a pattern fill, rows and columns of each
 * colour; `excel-keep.applescript`), and operations of the structure over it (`FORMATS`) that Excel does
 * (`excel-keep-struct.applescript`) and `writeKept` writes. What the files give each cell, row and column is compared
 * by the content of its cell format (`formatsOf`), not by its number: Excel may number them otherwise when it saves.
 */
import { createHash } from 'node:crypto'
import { openSource } from '../../../src/xlsx/keep/source.js'
import { partText } from '../../../src/xlsx/read/package.js'
import { attributes, element, elements } from '../../../src/xlsx/xml/scan.js'
import type { Made, Struct } from './books.js'

export const MADE_FORMATS: Made[] = [{ id: 'formats', features: ['formats'] }]

const rows = (id: string, row: number, count: number, op: 'insertRows' | 'deleteRows'): Struct => ({ id: `formats-${id}`, made: 'formats', op: { op, sheet: 'Data', row, count }, excel: [op, 'Data', `A${row}:A${row + count - 1}`, '', ''] })
const columns = (id: string, column: string, count: number, op: 'insertColumns' | 'deleteColumns'): Struct => {
  const last = String.fromCharCode(column.charCodeAt(0) + count - 1)
  return { id: `formats-${id}`, made: 'formats', op: { op, sheet: 'Data', column, count }, excel: [op, 'Data', `${column}1:${last}1`, '', ''] }
}

/**
 * The operations over the made file: rows inserted above every block and inside the quote prefix's, rows deleted inside
 * it, columns inserted and deleted before the blocks, a row inserted at the row of a colour and one below it, a column
 * inserted after the column of a colour.
 */
export const FORMATS: Struct[] = [
  rows('insert-above', 2, 2, 'insertRows'),
  rows('insert-in', 8, 1, 'insertRows'),
  rows('delete-in', 7, 2, 'deleteRows'),
  columns('insert-column', 'B', 1, 'insertColumns'),
  columns('delete-column', 'A', 1, 'deleteColumns'),
  rows('insert-at-row', 25, 1, 'insertRows'),
  rows('insert-below-row', 26, 1, 'insertRows'),
  columns('insert-after-column', 'I', 1, 'insertColumns'),
  // held out after the cases above were measured: rows and a column inserted inside the blocks, rows deleted across two,
  // a row inserted below a block, the column of a colour deleted
  rows('insert-in-colour', 13, 1, 'insertRows'),
  columns('insert-in-blocks', 'C', 1, 'insertColumns'),
  rows('delete-across', 11, 2, 'deleteRows'),
  rows('insert-below-block', 14, 1, 'insertRows'),
  columns('delete-colour-column', 'H', 1, 'deleteColumns'),
]

/** A cell format's content: its number format, font, fill and border as written, its cell style by name, its other attributes and its children. */
function contents(styles: string): string[] {
  const list = (local: string, item: string) => elements(element(styles, local)?.body ?? '', item)
  const xml = (e: { attrs: string; body: string | null } | undefined, local: string) => (e ? `<${local}${e.attrs}${e.body === null ? '/>' : `>${e.body}</${local}>`}` : '-')
  const [fonts, fills, borders] = [list('fonts', 'font'), list('fills', 'fill'), list('borders', 'border')]
  const codes = new Map(list('numFmts', 'numFmt').map((e) => [attributes(e.attrs).numFmtId, attributes(e.attrs).formatCode]))
  const named = new Map(list('cellStyles', 'cellStyle').map((e) => [attributes(e.attrs).xfId, attributes(e.attrs).name]))
  return list('cellXfs', 'xf').map((e) => {
    const { numFmtId = '0', fontId = '0', fillId = '0', borderId = '0', xfId = '0', ...rest } = attributes(e.attrs)
    const other = Object.entries(rest).sort(([a], [b]) => (a < b ? -1 : 1)).map(([k, v]) => `${k}=${v}`)
    return [`numFmt ${codes.get(numFmtId) ?? numFmtId}`, xml(fonts[Number(fontId)], 'font'), xml(fills[Number(fillId)], 'fill'), xml(borders[Number(borderId)], 'border'), `style ${named.get(xfId) ?? xfId}`, ...other, e.body ?? ''].join(' ')
  })
}

/** Columns past which a `<col>` is not counted column by column. */
const COLUMNS = 30

/**
 * What a file gives the cells, rows and columns of its sheet Data that are not in its first cell format: by address
 * (`B7`, `row 23`, `column 7`), a hash of the cell format's content, and the contents by hash.
 */
export async function formatsOf(data: Uint8Array): Promise<{ at: Record<string, string>; formats: Record<string, string> }> {
  const src = await openSource(data)
  const all = contents((await partText(src.archive, src.wbRels, 'styles')) ?? '')
  const sheet = src.sheets.find((s) => src.decls[s.decl].name === 'Data')!.xml
  const at: Record<string, string> = {}
  const formats: Record<string, string> = {}
  const put = (where: string, s: string | undefined) => {
    const c = all[Number(s ?? 0)] ?? `missing ${s}`
    if (c === all[0]) return
    const h = createHash('sha256').update(c).digest('hex').slice(0, 10)
    formats[h] = c
    at[where] = h
  }
  for (const m of sheet.matchAll(/<c r="([A-Z]+\d+)"([^>]*)>/g)) put(m[1], attributes(m[2]).s)
  for (const m of sheet.matchAll(/<row ([^>]*)>/g)) {
    const a = attributes(m[1])
    if (a.customFormat === '1') put(`row ${a.r}`, a.s)
  }
  for (const m of sheet.matchAll(/<col ([^>]*)>/g)) {
    const a = attributes(m[1])
    for (let c = Number(a.min); c <= Math.min(Number(a.max), COLUMNS); c++) put(`column ${c}`, a.style)
  }
  return { at, formats }
}
