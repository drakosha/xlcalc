/**
 * A link to another workbook as Excel writes its part (`xl/externalLinks/externalLinkN.xml`): the book's sheets' names,
 * its names and the values stored of its sheets' cells, a number without a type, text as `t="str"`, a boolean as
 * `t="b"`, an error as `t="e"`, a cell stored empty as `<cell r="A4"/>`, the mark of a failed update as
 * `refreshError="1"`. A sheet without `cells` gets no data, as every cell of it is `#REF!` then. The part points at the
 * other book through its relationship `rId1` (`book.ts` writes it with the target as the book holds it).
 */
import type { ExternalBook, ExternalSheet, ExternalValue } from '../../contract.js'
import { parseAddress, sortAddresses } from '../../model.js'
import { ERROR_TEXT } from '../../syntax/errors.js'
import { escapeAttr, escapeText } from '../xml/text.js'
import { HEAD, MAIN } from './xml.js'

function cellXml(addr: string, v: ExternalValue | null): string {
  if (v === null) return `<cell r="${addr}"/>`
  if (typeof v === 'number') return `<cell r="${addr}"><v>${v}</v></cell>`
  if (typeof v === 'boolean') return `<cell r="${addr}" t="b"><v>${v ? 1 : 0}</v></cell>`
  if (typeof v === 'string') return `<cell r="${addr}" t="str"><v>${escapeText(v)}</v></cell>`
  return `<cell r="${addr}" t="e"><v>${ERROR_TEXT[v.error]}</v></cell>`
}

/** A sheet's stored values by row, `sheetId` its index; none for a sheet without `cells`. */
function sheetData(s: ExternalSheet, i: number): string {
  if (!s.cells) return ''
  const rows: string[] = []
  let row = -1
  for (const a of sortAddresses(Object.keys(s.cells))) {
    const r = parseAddress(a).row
    if (r !== row) rows.push(`${row < 0 ? '' : '</row>'}<row r="${(row = r) + 1}">`)
    rows.push(cellXml(a, s.cells[a]))
  }
  const mark = s.refreshError ? ' refreshError="1"' : ''
  return rows.length ? `<sheetData sheetId="${i}"${mark}>${rows.join('')}</row></sheetData>` : `<sheetData sheetId="${i}"${mark}/>`
}

/** The part of a link; each list is left out when empty, as the file's schema has no empty one. */
export function externalLinkXml(b: ExternalBook): string {
  const sheets = b.sheets.map((s) => `<sheetName val="${escapeAttr(s.name)}"/>`).join('')
  const names = (b.names ?? []).map((n) => `<definedName name="${escapeAttr(n.name)}" refersTo="=${escapeAttr(n.ref)}"${n.sheet === undefined ? '' : ` sheetId="${n.sheet}"`}/>`).join('')
  const data = b.sheets.map(sheetData).join('')
  const body = (sheets ? `<sheetNames>${sheets}</sheetNames>` : '') + (names ? `<definedNames>${names}</definedNames>` : '') + (data ? `<sheetDataSet>${data}</sheetDataSet>` : '')
  const book = `<externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"${body ? `>${body}</externalBook>` : '/>'}`
  return `${HEAD}<externalLink xmlns="${MAIN}">${book}</externalLink>`
}
