/**
 * A link to another workbook as Excel writes its part (`xl/externalLinks/externalLinkN.xml`): the book's sheets' names,
 * its names and the values stored of its sheets' cells, a number without a type, text as `t="str"`, a boolean as
 * `t="b"`, an error as `t="e"` (#SPILL! and #CALC! as `#VALUE!`: Excel repairs a file that holds their text there, and
 * shows and saves `#VALUE!` for them with a rich value, the set `links/written`), a cell stored empty as `<cell r="A4"/>`,
 * the mark of a failed update as `refreshError="1"`. A sheet without `cells` gets no data, as every cell of it is `#REF!` then, but empty data when its book is held open (`open`). The
 * part points at the other book through its relationship `rId1` (`book.ts` writes it with the target `linkTarget` gives).
 */
import type { ErrorCode, ExternalBook, ExternalSheet, ExternalValue } from '../../contract.js'
import { parseAddress, sortAddresses } from '../../model.js'
import type { Issues } from '../issues.js'
import { escapeAttr, escapeText } from '../xml/text.js'
import { errorXml } from './sheet.js'
import { HEAD, MAIN, REL } from './xml.js'

function cellXml(addr: string, v: ExternalValue | null, error: (code: ErrorCode) => { t: string; v: string }): string {
  if (v === null) return `<cell r="${addr}"/>`
  const x = typeof v === 'object' ? error(v.error) : typeof v === 'boolean' ? { t: ' t="b"', v: v ? '1' : '0' } : typeof v === 'string' ? { t: ' t="str"', v: escapeText(v) } : { t: '', v: String(v) }
  return `<cell r="${addr}"${x.t}><v>${x.v}</v></cell>`
}

/**
 * A sheet's stored values by row, `sheetId` its index; none for a sheet without `cells`, but an empty one of a book held
 * open, which Excel reads as an empty sheet. `at` names its cells in issues (`[1]Sheet1!`).
 */
function sheetData(s: ExternalSheet, i: number, at: string, issues: Issues, open?: true): string {
  if (!s.cells && !open) return ''
  const rows: string[] = []
  let row = -1
  for (const a of sortAddresses(Object.keys(s.cells ?? {}))) {
    const r = parseAddress(a).row
    if (r !== row) rows.push(`${row < 0 ? '' : '</row>'}<row r="${(row = r) + 1}">`)
    rows.push(cellXml(a, s.cells![a], (code) => errorXml(code, (key) => issues.add('errorValue', { key }, { cell: `${at}${a}` }))))
  }
  const mark = s.refreshError ? ' refreshError="1"' : ''
  return rows.length ? `<sheetData sheetId="${i}"${mark}>${rows.join('')}</row></sheetData>` : `<sheetData sheetId="${i}"${mark}/>`
}

/** A link's target as Excel writes it in the relationship: a space as `%20`, the rest as it is (Cyrillic, backslashes, `%`, an escape already made). */
export const linkTarget = (b: ExternalBook): string => b.target.replaceAll(' ', '%20')

/** The part of the link `[n]`; each list is left out when empty, as the file's schema has no empty one. */
export function externalLinkXml(b: ExternalBook, n: number, issues: Issues): string {
  const sheets = b.sheets.map((s) => `<sheetName val="${escapeAttr(s.name)}"/>`).join('')
  const names = (b.names ?? []).map((x) => `<definedName name="${escapeAttr(x.name)}" refersTo="=${escapeAttr(x.ref)}"${x.sheet === undefined ? '' : ` sheetId="${x.sheet}"`}/>`).join('')
  const data = b.sheets.map((s, i) => sheetData(s, i, `[${n}]${s.name}!`, issues, b.open)).join('')
  const body = (sheets ? `<sheetNames>${sheets}</sheetNames>` : '') + (names ? `<definedNames>${names}</definedNames>` : '') + (data ? `<sheetDataSet>${data}</sheetDataSet>` : '')
  const book = `<externalBook xmlns:r="${REL}" r:id="rId1"${body ? `>${body}</externalBook>` : '/>'}`
  return `${HEAD}<externalLink xmlns="${MAIN}">${book}</externalLink>`
}
