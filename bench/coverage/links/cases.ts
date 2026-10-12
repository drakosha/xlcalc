/**
 * Links to other workbooks: each case is a book whose formulas refer to other books (`[1]Sheet1!A1`), with what Excel
 * stores of those books in the file itself (`xl/externalLinks/externalLinkN.xml`: their sheets' names, the values of
 * the cells the formulas read, their defined names). `writeXlsx` writes the book's own cells, then the formulas with
 * links, the link parts and this book's names over them are put in by hand (`linkFile`), as Excel writes them, so the
 * files do not depend on the writer; a case marked `written` is the file `writeXlsx` writes itself from the case's book
 * with its links (`writtenBook`), to check the writer. `npx tsx bench/coverage/run.ts links <set>` has Excel open each
 * file without updating its links, as a person who answers "Don't Update" does, and read what the cells show.
 */
import type { BookInput, CellValue, ExternalBook, RangeRef, SheetInput } from '../../../src/contract.js'
import { fromFile } from '../../../src/xlsx/formulas/from-file.js'
import { ERROR_TEXT } from '../../../src/syntax/errors.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { escapeAttr, escapeText } from '../../../src/xlsx/xml/text.js'
import { openArchive } from '../../../src/xlsx/zip/read.js'
import { zip } from '../../../src/xlsx/zip/write.js'
import { cellsOf, MIRRORS, mirrorFormulas, mirrorRanges, type RectsAnswer } from '../case-book.js'

/** A value Excel stores of another book's cell: a number, a text, a boolean or an error. */
export type CachedValue = Exclude<CellValue, null | { date: string }>

/** Another book as a file stores it beside the formulas that read it. */
export interface LinkedBook {
  /** The link's target as the file holds it (`externalLinkN.xml.rels`): a file name, a relative or an absolute path. */
  target: string
  /**
   * Its sheets in order, with the values of the cells Excel stored (null for a cell stored empty, `<cell r="A4"/>`);
   * `refreshError` marks the sheet's data as Excel marks a failed update, `data: false` leaves its data out (a sheet the
   * link names but holds no values of).
   */
  sheets: { name: string; cells?: Record<string, CachedValue | null>; refreshError?: true; data?: false }[]
  /** Its defined names: `ref` as the file holds it (`=Sheet1!$B$1`), `sheet` the index of a sheet's own name. */
  names?: { name: string; ref: string; sheet?: number }[]
}

export interface LinksCase {
  id: string
  about: string
  /** This book's sheets, with the cells the writer holds. */
  sheets: SheetInput[]
  /** Formulas with links, as the file holds them (`[1]Sheet1!A1+1`), by sheet and cell; with `array` a dynamic array formula over that range. */
  linked: Record<string, Record<string, LinkedFormula>>
  /** This book's defined names whose definitions the writer does not hold, as the file holds them (`[1]Sheet1!$A$1`). */
  names?: { name: string; ref: string; sheet?: string }[]
  /** The other books, the first answering `[1]`. */
  links: LinkedBook[]
  /** The rectangles read. */
  read: RangeRef[]
  /** Operations Excel takes after the first read, before the full calculation and the second read. */
  ops?: LinksOp[]
  /** The file is the one `writeXlsx` writes of the case's book with its links (`writtenBook`), not one put together by hand. */
  written?: true
  /**
   * The values predicted before Excel ran the case, by `Sheet!A1`, for the second read: the rule over the stored values,
   * and where the rule was not known the other answers thought possible (`alternatives`).
   */
  predicted?: Record<string, CachedValue | null>
  alternatives?: Record<string, (CachedValue | null)[]>
}

export type LinkedFormula = string | { formula: string; array: string }

/** An operation on this book: a fill down of `to`, a copy of `range` to `to`, rows inserted over `range`, the sheet renamed `to`. */
export type LinksOp = { op: 'fillDown'; sheet: string; to: string } | { op: 'copyRange'; sheet: string; range: string; to: string } | { op: 'insertRows'; sheet: string; range: string } | { op: 'renameSheet'; sheet: string; to: string }

export interface LinksCases {
  cases: LinksCase[]
}

export interface LinksSet {
  buildLinksCases(): LinksCases
}

/** One read of the rectangles: values through the mirror sheets, formulas, formulas in the user's language and the shown text. */
export interface LinksRead extends RectsAnswer {
  local: Record<string, string>
  shown: Record<string, string>
}

export interface LinksAnswer {
  /** The linked books as Excel lists them (`LinkSources`). */
  sources: string[]
  /** The folder of the opened book as Excel names it: a closed book's link is shown with its path from there. */
  folder: string
  /** The operations Excel refused, by their index: its error. */
  errors?: Record<number, string>
  /** After opening, then after the operations and a full calculation. */
  reads: LinksRead[]
}

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const PACKAGE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


/** A cached cell as `externalLinkN.xml` holds it. */
function cachedCell(addr: string, v: CachedValue | null): string {
  if (v === null) return `<cell r="${addr}"/>`
  if (typeof v === 'number') return `<cell r="${addr}"><v>${v}</v></cell>`
  if (typeof v === 'boolean') return `<cell r="${addr}" t="b"><v>${v ? 1 : 0}</v></cell>`
  if (typeof v === 'string') return `<cell r="${addr}" t="str"><v>${escapeText(v)}</v></cell>`
  return `<cell r="${addr}" t="e"><v>${ERROR_TEXT[v.error]}</v></cell>`
}


/** A sheet's cached cells by row, as Excel writes them. */
function sheetData(i: number, s: LinkedBook['sheets'][number]): string {
  if (s.data === false) return ''
  const cells = s.cells ?? {}
  const rows = new Map<number, string[]>()
  const keys = Object.keys(cells).sort((a, b) => rowOf(a) - rowOf(b) || colOf(a) - colOf(b))
  for (const k of keys) rows.set(rowOf(k), [...(rows.get(rowOf(k)) ?? []), cachedCell(k, cells[k])])
  return `<sheetData sheetId="${i}"${s.refreshError ? ' refreshError="1"' : ''}>${[...rows].map(([r, cs]) => `<row r="${r}">${cs.join('')}</row>`).join('')}</sheetData>`
}

const rowOf = (a: string) => Number(a.replace(/^[A-Z]+/, ''))
const colOf = (a: string) => [...a.replace(/\d+$/, '')].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0)

/** `externalLinkN.xml` of a linked book. */
export function externalLinkXml(b: LinkedBook): string {
  const names = b.names?.length ? `<definedNames>${b.names.map((n) => `<definedName name="${escapeAttr(n.name)}" refersTo="${escapeAttr(n.ref)}"${n.sheet === undefined ? '' : ` sheetId="${n.sheet}"`}/>`).join('')}</definedNames>` : ''
  const sheets = `<sheetNames>${b.sheets.map((s) => `<sheetName val="${escapeAttr(s.name)}"/>`).join('')}</sheetNames>`
  return `${XML_HEAD}<externalLink xmlns="${MAIN}"><externalBook xmlns:r="${REL}" r:id="rId1">${sheets}${names}<sheetDataSet>${b.sheets.map((s, i) => sheetData(i, s)).join('')}</sheetDataSet></externalBook></externalLink>`
}

const linkRels = (b: LinkedBook) => `${XML_HEAD}<Relationships xmlns="${PACKAGE_REL}"><Relationship Id="rId1" Type="${REL}/externalLinkPath" Target="${escapeAttr(b.target)}" TargetMode="External"/></Relationships>`

/** A formula cell of the sheet's part put in by hand over what the writer wrote there (the cell's value). */
function withFormula(xml: string, addr: string, f: LinkedFormula): string {
  const cell = new RegExp(`<c r="${addr}"(?: [^>]*?)?(?:/>|>.*?</c>)`)
  const put = typeof f === 'string' ? `<c r="${addr}"><f>${escapeText(f)}</f></c>` : `<c r="${addr}" cm="1"><f t="array" ref="${f.array}">${escapeText(f.formula)}</f></c>`
  if (cell.test(xml)) return xml.replace(cell, put)
  throw new Error(`linkFile: no cell ${addr} written (give it a value in the case's sheet)`)
}

/** The book of a case as `writeXlsx` writes it, with mirror sheets over the rectangles read. */
function caseSheets(c: LinksCase): SheetInput[] {
  const at = mirrorRanges(c.read)
  const mirrors: SheetInput[] = MIRRORS.map((name) => ({ name, cells: {} }))
  c.read.forEach((g, i) => {
    const dst = cellsOf(at[i]).addrs
    cellsOf(g.range).addrs.forEach((a, k) => mirrorFormulas({ sheet: g.sheet, range: a }).forEach((f, m) => (mirrors[m].cells[dst[k]] = { formula: f.slice(1) })))
  })
  // a cell the case links gets a placeholder value, so the writer writes the cell the formula goes into
  const own = c.sheets.map((s) => ({ ...s, cells: { ...Object.fromEntries(Object.keys(c.linked[s.name] ?? {}).map((a) => [a, { value: 0 }])), ...s.cells } }))
  return [...own, ...mirrors]
}

/** The file of a case: the writer's book, then the formulas with links, the link parts and the names put in by hand. */
export async function linkFile(c: LinksCase): Promise<Uint8Array> {
  if (c.written) return (await writeXlsx({ book: writtenBook(c) })).data
  const sheets = caseSheets(c)
  const { data } = await writeXlsx({ book: { sheets } })
  const archive = openArchive(data)
  const files: { name: string; data: Uint8Array }[] = []
  const enc = new TextEncoder()
  for (const name of archive.names) {
    let text = name.endsWith('.xml') || name.endsWith('.rels') ? (await archive.text(name))! : undefined
    const sheet = /^xl\/worksheets\/sheet(\d+)\.xml$/.exec(name)
    if (text !== undefined && sheet) {
      for (const [addr, f] of Object.entries(c.linked[sheets[Number(sheet[1]) - 1].name] ?? {})) text = withFormula(text, addr, f)
    } else if (name === 'xl/workbook.xml') text = workbookXml(text!, c, sheets)
    else if (name === 'xl/_rels/workbook.xml.rels') text = text!.replace('</Relationships>', `${c.links.map((_, i) => `<Relationship Id="rIdL${i + 1}" Type="${REL}/externalLink" Target="externalLinks/externalLink${i + 1}.xml"/>`).join('')}</Relationships>`)
    else if (name === '[Content_Types].xml') text = text!.replace('</Types>', `${c.links.map((_, i) => `<Override PartName="/xl/externalLinks/externalLink${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>`).join('')}</Types>`)
    files.push({ name, data: text === undefined ? (await archive.bytes(name))! : enc.encode(text) })
  }
  c.links.forEach((b, i) => {
    files.push({ name: `xl/externalLinks/externalLink${i + 1}.xml`, data: enc.encode(externalLinkXml(b)) })
    files.push({ name: `xl/externalLinks/_rels/externalLink${i + 1}.xml.rels`, data: enc.encode(linkRels(b)) })
  })
  return zip(files)
}

/** workbook.xml with the external references after the sheets and the case's names among the defined names. */
function workbookXml(xml: string, c: LinksCase, sheets: SheetInput[]): string {
  const refs = `<externalReferences>${c.links.map((_, i) => `<externalReference r:id="rIdL${i + 1}"/>`).join('')}</externalReferences>`
  let out = xml.replace('</sheets>', `</sheets>${refs}`)
  if (!c.names?.length) return out
  const defs = c.names.map((n) => `<definedName name="${escapeAttr(n.name)}"${n.sheet === undefined ? '' : ` localSheetId="${sheets.findIndex((s) => s.name === n.sheet)}"`}>${escapeText(n.ref)}</definedName>`).join('')
  out = out.includes('</definedNames>') ? out.replace('</definedNames>', `${defs}</definedNames>`) : out.replace(refs, `${refs}<definedNames>${defs}</definedNames>`)
  return out
}

/** A case's book as the engine takes it: its cells, its formulas with links and names in the engine's spelling, the stored values of its links. */
export function linksBook(c: LinksCase): BookInput {
  const formula = (f: LinkedFormula) => fromFile(typeof f === 'string' ? f : f.formula).formula
  const sheets = c.sheets.map((s) => ({ ...s, cells: { ...s.cells, ...Object.fromEntries(Object.entries(c.linked[s.name] ?? {}).map(([a, f]) => [a, { formula: formula(f) }])) } }))
  const externalBooks: ExternalBook[] = c.links.map((b) => ({
    target: b.target,
    sheets: b.sheets.map((s) => ({ name: s.name, ...(s.data === false ? {} : { cells: s.cells ?? {} }), ...(s.refreshError ? { refreshError: true as const } : {}) })),
    ...(b.names?.length ? { names: b.names.map((n) => ({ name: n.name, ref: n.ref.replace(/^=/, ''), ...(n.sheet === undefined ? {} : { sheet: n.sheet }) })) } : {}),
  }))
  return { sheets, ...(c.names?.length ? { names: c.names.map((n) => ({ name: n.name, ref: fromFile(n.ref).formula, ...(n.sheet ? { scope: n.sheet } : {}) })) } : {}), externalBooks }
}

/** The book `writeXlsx` writes of a case marked `written`: the case's book with its links (`linksBook`) and the mirror sheets. */
export function writtenBook(c: LinksCase): BookInput {
  const book = linksBook(c)
  return { ...book, sheets: [...book.sheets, ...caseSheets(c).slice(c.sheets.length)] }
}
