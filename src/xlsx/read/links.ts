/**
 * The book's links to other workbooks (`xl/externalLinks/externalLinkN.xml`), in the order the workbook lists them
 * (`<externalReferences>`), as `[1]` names the first: where the other book is (the target of the link's own
 * relationship, whatever its kind, `xlPathMissing` too; Excel's alternate addresses `xxl21:alternateUrls` do not change
 * it), its sheets' names, its defined names and the values Excel stored of its cells. A link of another kind (DDE, OLE)
 * or one whose part is missing keeps its number as a book without sheets and with an empty target, reported.
 */
import type { ExternalBook, ExternalName, ExternalSheet, ExternalValue } from '../../contract.js'
import type { Issues } from '../issues.js'
import { attr, attributes, element, elements, flag, ownEntry } from '../xml/scan.js'
import { readText } from '../xml/text.js'
import type { Archive } from '../zip/read.js'
import { ERRORS } from './cells.js'
import { relsOf, type Rel } from './package.js'
import { fileCell } from './refs.js'
import { stringItem } from './strings.js'

export interface ReadLinks {
  books: ExternalBook[]
  /** The numbers `[N]` of the links that are workbooks, which formulas may refer to. */
  known: Set<number>
}

/** The place of a link the model has no workbook for. */
const placeholder = (): ExternalBook => ({ target: '', sheets: [] })

/**
 * A stored value of another book's cell, null for a cell stored empty; an error the model has no code for is reported, and
 * so is a rich value (`vm`): Excel shows the value the cell stores, `#VALUE!` for a rich value of #SPILL! too (the set
 * `links/written`).
 */
function cachedValue(e: { attrs: string; body: string | null }, where: string, issues: Issues): ExternalValue | null {
  const a = attributes(e.attrs)
  if (a.vm !== undefined) issues.add('cellMetadata', { key: 'valueMetadata' }, { cell: where })
  const body = e.body ?? ''
  const t = a.t ?? 'n'
  if (t === 'inlineStr' && element(body, 'is')) return stringItem(element(body, 'is')!.body ?? '').text
  const v = element(body, 'v')
  if (!v) return null
  const raw = v.body ?? ''
  // text is text even when empty, as a cell's `<v/>` of an empty text result
  if (t === 'str' || t === 's' || t === 'inlineStr') return readText(raw)
  if (raw === '') return null
  if (t === 'b') return flag(raw)
  if (t === 'e') {
    const code = ownEntry(ERRORS, readText(raw).trim())
    if (!code) issues.add('errorValue', { key: 'errorUnknown' }, { cell: where })
    return { error: code ?? 'ERROR' }
  }
  const n = Number(raw)
  return Number.isFinite(n) ? n : null
}

/** The stored values of one sheet's data (`<sheetData>`), by address. */
function sheetCells(body: string, at: string, issues: Issues): Record<string, ExternalValue | null> {
  const out: Record<string, ExternalValue | null> = {}
  for (const c of elements(body, 'cell')) {
    const addr = fileCell(attributes(c.attrs).r ?? '')
    if (addr) out[addr] = cachedValue(c, `${at}${addr}`, issues)
  }
  return out
}

/** A link's `externalBook`: its sheets with their data, and its names; `target` from the part's relationships. */
function externalBook(xml: string, rels: Rel[], n: number, issues: Issues): ExternalBook {
  const root = element(xml, 'externalBook')!
  const id = attr(attributes(root.attrs), 'id')
  const body = root.body ?? ''
  const sheets: ExternalSheet[] = elements(element(body, 'sheetNames')?.body ?? '', 'sheetName').map((e) => ({ name: readText(attributes(e.attrs).val ?? '') }))
  for (const d of elements(element(body, 'sheetDataSet')?.body ?? '', 'sheetData')) {
    const a = attributes(d.attrs)
    const sheet = sheets[Number(a.sheetId)]
    if (!sheet) continue
    sheet.cells = sheetCells(d.body ?? '', `[${n}]${sheet.name}!`, issues)
    if (flag(a.refreshError)) sheet.refreshError = true
  }
  const names: ExternalName[] = []
  for (const e of elements(element(body, 'definedNames')?.body ?? '', 'definedName')) {
    const a = attributes(e.attrs)
    const name: ExternalName = { name: readText(a.name ?? ''), ref: readText(a.refersTo ?? '').replace(/^=/, '') }
    if (a.sheetId !== undefined) {
      const k = Number(a.sheetId)
      if (!Number.isInteger(k) || !sheets[k]) {
        issues.add('name', { key: 'nameOfSheetLeftOut' }, { name: `[${n}]${name.name}` })
        continue
      }
      name.sheet = k
    }
    names.push(name)
  }
  const book: ExternalBook = { target: rels.find((r) => r.id === id)?.target ?? '', sheets }
  if (names.length) book.names = names
  return book
}

/** The links of the workbook `wb`, whose relationships are `rels`. */
export async function readLinks(archive: Archive, wb: string, rels: Rel[], issues: Issues): Promise<ReadLinks> {
  const out: ReadLinks = { books: [], known: new Set() }
  const byId = new Map(rels.map((r) => [r.id, r]))
  for (const e of elements(element(wb, 'externalReferences')?.body ?? '', 'externalReference')) {
    const n = out.books.length + 1
    const r = byId.get(attr(attributes(e.attrs), 'id') ?? '')
    const xml = r && !r.external ? await archive.text(r.target) : undefined
    if (xml === undefined || !element(xml, 'externalBook')) {
      issues.add('externalLink', { key: 'externalLinks' })
      out.books.push(placeholder())
      continue
    }
    out.books.push(externalBook(xml, await relsOf(archive, r!.target), n, issues))
    out.known.add(n)
  }
  return out
}
