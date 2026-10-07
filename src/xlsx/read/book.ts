/**
 * `readXlsx`: an xlsx file as a book. The workbook part is found through the package's relationships, its
 * sheets taken in the order the workbook lists them (not by part names), each linked to its part by id; shared
 * strings, styles, the theme and the cell metadata of dynamic arrays come from the workbook's relationships.
 * Everything the book cannot hold is reported in `issues`, nothing is dropped silently.
 */
import type { BookInput, NameInput, SheetInput } from '../../contract.js'
import { checkBook, InputError } from '../../input.js'
import { LOCALES } from '../../locale/registry.js'
import { Issues } from '../issues.js'
import { LOCALE_PROPERTY } from '../property.js'
import { palette } from '../styles/color.js'
import { readStyles, type Xf } from '../styles/read.js'
import type { XlsxRead } from '../types.js'
import { attr, attributes, element, elements, flag, has } from '../xml/scan.js'
import { decodeXml, readText } from '../xml/text.js'
import { openArchive, type Archive } from '../zip/read.js'
import { readFormula } from './formulas.js'
import { checkIntersections } from './intersect.js'
import { cellMetadata } from './metadata.js'
import { relsOf, type Rel } from './package.js'
import { readSheet } from './sheet.js'
import { readStrings } from './strings.js'


/** Days between the 1904 and the 1900 date systems' day 0. */
const DAYS_1904 = 1462

/**
 * A number format that shows a date (a time alone, or a duration, is the same in both date systems): a year, a
 * day or month with a separator, or a month's or day's name. A unit such as `0 m` is no date.
 */
function hasDate(code: string): boolean {
  const s = code.replace(/"[^"]*"|\\.|\[[^\]]*\]/g, '')
  return /y/i.test(s) || /(?:d{1,4}|m{1,5})\s*[/.\- ]\s*(?:d{1,4}|m{1,5}|y{1,4})/i.test(s) || /m{3,}|d{3,}/i.test(s)
}

/** Parts the workbook links to that the model has no place for, by relationship kind. */
const LINKED: Record<string, ['macros' | 'pivotTable' | 'connection' | 'externalLink', string]> = {
  vbaProject: ['macros', 'macros (a VBA project) are not in the model'],
  pivotCacheDefinition: ['pivotTable', 'pivot tables are not in the model: their cells are read as stored'],
  connections: ['connection', 'data connections and queries are not in the model'],
  externalLink: ['externalLink', 'links to other workbooks are not in the model'],
}

async function locale(archive: Archive, rels: Rel[], issues: Issues): Promise<string | undefined> {
  const part = rels.find((r) => r.kind === 'custom-properties' && !r.external)
  const xml = part && (await archive.text(part.target))
  const p = xml && elements(xml, 'property').find((e) => decodeXml(attributes(e.attrs).name ?? '') === LOCALE_PROPERTY)
  if (!p) return undefined
  const tag = readText(element(p.body ?? '', 'lpwstr')?.body ?? '')
  if (LOCALES.has(tag)) return tag
  issues.add('locale', `the locale ${JSON.stringify(tag)} is not one of the package's, left out`)
  return undefined
}

interface SheetDecl {
  name: string
  kind?: string
  part?: string
  hidden: boolean
}

function readNames(xml: string, decls: SheetDecl[], read: Set<string>, issues: Issues): NameInput[] {
  const out: NameInput[] = []
  for (const e of elements(element(xml, 'definedNames')?.body ?? '', 'definedName')) {
    const a = attributes(e.attrs)
    const name = readText(a.name ?? '')
    // Excel's placeholders for functions newer than the file's readers: not names of the book
    if (/^_xlfn\./i.test(name)) continue
    if (/^_xlnm\./i.test(name)) {
      issues.add('name', 'a built-in name (a print area or titles, the range of a filter) is not in the model', { name })
      continue
    }
    const scope = a.localSheetId === undefined ? undefined : decls[Number(a.localSheetId)]?.name
    if (a.localSheetId !== undefined && (scope === undefined || !read.has(scope))) {
      issues.add('name', 'a name of a sheet left out is left out', { name })
      continue
    }
    const f = readFormula(readText(e.body ?? ''), issues, { name })
    const ok = f.ast !== undefined && ((f.ast.k === 'ref' && f.ast.sheet !== null) || (f.ast.k === 'err' && f.ast.code === 'REF'))
    if (!ok) issues.add('name', 'the engine takes only a reference in a name: a constant or a formula is kept as written, and formulas using the name compute to an error', { name })
    if (flag(a.hidden)) issues.add('name', 'a hidden name is read as a visible one', { name })
    out.push(scope === undefined ? { name, ref: f.formula } : { name, ref: f.formula, scope })
  }
  return out
}

function workbookIssues(xml: string, rels: Rel[], issues: Issues) {
  if (flag(attributes(element(xml, 'workbookPr')?.attrs ?? '').date1904)) issues.add('date1904', 'the book uses the 1904 date system: numbers in date formats are moved to the 1900 system the engine uses')
  const calc = attributes(element(xml, 'calcPr')?.attrs ?? '')
  if (flag(calc.iterate)) issues.add('calculation', 'iterative calculation is not in the engine: a circular reference computes to #CYCLE!')
  if (calc.calcMode === 'manual') issues.add('calculation', 'the manual calculation mode is not kept')
  if (has(xml, 'workbookProtection')) issues.add('protection', 'workbook protection is not in the model')
  if (has(xml, 'externalReferences')) issues.add('externalLink', 'links to other workbooks are not in the model')
  if (has(xml, 'pivotCaches') || has(xml, 'slicerCaches')) issues.add('pivotTable', 'pivot tables are not in the model: their cells are read as stored')
  for (const r of rels) if (LINKED[r.kind]) issues.add(...LINKED[r.kind])
}

/** Reads an xlsx file (bytes of the file) into a book, the values Excel stored and what the book could not hold. */
export async function readXlsx(data: Uint8Array | ArrayBuffer): Promise<XlsxRead> {
  const archive = openArchive(data instanceof Uint8Array ? data : new Uint8Array(data))
  const issues = new Issues()
  const root = await relsOf(archive, '')
  const wbPath = root.find((r) => r.kind === 'officeDocument' && !r.external)?.target ?? 'xl/workbook.xml'
  const wb = await archive.text(wbPath)
  if (wb === undefined) throw new Error(`xlsx: the file has no workbook part (${wbPath})`)
  const rels = await relsOf(archive, wbPath)
  const part = async (kind: string) => {
    const r = rels.find((x) => x.kind === kind && !x.external)
    return r ? archive.text(r.target) : undefined
  }
  const [stringsXml, stylesXml, themeXml, metaXml, richValues, richStructures] = await Promise.all(
    ['sharedStrings', 'styles', 'theme', 'sheetMetadata', 'rdRichValue', 'rdRichValueStructure'].map(part),
  )
  const table = readStyles(stylesXml, palette(themeXml, stylesXml))
  if (table.font && (table.font.name !== 'Calibri' || table.font.size !== 11))
    issues.add('defaultFont', `the default font ${table.font.name ?? '(no name)'} ${table.font.size ?? '(no size)'} is read as Calibri 11, the model's default`)
  workbookIssues(wb, rels, issues)
  const date1904 = flag(attributes(element(wb, 'workbookPr')?.attrs ?? '').date1904)
  const dateOf = (n: number, xf: Xf | undefined) => (date1904 && xf?.format !== undefined && hasDate(xf.format) ? n + DAYS_1904 : n)
  const byId = new Map(rels.map((r) => [r.id, r]))
  const decls: SheetDecl[] = elements(element(wb, 'sheets')?.body ?? '', 'sheet').map((e) => {
    const a = attributes(e.attrs)
    const r = byId.get(attr(a, 'id') ?? '')
    return { name: readText(a.name ?? ''), kind: r?.kind, part: r?.external ? undefined : r?.target, hidden: a.state === 'hidden' || a.state === 'veryHidden' }
  })
  const strings = readStrings(stringsXml)
  const meta = cellMetadata(metaXml, richValues, richStructures)
  const all = decls.filter((d) => d.kind === 'worksheet').map((d) => d.name)
  const names = readNames(wb, decls, new Set(all), issues)
  const sheets: SheetInput[] = []
  const cached: XlsxRead['cached'] = []
  for (const d of decls) {
    if (d.kind !== 'worksheet') {
      issues.add('sheetKind', 'a chart, dialog or macro sheet is left out', { sheet: d.name })
      continue
    }
    const xml = d.part === undefined ? undefined : await archive.text(d.part)
    if (xml === undefined) {
      issues.add('part', "a sheet whose part is missing is left out", { sheet: d.name })
      continue
    }
    if (d.hidden) issues.add('hidden', 'a hidden sheet is read as a visible one', { sheet: d.name })
    const read = readSheet(
      { name: d.name, xml, rels: await relsOf(archive, d.part!), sheets: all },
      { sheet: d.name, strings, xfs: table.xfs, meta, date: dateOf, issues },
    )
    sheets.push(read.sheet)
    cached.push({ name: d.name, cells: read.cached })
    checkIntersections(d.name, read.old, names, issues)
  }
  const book: BookInput = { sheets }
  if (names.length) book.names = names
  if (table.styles.length) book.styles = table.styles
  const tag = await locale(archive, root, issues)
  if (tag) book.locale = tag
  try {
    checkBook(book)
  } catch (e) {
    if (e instanceof InputError) throw new Error(`xlsx: the file's book is not valid: ${e.message}`)
    throw e
  }
  return { book, cached, issues: issues.list() }
}
