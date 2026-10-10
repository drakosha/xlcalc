/**
 * `readXlsx`: an xlsx file as a book. The workbook part is found through the package's relationships, its
 * sheets taken in the order the workbook lists them (not by part names), each linked to its part by id; shared
 * strings, styles, the theme and the cell metadata of dynamic arrays come from the workbook's relationships.
 * Everything the book cannot hold is reported in `issues`, nothing is dropped silently.
 */
import { caseKey } from '../../collate.js'
import type { BookInput, NameInput, SheetInput } from '../../contract.js'
import { epochOf } from '../../dates.js'
import { checkBook, InputError } from '../../input.js'
import { LOCALE_FILES } from '../../locale/registry.js'
import { Issues } from '../issues.js'
import { LOCALE_PROPERTY } from '../property.js'
import { palette } from '../styles/color.js'
import { readStyles } from '../styles/read.js'
import { readDxfs } from '../styles/dxf.js'
import type { XlsxRead } from '../types.js'
import { attr, attributes, element, elements, flag, has } from '../xml/scan.js'
import { decodeXml, readText } from '../xml/text.js'
import { openArchive, type Archive } from '../zip/read.js'
import { legacyConverter } from '../../legacy/convert.js'
import { readFormula } from './formulas.js'
import { cellMetadata } from './metadata.js'
import { relsOf, type Rel } from './package.js'
import { readSheet } from './sheet.js'
import { readStrings } from './strings.js'
import { formatAddress } from '../../model.js'
import { parseRange } from '../../structure/area.js'

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
  if (LOCALE_FILES.has(tag)) return tag
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
    const f = readFormula(readText(e.body ?? ''), issues, { name }, true)
    if (flag(a.hidden)) issues.add('name', 'a hidden name is read as a visible one', { name })
    out.push(scope === undefined ? { name, ref: f.formula } : { name, ref: f.formula, scope })
  }
  return out
}

/** The calculation settings the book keeps (iterative calculation), and the others reported. */
function calculation(xml: string, issues: Issues): BookInput['iteration'] {
  const calc = attributes(element(xml, 'calcPr')?.attrs ?? '')
  if (calc.calcMode === 'manual') issues.add('calculation', 'the manual calculation mode is not kept')
  if (calc.calcMode === 'autoNoTable') issues.add('calculation', 'the calculation mode "automatic except for data tables" is not kept')
  if (!flag(calc.fullPrecision, true)) issues.add('calculation', 'precision as displayed is not in the engine: values keep their full precision')
  if (!flag(calc.iterate)) return undefined
  const out: NonNullable<BookInput['iteration']> = {}
  const count = Number(calc.iterateCount)
  if (calc.iterateCount !== undefined && Number.isInteger(count) && count >= 1 && count <= 32_767) out.maxIterations = count
  else if (calc.iterateCount !== undefined) issues.add('calculation', `the maximum of iterations ${JSON.stringify(calc.iterateCount)} is not one of 1 to 32,767: the default, 100, is read`)
  const change = Number(calc.iterateDelta)
  if (calc.iterateDelta !== undefined && change >= 0 && change <= Number.MAX_VALUE) out.maxChange = change
  else if (calc.iterateDelta !== undefined) issues.add('calculation', `the maximum change of iterations ${JSON.stringify(calc.iterateDelta)} is not a number from 0: the default, 0.001, is read`)
  return out
}

function workbookIssues(xml: string, rels: Rel[], issues: Issues) {
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
  const [stringsXml, stylesXml, themeXml, metaXml, richValues, richStructures, persons] = await Promise.all(
    ['sharedStrings', 'styles', 'theme', 'sheetMetadata', 'rdRichValue', 'rdRichValueStructure', 'person'].map(part),
  )
  const colors = palette(themeXml, stylesXml)
  const table = readStyles(stylesXml, colors)
  const cond = { dxfs: readDxfs(stylesXml, colors), palette: colors }
  if (table.font && (table.font.name !== 'Calibri' || table.font.size !== 11))
    issues.add('defaultFont', `the default font ${table.font.name ?? '(no name)'} ${table.font.size ?? '(no size)'} is read as Calibri 11, the model's default`)
  const iteration = calculation(wb, issues)
  workbookIssues(wb, rels, issues)
  const dateSystem = flag(attributes(element(wb, 'workbookPr')?.attrs ?? '').date1904) ? 1904 : 1900
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
  const legacy = legacyConverter(names)
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
    const sheetRels = await relsOf(archive, d.part!)
    const tableParts = await Promise.all(sheetRels.filter((r) => r.kind === 'table' && !r.external).map((r) => archive.text(r.target)))
    // the VML drawing that shows the notes is the sheet's legacy drawing
    const drawing = attr(attributes(element(xml, 'legacyDrawing')?.attrs ?? ''), 'id')
    const linked = async (kind: string, id?: string) => {
      const r = sheetRels.find((x) => x.kind === kind && !x.external && (id === undefined || x.id === id))
      return r ? archive.text(r.target) : undefined
    }
    const [comments, vml, threads] = await Promise.all([linked('comments'), drawing === undefined ? undefined : linked('vmlDrawing', drawing), linked('threadedComment')])
    const read = readSheet(
      { name: d.name, xml, rels: sheetRels, sheets: all, legacy, tables: tableParts.filter((x): x is string => x !== undefined), parts: { comments, vml, threads, persons }, cond },
      { sheet: d.name, strings, xfs: table.xfs, meta, epoch: epochOf(dateSystem), issues },
    )
    sheets.push(read.sheet)
    cached.push({ name: d.name, cells: read.cached })
  }
  // table names are unique through the book and apart from the defined names
  const taken = new Set(names.map((n) => n.name.toUpperCase()))
  for (const sh of sheets) {
    if (!sh.tables) continue
    sh.tables = sh.tables.filter((t) => {
      const k = caseKey(t.name)
      if (!taken.has(k)) return taken.add(k)
      issues.add('table', `the table ${t.name} is left out: its name is another table's or a defined name`, { sheet: sh.name })
      return false
    })
    if (!sh.tables.length) delete sh.tables
  }
  const book: BookInput = { sheets }
  if (names.length) book.names = names
  if (table.styles.length) book.styles = table.styles
  const tag = await locale(archive, root, issues)
  if (tag) book.locale = tag
  if (dateSystem === 1904) book.dateSystem = 1904
  if (iteration) book.iteration = iteration
  // the cells are valid as the reader builds them: the rest of the book is checked (the cells' check is most of the time)
  try {
    checkBook({ ...book, sheets: sheets.map((s) => ({ ...s, cells: headerCells(s) })) })
  } catch (e) {
    if (e instanceof InputError) throw new Error(`xlsx: the file's book is not valid: ${e.message}`)
    throw e
  }
  return { book, cached, issues: issues.list() }
}

/** The header cells of a sheet's tables, which the check of its tables reads. */
function headerCells(s: SheetInput): SheetInput['cells'] {
  const out: SheetInput['cells'] = {}
  for (const t of s.tables ?? []) {
    const x = parseRange(t.range)
    if (!x || t.headers === false) continue
    for (let c = x.c1; c <= x.c2; c++) {
      const a = formatAddress(x.r1, c)
      if (s.cells[a]) out[a] = s.cells[a]
    }
  }
  return out
}
