/**
 * `readXlsx`: an xlsx file as a book. The workbook part is found through the package's relationships, its
 * sheets taken in the order the workbook lists them (not by part names), each linked to its part by id; shared
 * strings, styles, the theme, the cell metadata of dynamic arrays and the links to other workbooks (`links.ts`) come from
 * the workbook's relationships.
 * Everything the book cannot hold is reported in `issues`, nothing is dropped silently.
 */
import { caseKey } from '../../collate.js'
import type { BookInput, NameInput, SheetInput } from '../../contract.js'
import { epochOf } from '../../dates.js'
import { checkBook, checkNumber, InputError, ITERATION, object, onlyKeys, onlyOptions, SHEET_STATES, what } from '../../input.js'
import { LOCALES, PACKAGE_LOCALES } from '../../locale/registry.js'
import { issueLocale, Issues, type XlsxMessage } from '../issues.js'
import { LOCALE_PROPERTY } from '../property.js'
import { palette } from '../styles/color.js'
import { readStyles } from '../styles/read.js'
import { readDxfs } from '../styles/dxf.js'
import type { ReadXlsxOptions, XlsxLimits, XlsxRead } from '../types.js'
import { attr, attributes, element, elements, flag, has, ownEntry } from '../xml/scan.js'
import { decodeXml, readText } from '../xml/text.js'
import { openArchive, type Archive } from '../zip/read.js'
import { legacyConverter } from '../../legacy/convert.js'
import { readFormula } from './formulas.js'
import { readLinks } from './links.js'
import { cellMetadata } from './metadata.js'
import { relsOf, type Rel } from './package.js'
import { readSheet } from './sheet.js'
import { readActiveSheet } from '../view.js'
import { readStrings } from './strings.js'
import { formatAddress } from '../../model.js'
import { parseRange } from '../../structure/area.js'

/** Parts the workbook links to that the model has no place for, by relationship kind. */
const LINKED: Record<string, ['macros' | 'pivotTable' | 'connection', XlsxMessage]> = {
  vbaProject: ['macros', { key: 'macros' }],
  pivotCacheDefinition: ['pivotTable', { key: 'pivotTables' }],
  connections: ['connection', { key: 'connections' }],
}

async function locale(archive: Archive, rels: Rel[], issues: Issues): Promise<string | undefined> {
  const part = rels.find((r) => r.kind === 'custom-properties' && !r.external)
  const xml = part && (await archive.text(part.target))
  const p = xml && elements(xml, 'property').find((e) => decodeXml(attributes(e.attrs).name ?? '') === LOCALE_PROPERTY)
  if (!p) return undefined
  const tag = readText(element(p.body ?? '', 'lpwstr')?.body ?? '')
  // the book's locale is data: a package's locale not loaded stays, and computing the book names its module
  if (PACKAGE_LOCALES.includes(tag) || LOCALES.has(tag)) return tag
  issues.add('locale', { key: 'localeUnknown', tag })
  return undefined
}

interface SheetDecl {
  name: string
  kind?: string
  part?: string
  /** xlsx `state`: hidden, or very hidden. */
  hidden?: true | 'very'
}

/** A sheet's `hidden` by its state in the file. */
const STATES = new Map<string | undefined, true | 'very'>([...SHEET_STATES].map(([hidden, state]) => [state, hidden]))

function readNames(xml: string, decls: SheetDecl[], read: Set<string>, issues: Issues, links: Set<number>): NameInput[] {
  const out: NameInput[] = []
  for (const e of elements(element(xml, 'definedNames')?.body ?? '', 'definedName')) {
    const a = attributes(e.attrs)
    const name = readText(a.name ?? '')
    // Excel's placeholders for functions newer than the file's readers: not names of the book
    if (/^_xlfn\./i.test(name)) continue
    if (/^_xlnm\./i.test(name)) {
      issues.add('name', { key: 'builtInName' }, { name })
      continue
    }
    const scope = a.localSheetId === undefined ? undefined : decls[Number(a.localSheetId)]?.name
    if (a.localSheetId !== undefined && (scope === undefined || !read.has(scope))) {
      issues.add('name', { key: 'nameOfSheetLeftOut' }, { name })
      continue
    }
    const f = readFormula(readText(e.body ?? ''), issues, { name }, { union: true, links })
    if (flag(a.hidden)) issues.add('name', { key: 'hiddenName' }, { name })
    out.push(scope === undefined ? { name, ref: f.formula } : { name, ref: f.formula, scope })
  }
  return out
}

/** The calculation settings the book keeps (iterative calculation), and the others reported. */
function calculation(xml: string, issues: Issues): BookInput['iteration'] {
  const calc = attributes(element(xml, 'calcPr')?.attrs ?? '')
  if (calc.calcMode === 'manual') issues.add('calculation', { key: 'manualCalculation' })
  if (calc.calcMode === 'autoNoTable') issues.add('calculation', { key: 'autoNoTable' })
  if (!flag(calc.fullPrecision, true)) issues.add('calculation', { key: 'precisionAsDisplayed' })
  if (!flag(calc.iterate)) return undefined
  const out: NonNullable<BookInput['iteration']> = {}
  const count = Number(calc.iterateCount)
  const [least, most] = ITERATION.maxIterations
  if (calc.iterateCount !== undefined && Number.isInteger(count) && count >= least && count <= most) out.maxIterations = count
  else if (calc.iterateCount !== undefined) issues.add('calculation', { key: 'iterationCount', count: calc.iterateCount })
  const change = Number(calc.iterateDelta)
  if (calc.iterateDelta !== undefined && change >= ITERATION.maxChange[0] && change <= ITERATION.maxChange[1]) out.maxChange = change
  else if (calc.iterateDelta !== undefined) issues.add('calculation', { key: 'iterationChange', change: calc.iterateDelta })
  return out
}

function workbookIssues(xml: string, rels: Rel[], issues: Issues) {
  if (has(xml, 'workbookProtection')) issues.add('protection', { key: 'workbookProtection' })
  if (has(xml, 'pivotCaches') || has(xml, 'slicerCaches')) issues.add('pivotTable', { key: 'pivotTables' })
  for (const r of rels) {
    const linked = ownEntry(LINKED, r.kind)
    if (linked) issues.add(...linked)
  }
}

/** The file's bytes: a `Uint8Array` (a Node `Buffer` is one), another view of bytes, or an `ArrayBuffer`. */
function fileBytes(data: unknown): Uint8Array {
  if (data instanceof Uint8Array) return data
  if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength)
  if (data instanceof ArrayBuffer || Object.prototype.toString.call(data) === '[object ArrayBuffer]') return new Uint8Array(data as ArrayBuffer)
  const hint = typeof (data as { arrayBuffer?: unknown } | null)?.arrayBuffer === 'function' ? 'await blob.arrayBuffer()' : 'await readFile(path)'
  throw new InputError('data', `expected the file's bytes, got ${what(data)}: pass the file's bytes: Uint8Array, ArrayBuffer or Buffer, e.g. ${hint}`)
}

const LIMITS = ['entryBytes', 'totalBytes'] as const

/** `opts.limits`: whole numbers of bytes of at least 1. */
function checkLimits(v: unknown): XlsxLimits {
  const o = object(v, 'opts.limits', 'limits such as { entryBytes: 16777216, totalBytes: 67108864 }')
  onlyKeys(o, 'opts.limits', LIMITS)
  for (const k of LIMITS) if (o[k] !== undefined) checkNumber(o[k], `opts.limits.${k}`, 1, Number.MAX_SAFE_INTEGER, true)
  return o as XlsxLimits
}

/**
 * Reads an xlsx file (bytes of the file) into a book, the values Excel stored and what the book could not hold, its
 * messages in `opts.locale` (en-US by default).
 */
export async function readXlsx(data: Uint8Array | ArrayBuffer, opts?: ReadXlsxOptions): Promise<XlsxRead> {
  const bytes = fileBytes(data)
  const o = opts === undefined ? {} : object(opts, 'opts', "options such as { locale: 'ru-RU' }")
  onlyOptions(o, 'opts', ['locale', 'limits'])
  const language = issueLocale(o)
  if (o.limits !== undefined) checkLimits(o.limits)
  const archive = openArchive(bytes)
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
  // a cell or a row that a crossing line's look does not reach is read with a style with nothing set
  let plainStyle: number | undefined
  const plain = () => (plainStyle ??= table.styles.push({}) - 1)
  const isPlain = (style: number) => style === plainStyle
  const cond = { dxfs: readDxfs(stylesXml, colors), palette: colors }
  if (table.font && (table.font.name !== 'Calibri' || table.font.size !== 11))
    issues.add('defaultFont', { key: 'defaultFont', ...table.font })
  const iteration = calculation(wb, issues)
  workbookIssues(wb, rels, issues)
  const dateSystem = flag(attributes(element(wb, 'workbookPr')?.attrs ?? '').date1904) ? 1904 : 1900
  const byId = new Map(rels.map((r) => [r.id, r]))
  const decls: SheetDecl[] = elements(element(wb, 'sheets')?.body ?? '', 'sheet').map((e) => {
    const a = attributes(e.attrs)
    const r = byId.get(attr(a, 'id') ?? '')
    return { name: readText(a.name ?? ''), kind: r?.kind, part: r?.external ? undefined : r?.target, ...(STATES.has(a.state) ? { hidden: STATES.get(a.state) } : {}) }
  })
  const strings = readStrings(stringsXml)
  const meta = cellMetadata(metaXml, richValues, richStructures)
  const links = await readLinks(archive, wb, rels, issues)
  const all = decls.filter((d) => d.kind === 'worksheet').map((d) => d.name)
  const names = readNames(wb, decls, new Set(all), issues, links.known)
  const legacy = legacyConverter(names)
  const sheets: SheetInput[] = []
  const cached: XlsxRead['cached'] = { sheets: [] }
  for (const d of decls) {
    if (d.kind !== 'worksheet') {
      issues.add('sheetKind', { key: 'sheetKind' }, { sheet: d.name })
      continue
    }
    const xml = d.part === undefined ? undefined : await archive.text(d.part)
    if (xml === undefined) {
      issues.add('part', { key: 'sheetPartMissing' }, { sheet: d.name })
      continue
    }
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
      { name: d.name, xml, rels: sheetRels, sheets: all, legacy, tables: tableParts.filter((x): x is string => x !== undefined), parts: { comments, vml, threads, persons }, cond, formulas: { links: links.known } },
      { sheet: d.name, strings, xfs: table.xfs, meta, epoch: epochOf(dateSystem), issues, plain, isPlain },
    )
    sheets.push(d.hidden ? { ...read.sheet, hidden: d.hidden } : read.sheet)
    cached.sheets.push({ name: d.name, cells: read.cached })
  }
  // a book keeps one visible sheet: a file whose visible sheets are all charts shows its first worksheet
  if (sheets.length && sheets.every((s) => s.hidden)) {
    issues.add('hidden', { key: 'allSheetsHidden' }, { sheet: sheets[0].name })
    delete sheets[0].hidden
  }
  // table names are unique through the book and apart from the defined names
  const taken = new Set(names.map((n) => caseKey(n.name)))
  for (const sh of sheets) {
    if (!sh.tables) continue
    sh.tables = sh.tables.filter((t) => {
      const k = caseKey(t.name)
      if (!taken.has(k)) return taken.add(k)
      issues.add('table', { key: 'tableNameTwice', table: t.name }, { sheet: sh.name })
      return false
    })
    if (!sh.tables.length) delete sh.tables
  }
  const book: BookInput = { sheets }
  const activeSheet = readActiveSheet(wb, decls, sheets)
  if (activeSheet) book.activeSheet = activeSheet
  if (names.length) book.names = names
  if (table.styles.length) book.styles = table.styles
  const tag = await locale(archive, root, issues)
  if (tag) book.locale = tag
  if (dateSystem === 1904) book.dateSystem = 1904
  if (iteration) book.iteration = iteration
  if (links.books.length) book.externalBooks = links.books
  // the cells are valid as the reader builds them: the rest of the book is checked (the cells' check is most of the time),
  // by a locale not loaded as by en-US
  try {
    checkBook({ ...book, locale: tag && LOCALES.has(tag) ? tag : undefined, sheets: sheets.map((s) => ({ ...s, cells: headerCells(s) })) })
  } catch (e) {
    if (e instanceof InputError) throw new Error(`xlsx: the file's book is not valid: ${e.message}`)
    throw e
  }
  return { book, cached, issues: issues.list(language) }
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
