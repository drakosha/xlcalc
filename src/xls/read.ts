import { hasCompoundSignature, readStyledArchive, workbookInput, type ReadWorksheetPart } from '../xlsx/read/book.js'
import { LOCALE_PROPERTY } from '../xlsx/property.js'
import type { Rel } from '../xlsx/read/package.js'
import { OFFICE_THEME, THEME_ORDER } from '../xlsx/styles/color.js'
import type { ReadXlsxOptions, XlsxLimits, XlsxRead } from '../xlsx/types.js'
import { escapeAttr, escapeText } from '../xlsx/xml/text.js'
import { concat } from '../xlsx/zip/platform.js'
import { ENTRY_BYTES, openArchive, TOTAL_BYTES, type Archive } from '../xlsx/zip/read.js'
import { Xml } from '../xlsb/xml.js'
import { conditionalRules } from './conditional.js'
import { readFilters, type SheetFilters } from './filter.js'
import { cacheParts } from './pivot-cache.js'
import { olapCacheIndexes, pivotViews, viewPart, type ViewCache } from './pivot-table.js'
import { notes } from './notes.js'
import type { ReadPrintMarkup } from './printing.js'
import { CompoundLayoutError, openCompound, type Compound } from './compound.js'
import { localeProperty } from './properties.js'
import { isBof } from './records.js'
import { globals, Losses, type Sheet } from './workbook.js'
import { restoreCalculatedColumns, tablePart, tables, type Table } from './tables.js'
import { sheetRecords, worksheet } from './worksheet.js'

const WORKBOOK = 'workbook.xml'
const STRINGS = 'strings.xml'
const STYLES = 'styles.xml'
const THEME = 'theme.xml'
const PROPERTIES = 'customProperties.xml'
const RELS = '_rels/workbook.xml.rels'
const OFFICE_2007 = ['FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646', '0000FF', '800080']

function defaultTheme(version: number | undefined): string {
  // MS-XLS lists both versions of the legacy default theme. With no record Excel uses the current palette.
  const colors = version === 124226 || version === 123820 ? OFFICE_2007 : OFFICE_THEME
  const scheme = THEME_ORDER.map((name, i) => `<${name}><srgbClr val="${colors[i]}"/></${name}>`).join('')
  return `<theme><clrScheme>${scheme}</clrScheme></theme>`
}

function workbookBytes(bytes: Uint8Array, limits: XlsxLimits): Compound {
  if (!hasCompoundSignature(bytes) && bytes.length >= 4 && isBof(bytes[0] | (bytes[1] << 8))) {
    const entry = limits.entryBytes ?? ENTRY_BYTES
    const total = limits.totalBytes ?? TOTAL_BYTES
    if (bytes.length > entry) throw new Error(`xls: Workbook length exceeds ${entry} bytes (limits.entryBytes)`)
    if (bytes.length > total) throw new Error(`xls: Workbook exceeds the ${total} byte total (limits.totalBytes)`)
    return { workbook: bytes, stream: () => undefined }
  }
  return openCompound(bytes, limits)
}

interface XlsParts {
  readPrint?: ReadPrintMarkup
  visitWorksheet?: ReadWorksheetPart
}

export async function readXls(
  data: Uint8Array | ArrayBuffer, opts: ReadXlsxOptions | undefined, parts?: XlsParts,
): Promise<{ read: XlsxRead; workbook: ReturnType<typeof globals> }> {
  const input = workbookInput(data, opts, true)
  const limits = input.limits ?? {}
  const compound = workbookBytes(input.bytes, limits)
  const bytes = compound.workbook
  const losses = new Losses(input.language?.tag === 'ru-RU')
  let properties: Uint8Array | undefined
  try {
    properties = compound.stream('', '\x05DocumentSummaryInformation')
  } catch (error) {
    if (!(error instanceof CompoundLayoutError)) throw error
    losses.add('part', 'invalid optional workbook properties were left out', 'Повреждённые дополнительные свойства книги пропущены')
  }
  const locale = localeProperty(properties)
  const budget = { spent: bytes.length + (properties?.length ?? 0) }
  const xml = (part: string) => new Xml(part, limits, budget, 'xls')
  const wbXml = xml(WORKBOOK)
  const book = globals(bytes, wbXml, () => xml(STRINGS), losses)
  const sheetParts = new Map<string, Sheet>()
  const rels = xml(RELS)
  rels.add('<Relationships>')
  const relate = (id: string, kind: string, path: string) => rels.add(`<Relationship Id="${id}" Type="${kind}" Target="${path}"/>`)
  relate('styles', 'styles', STYLES)
  let theme: string | undefined
  if (book.theme) {
    const themeArchive = openArchive(concat(book.theme), limits, budget)
    const path = themeArchive.names.find((name) => /(?:^|\/)theme[0-9]*\.xml$/i.test(name))
    if (path) theme = await themeArchive.text(path)
  }
  if (theme === undefined) {
    const defaults = xml(THEME)
    defaults.add(defaultTheme(book.themeVersion))
    theme = defaults.text()
  }
  relate('theme', 'theme', THEME)
  if (book.stringsXml !== undefined) relate('strings', 'sharedStrings', STRINGS)
  const extraParts = new Map<string, string>()
  const root: Rel[] = []
  if (locale !== undefined) {
    const part = xml(PROPERTIES)
    part.add(`<Properties><property name="${LOCALE_PROPERTY}"><lpwstr>${escapeText(locale)}</lpwstr></property></Properties>`)
    extraParts.set(PROPERTIES, part.text())
    root.push({ id: 'properties', kind: 'custom-properties', target: PROPERTIES, external: false })
  }
  const connectionIds = new Map<number, { id: number; olap: boolean }>()
  for (const [index, cache] of book.connectionCaches.entries()) {
    if (cache !== undefined && !connectionIds.has(cache.id)) connectionIds.set(cache.id, { id: index + 1, olap: cache.olap })
  }
  if (book.caches.some((cache) => cache.type === 2 && !connectionIds.has(cache.id))) {
    for (const sheet of book.sheets) {
      if (sheet.kind !== 'worksheet') continue
      for (const index of olapCacheIndexes(sheetRecords(book, sheet))) {
        const cache = book.caches[index]
        if (cache?.type === 2) cache.olap = true
      }
    }
  }
  if (book.connectionCaches.length || book.caches.some((cache) => cache.type === 2)) {
    // The shared reader reports the presence of connections from the relationship.
    relate('connections', 'connections', 'connections.xml')
  }
  const pivotCaches: (ViewCache | undefined)[] = []
  if (book.caches.length) {
    wbXml.add('<pivotCaches>')
    for (const [index, cache] of book.caches.entries()) {
      let data: Uint8Array | undefined
      try {
        data = compound.stream('_SX_DB_CUR', cache.id.toString(16).padStart(4, '0'))
      } catch (error) {
        if (!(error instanceof CompoundLayoutError)) throw error
        // The affected PivotTable reports its unavailable cache when its sheet is read.
        continue
      }
      if (!data) continue
      const total = limits.totalBytes ?? TOTAL_BYTES
      if ((budget.spent += data.length) > total) throw new Error(`xls: pivot cache exceeds the ${total} byte total (limits.totalBytes)`)
      const id = `pivotCache${index}`
      const part = id + '.xml'
      const def = xml(part)
      const recordsPath = id + '.records'
      const cacheRecords = xml(recordsPath)
      const connection = connectionIds.get(cache.id)
      cache.connection = connection?.id
      cache.olap ||= connection?.olap
      const stored = cacheParts(cache, data, def, cacheRecords, book.context)
      if (!stored) continue
      pivotCaches[index] = { id: cache.id, fields: new Map(stored.names.map((name, index) => [name, index])) }
      relate(id, 'pivotCacheDefinition', part)
      wbXml.add(`<pivotCache cacheId="${cache.id}" r:id="${id}"/>`)
      extraParts.set(part, def.text())
      if (stored.saved) extraParts.set(recordsPath, cacheRecords.text())
      if (stored.saved || stored.books.length) {
        const relationsPath = `_rels/${part}.rels`
        const relations = xml(relationsPath)
        relations.add('<Relationships>')
        if (stored.saved) relations.add(`<Relationship Id="records" Type="pivotCacheRecords" Target="${recordsPath}"/>`)
        for (const [i, target] of stored.books.entries()) {
          relations.add(`<Relationship Id="source${i}" Type="externalLinkPath" Target="${escapeAttr(target)}" TargetMode="External"/>`)
        }
        relations.add('</Relationships>')
        extraParts.set(relationsPath, relations.text())
      }
    }
    wbXml.add('</pivotCaches>')
  }
  let calcRead = false
  const metadata = new Map<Sheet, { filters: SheetFilters; conditional?: Xml; pivots: ReturnType<typeof pivotViews> }>()
  for (let i = 0; i < book.sheets.length; i++) {
    const sheet = book.sheets[i]
    if (sheet.kind === 'worksheet') {
      for (const { type, payload } of sheetRecords(book, sheet)) {
        if (type === 0x0081) {
          if (payload.u16() & 16) sheet.kind = 'dialogsheet'
          break
        }
        if (!calcRead) book.calc.take(type, payload)
      }
      if (sheet.kind === 'worksheet') calcRead = true
    }
    const path = `sheet${i}.xml`
    relate(`sheet${i}`, sheet.kind, path)
    if (sheet.kind === 'worksheet') {
      sheetParts.set(path, sheet)
      if (book.encoding?.version === 5) {
        const links = book.external.legacy(book.encoding)
        for (const { type, payload } of sheetRecords(book, sheet)) links.take(type, payload)
        sheet.links = links.references
      }
      const filters = readFilters(sheetRecords(book, sheet), book, sheet)
      const pivots = pivotViews(sheetRecords(book, sheet), pivotCaches, book.styles)
      const pivotPriorities = new Set(pivots.flatMap((pivot) => pivot.conditional.priorities()))
      const rules = xml(path + '.rules')
      conditionalRules(sheetRecords(book, sheet), book, sheet, rules, pivotPriorities)
      metadata.set(sheet, { filters, pivots, conditional: rules.empty ? undefined : rules })
    }
  }
  const externalParts = new Map<string, number>()
  const externalRelations = new Map<string, number>()
  if (book.external.links.length) {
    wbXml.add('<externalReferences>')
    for (const [index] of book.external.links.entries()) {
      const id = `external${index}`
      const part = `${id}.xml`
      relate(id, 'externalLink', part)
      externalParts.set(part, index)
      externalRelations.set(`_rels/${part}.rels`, index)
      wbXml.add(`<externalReference r:id="${id}"/>`)
    }
    wbXml.add('</externalReferences>')
  }
  rels.add('</Relationships>')
  const relsText = rels.text()
  wbXml.add(book.calc.write() + '</workbook>')
  const workbook = wbXml.text()
  const fallback = new Map<Sheet, Set<string>>()
  const calculated = new Map<Sheet, Table[]>()
  const archive: Archive = {
    names: [WORKBOOK, STYLES, STRINGS, ...sheetParts.keys()],
    async bytes() { return undefined },
    marked() { return false },
    async text(path) {
      if (path === WORKBOOK) return workbook
      if (path === STRINGS) return book.stringsXml
      if (path === THEME) return theme
      if (path === STYLES) return book.styles.write(xml(path))
      if (path === RELS) return relsText
      const external = externalParts.get(path)
      if (external !== undefined) return book.external.write(external, xml(path))
      const relation = externalRelations.get(path)
      if (relation !== undefined) {
        const target = book.external.links[relation].target
        if (!target) return undefined
        const part = xml(path)
        part.add(`<Relationships><Relationship Id="target" Type="externalLinkPath" TargetMode="External" Target="${escapeAttr(target)}"/></Relationships>`)
        return part.text()
      }
      const sheet = sheetParts.get(path)
      if (!sheet) return extraParts.get(path)
      const notePath = path + '.notes'
      const vmlPath = path + '.vml'
      const noteXml = xml(notePath)
      const vml = xml(vmlPath)
      const hasNotes = notes(sheetRecords(book, sheet), book, sheet, noteXml, vml)
      const relPath = '_rels/' + path + '.rels'
      const relations = xml(relPath)
      relations.add('<Relationships>')
      const relateSheet = (id: string, kind: string, target: string, external = false) => {
        relations.add(`<Relationship Id="${id}" Type="${kind}" Target="${escapeAttr(target)}"${external ? ' TargetMode="External"' : ''}/>`)
      }
      if (hasNotes) {
        relateSheet('notes', 'comments', notePath)
        relateSheet('noteDrawing', 'vmlDrawing', vmlPath)
        extraParts.set(notePath, noteXml.text())
        extraParts.set(vmlPath, vml.text())
      }
      const { filters, conditional, pivots } = metadata.get(sheet)!
      const sheetTables = tables(sheetRecords(book, sheet))
      if (sheetTables.tables.some((table) => table.calculated.length)) calculated.set(sheet, sheetTables.tables)
      for (const [index, table] of sheetTables.tables.entries()) {
        const id = `table${index}`
        const part = `${path}.${id}`
        extraParts.set(part, tablePart(table, xml(part), filters.tables.get(table.id)))
        relateSheet(id, 'table', part)
      }
      if (sheetTables.missed) losses.add('table', 'a BIFF table was left out; its cells were read as stored', 'Таблица BIFF пропущена, ячейки прочитаны как в файле', sheet.name)
      for (const [index, pivot] of pivots.entries()) {
        const id = `pivot${index}`
        const part = `${path}.${id}`
        const description = xml(part)
        if (viewPart(pivot, description)) {
          extraParts.set(part, description.text())
          relateSheet(id, 'pivotTable', part)
          if (pivot.leftOut) losses.add('pivotTable', 'some of the stored pivot layout was left out', 'Часть разметки сводной таблицы пропущена', sheet.name)
        } else losses.add('pivotTable', 'a pivot table with an unsupported cache or layout was left out', 'Сводная таблица с неподдерживаемым кэшем или разметкой пропущена', sheet.name)
      }
      const output = xml(path)
      const print = parts?.readPrint?.(sheetRecords(book, sheet), book.encoding, sheet.name)
      const missing = worksheet(book, sheet, output, hasNotes, relateSheet, filters, conditional, print)
      if (missing.size) fallback.set(sheet, missing)
      relations.add('</Relationships>')
      extraParts.set(relPath, relations.text())
      return output.text()
    },
  }
  const { read } = await readStyledArchive(archive, root, WORKBOOK, input.language, parts?.visitWorksheet)
  const refreshFromView = new Set<number>()
  for (const [index, view] of pivotCaches.entries()) {
    if (view?.legacyViewEnablesRefresh && book.caches[index].enableRefresh === undefined) refreshFromView.add(view.id)
  }
  // Older cache versions can enable refresh in the view's QsiSXTag, in addition to the cache stream.
  for (const cache of read.book.pivotCaches ?? []) {
    if (refreshFromView.has(cache.id)) delete cache.enableRefresh
  }
  let index = 0
  for (const sheet of sheetParts.values()) {
    const modelSheet = read.book.sheets[index]
    const tableColumns = calculated.get(sheet)
    if (tableColumns) restoreCalculatedColumns(tableColumns, modelSheet)
    const { cells } = modelSheet
    const saved = read.cached.sheets[index++].cells
    for (const at of fallback.get(sheet) ?? []) {
      const value = cells[at]?.value
      if (value !== undefined && typeof value !== 'object') saved[at] ??= { value }
    }
  }
  read.issues.push(...losses.list())
  return { read, workbook: book }
}
