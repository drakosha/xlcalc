import { relsOf } from '../xlsx/read/package.js'
import type { XlsxLimits } from '../xlsx/types.js'
import { attributes, elements } from '../xlsx/xml/scan.js'
import { decodeXml } from '../xlsx/xml/text.js'
import type { Archive } from '../xlsx/zip/read.js'
import { comments } from './comments.js'
import { externalIndex, externalLink, type ExternalSource } from './external.js'
import { metadata } from './metadata.js'
import { connections, workbook, type Context, type Part } from './parts.js'
import { cacheDefinition, cacheRecords } from './pivot-cache.js'
import { pivotTable } from './pivot-table.js'
import { records } from './records.js'
import { slicerCache, slicers } from './slicers.js'
import { sharedStrings } from './strings.js'
import { styles } from './styles.js'
import { table, tableIndex } from './table.js'
import { worksheet } from './worksheet.js'
import { Xml, type Budget } from './xml.js'

type Convert = (part: Part) => void

const converters: Readonly<Partial<Record<string, Convert>>> = {
  'application/vnd.ms-excel.worksheet': worksheet,
  'application/vnd.ms-excel.sharedstrings': sharedStrings,
  'application/vnd.ms-excel.styles': styles,
  'application/vnd.ms-excel.sheetmetadata': metadata,
  'application/vnd.ms-excel.table': table,
  'application/vnd.ms-excel.comments': comments,
  'application/vnd.ms-excel.connections': connections,
  'application/vnd.ms-excel.pivottable': pivotTable,
  'application/vnd.ms-excel.slicer': slicers,
  'application/vnd.ms-excel.slicercache': slicerCache,
}

/** The relationship graph keeps its original paths, so binary and XML parts can point to each other. */
export async function convertingArchive(
  source: Archive, workbookPath: string, limits: XlsxLimits, budget: Budget,
  unfinished: Set<string>, invalidExternalDefinition: (part: string, name: string) => void,
): Promise<Archive> {
  const types = new Map<string, string>()
  const external = new Map<string, ExternalSource>()
  const context: Context = {
    fonts: [], strings: 0, dxfs14: [], sheets: [], sheetsByRel: new Map(), names: [], xti: [], linkSources: [], external,
    metadata: { cell: 0, value: 0 }, tables: new Map(),
  }
  const contentTypes = await source.text('[Content_Types].xml')
  for (const { attrs } of elements(contentTypes ?? '', 'Override')) {
    const { PartName, ContentType } = attributes(attrs)
    if (PartName && ContentType) types.set(decodeXml(PartName).replace(/^\//, '').toLowerCase(), ContentType.toLowerCase())
  }
  const loaded = new Map<string, Uint8Array>()
  const cacheFields = new Map<string, ReturnType<typeof cacheDefinition>>()
  const slash = workbookPath.lastIndexOf('/') + 1
  const workbookRelsPath = `${workbookPath.slice(0, slash)}_rels/${workbookPath.slice(slash)}.rels`
  const relsText = await source.text(workbookRelsPath)
  const sourceText = (path: string) => path.toLowerCase() === workbookRelsPath.toLowerCase() ? Promise.resolve(relsText) : source.text(path)
  const links = await relsOf({ ...source, text: sourceText }, workbookPath)
  const sheetRels = new Map(links.filter((r) => r.kind === 'worksheet' && !r.external).map((r) => [r.target.toLowerCase(), r.id]))
  const stylesPath = links.find((r) => r.kind === 'styles' && !r.external)?.target.toLowerCase()
  for (const { id, kind, target, external: outside } of links) {
    if (kind !== 'externalLink' || outside || !target.toLowerCase().endsWith('.bin')) continue
    const bytes = await source.bytes(target)
    if (!bytes) continue
    loaded.set(target.toLowerCase(), bytes)
    external.set(id, externalIndex(bytes, target))
  }
  for (const [path, kind] of types) {
    if (kind !== 'application/vnd.ms-excel.table') continue
    const bytes = await source.bytes(path)
    if (!bytes) continue
    loaded.set(path, bytes)
    const index = tableIndex(bytes, path)
    if (index) context.tables.set(index.id, index)
  }
  let stylesXml: string | undefined
  if (stylesPath) {
    const bytes = await source.bytes(stylesPath)
    if (bytes) {
      const xml = new Xml(stylesPath, limits, budget)
      const pending = () => unfinished.add(stylesPath)
      styles({ bytes, path: stylesPath, xml, context, pending })
      stylesXml = xml.text()
    }
  }
  return {
    names: source.names,
    bytes: source.bytes,
    marked: source.marked,
    async text(path) {
      const key = path.toLowerCase()
      if (!key.endsWith('.bin')) return sourceText(path)
      if (key === stylesPath) return stylesXml
      const bytes = loaded.get(key) ?? await source.bytes(path)
      loaded.delete(key)
      if (bytes === undefined) return undefined
      const pending = () => unfinished.add(path)
      const xml = new Xml(path, limits, budget)
      const kind = types.get(key) ?? ''
      const sheetRel = sheetRels.get(key)
      const sheet = sheetRel === undefined ? undefined : context.sheetsByRel.get(sheetRel)
      const part = { bytes, path, xml, context, sheet, pending }
      if (kind === 'application/vnd.ms-excel.externallink') {
        externalLink(part, (name) => invalidExternalDefinition(path, name))
        return xml.text()
      }
      if (kind === 'application/vnd.ms-excel.pivotcachedefinition') {
        const fields = cacheDefinition(part)
        const data = (await relsOf(source, path)).find((r) => r.kind === 'pivotCacheRecords' && !r.external)
        if (data) cacheFields.set(data.target.toLowerCase(), fields)
        return xml.text()
      }
      if (kind === 'application/vnd.ms-excel.pivotcacherecords') {
        const fields = cacheFields.get(key)
        if (fields) cacheRecords(part, fields)
        else pending()
        return xml.text()
      }
      const convert = key === workbookPath.toLowerCase() ? workbook : Object.hasOwn(converters, kind) ? converters[kind] : undefined
      if (!convert) {
        for (const _record of records(bytes, path)) {
          // Check framing for a requested binary part even before its fields are translated.
        }
        pending()
        return undefined
      }
      convert(part)
      return xml.text()
    },
  }
}
