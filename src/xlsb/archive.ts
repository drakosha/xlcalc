import type { XlsxLimits } from '../xlsx/types.js'
import { attributes, elements } from '../xlsx/xml/scan.js'
import { decodeXml } from '../xlsx/xml/text.js'
import type { Archive } from '../xlsx/zip/read.js'
import { metadata } from './metadata.js'
import { workbook, type Context, type Part } from './parts.js'
import { records } from './records.js'
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
}

/** The relationship graph keeps its original paths, so binary and XML parts can point to each other. */
export async function convertingArchive(source: Archive, workbookPath: string, limits: XlsxLimits, budget: Budget, unfinished: Set<string>): Promise<Archive> {
  const types = new Map<string, string>()
  const context: Context = {
    fonts: [], strings: 0, sheets: [], names: [], xti: [], linkSources: [], metadata: { cell: 0, value: 0 }, tables: new Map(),
  }
  const contentTypes = await source.text('[Content_Types].xml')
  for (const { attrs } of elements(contentTypes ?? '', 'Override')) {
    const { PartName, ContentType } = attributes(attrs)
    if (PartName && ContentType) types.set(decodeXml(PartName).replace(/^\//, '').toLowerCase(), ContentType.toLowerCase())
  }
  const stylesPath = [...types].find(([, kind]) => kind === 'application/vnd.ms-excel.styles')?.[0]
  const loaded = new Map<string, Uint8Array>()
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
      if (!key.endsWith('.bin')) return source.text(path)
      if (key === stylesPath) return stylesXml
      const bytes = loaded.get(key) ?? await source.bytes(path)
      loaded.delete(key)
      if (bytes === undefined) return undefined
      const pending = () => unfinished.add(path)
      const xml = new Xml(path, limits, budget)
      const kind = types.get(key) ?? ''
      const convert = key === workbookPath.toLowerCase() ? workbook : Object.hasOwn(converters, kind) ? converters[kind] : undefined
      if (!convert) {
        for (const _record of records(bytes, path)) {
          // Check framing for a requested binary part even before its fields are translated.
        }
        pending()
        return undefined
      }
      convert({ bytes, path, xml, context, pending })
      return xml.text()
    },
  }
}
