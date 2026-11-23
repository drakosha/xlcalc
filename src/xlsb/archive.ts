import type { XlsxLimits } from '../xlsx/types.js'
import { attributes, elements } from '../xlsx/xml/scan.js'
import { decodeXml } from '../xlsx/xml/text.js'
import type { Archive } from '../xlsx/zip/read.js'
import { worksheet, workbook, type Part } from './parts.js'
import { records } from './records.js'
import { Xml, type Budget } from './xml.js'

type Convert = (part: Part) => void

const converters: Readonly<Partial<Record<string, Convert>>> = {
  'application/vnd.ms-excel.worksheet': worksheet,
}

/** The relationship graph keeps its original paths, so binary and XML parts can point to each other. */
export async function convertingArchive(source: Archive, workbookPath: string, limits: XlsxLimits, budget: Budget, unfinished: Set<string>): Promise<Archive> {
  const types = new Map<string, string>()
  const contentTypes = await source.text('[Content_Types].xml')
  for (const { attrs } of elements(contentTypes ?? '', 'Override')) {
    const { PartName, ContentType } = attributes(attrs)
    if (PartName && ContentType) types.set(decodeXml(PartName).replace(/^\//, '').toLowerCase(), ContentType.toLowerCase())
  }
  return {
    names: source.names,
    bytes: source.bytes,
    marked: source.marked,
    async text(path) {
      const key = path.toLowerCase()
      if (!key.endsWith('.bin')) return source.text(path)
      const bytes = await source.bytes(path)
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
      convert({ bytes, path, xml, pending })
      return xml.text()
    },
  }
}
