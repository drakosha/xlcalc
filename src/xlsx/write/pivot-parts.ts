/**
 * A workbook's PivotTables as parts of a package: each pivot cache's definition (`xl/pivotCache/pivotCacheDefinitionN.xml`)
 * with its rows when it saves them, each PivotTable (`xl/pivotTables/pivotTableN.xml`) pointing at its cache, and what
 * the workbook and its sheets add: the workbook's `pivotCaches` and the sheets' relationships to their PivotTables.
 */
import type { BookInput, ConditionalFormat, SheetInput } from '../../contract.js'
import { caseKey } from '../../collate.js'
import { pivotCacheXml, pivotRecordsXml } from './pivot-cache.js'
import { pivotTableXml, type PivotWriting } from './pivot.js'
import type { CellFormats } from './styles.js'
import { REL, relsXml as rels, TYPE } from './xml.js'
import type { Issues } from '../issues.js'
import { sheetKey } from '../../syntax/print.js'

/** The file a source on a deleted sheet links to: Excel writes the workbook's own path, which a file written anew does not know. */
const DELETED_SOURCE = 'book.xlsx'

export interface PivotFile {
  name: string
  text: string
}

/** What the PivotTables add to a package. */
export interface PivotPackage {
  files: PivotFile[]
  contentTypes: [part: string, type: string][]
  /** The workbook's relationships to the caches' definitions, by cache in the book's order. */
  caches: { id: number; target: string }[]
  /** Each sheet's relationships to its PivotTables (the target from the sheet's part). */
  sheets: string[][]
}


/** The rules of a sheet's PivotTable `name` with their priorities (`priority` gives the rule's at index i, `i + 1` without). */
export function pivotRules(sheet: SheetInput, name: string, priority: (i: number) => number = (i) => i + 1): PivotWriting['rules'] {
  const out: PivotWriting['rules'] = []
  ;(sheet.conditionalFormats ?? []).forEach((r: ConditionalFormat, i) => {
    const p = r.pivot
    if (p && caseKey(p.pivot) === caseKey(name)) out.push({ priority: priority(i), scope: p.scope, type: p.type, areas: p.areas })
  })
  return out
}

/** The parts of the book's caches and PivotTables; `styles` holds the number formats and the differential formats they use. */
export function pivotPackage(book: BookInput, styles: CellFormats, issues: Issues, slicerIds: ReadonlySet<number> = new Set()): PivotPackage {
  const out: PivotPackage = { files: [], contentTypes: [], caches: [], sheets: book.sheets.map(() => []) }
  const caches = book.pivotCaches ?? []
  if (!caches.length) return out
  const date1904 = book.dateSystem === 1904
  caches.forEach((c, k) => {
    const n = k + 1
    const name = `xl/pivotCache/pivotCacheDefinition${n}.xml`
    const saved = c.records !== undefined && c.saveData !== false
    const own: [string, string, boolean?][] = []
    if (saved) own.push([`${REL}/pivotCacheRecords`, `pivotCacheRecords${n}.xml`])
    // a source on a sheet the workbook no longer has: a link to a file by its path, as Excel writes it after deleting the sheet
    const src = c.source
    const gone = src.type === 'range' && src.book === undefined && !book.sheets.some((s) => sheetKey(s.name) === sheetKey(src.sheet))
    if (gone) issues.add('pivotTable', { key: 'pivotSourceLinked', sheet: src.sheet })
    const book_ = gone ? DELETED_SOURCE : 'book' in src ? src.book : undefined
    if (book_ !== undefined) own.push([`${REL}/externalLinkPath`, book_, true])
    out.files.push({ name, text: pivotCacheXml(c, (f) => styles.numFmt(f.numberFormat), saved ? 'rId1' : undefined, book_ === undefined ? undefined : `rId${own.length}`, slicerIds.has(c.id) ? c.id : undefined) })
    out.contentTypes.push([`/${name}`, `${TYPE}.pivotCacheDefinition+xml`])
    if (saved) {
      out.files.push({ name: `xl/pivotCache/pivotCacheRecords${n}.xml`, text: pivotRecordsXml(c) })
      out.contentTypes.push([`/xl/pivotCache/pivotCacheRecords${n}.xml`, `${TYPE}.pivotCacheRecords+xml`])
    }
    if (own.length) out.files.push({ name: `xl/pivotCache/_rels/pivotCacheDefinition${n}.xml.rels`, text: rels(own) })
    out.caches.push({ id: c.id, target: `pivotCache/pivotCacheDefinition${n}.xml` })
  })
  const part = new Map(caches.map((c, k) => [c.id, k + 1]))
  let t = 0
  book.sheets.forEach((s, i) => {
    for (const p of s.pivotTables ?? []) {
      const n = ++t
      const cache = caches.find((c) => c.id === p.cache)!
      const w: PivotWriting = { numFmt: (f) => styles.numFmt(f), dxf: (x) => styles.dxf(x), date1904, rules: pivotRules(s, p.name), locale: book.locale }
      out.files.push({ name: `xl/pivotTables/pivotTable${n}.xml`, text: pivotTableXml(p, cache, w) })
      out.files.push({ name: `xl/pivotTables/_rels/pivotTable${n}.xml.rels`, text: rels([[`${REL}/pivotCacheDefinition`, `../pivotCache/pivotCacheDefinition${part.get(p.cache)}.xml`]]) })
      out.contentTypes.push([`/xl/pivotTables/pivotTable${n}.xml`, `${TYPE}.pivotTable+xml`])
      out.sheets[i].push(`../pivotTables/pivotTable${n}.xml`)
    }
  })
  return out
}
