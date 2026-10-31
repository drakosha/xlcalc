/**
 * The workbook's PivotTables, caches, slicers and timelines, found through the relationships: the caches the workbook
 * lists (`<pivotCaches>`) with their records, each sheet's PivotTables, slicers and timelines with their places in its
 * drawing, and the workbook's caches of slicers and timelines. What the model holds of them only as stored (an external
 * source, the data model, consolidated ranges, OLAP hierarchies, chart formats) is reported.
 */
import { caseKey } from '../../collate.js'
import type { BookInput, NumberFormat, PivotCache, SheetInput } from '../../contract.js'
import type { DateSystem } from '../../dates.js'
import { InputError } from '../../input-error.js'
import { checkPivots } from '../../input/pivot.js'
import { checkPivotShapes } from '../../pivot/check.js'
import type { Issues } from '../issues.js'
import { numberFormatOf, type StyleTable } from '../styles/read.js'
import type { Dxf } from '../styles/dxf.js'
import { attr, attributes, attrsAt, elements, has, within } from '../xml/scan.js'
import type { Archive } from '../zip/read.js'
import { relsOf, type Rel } from './package.js'
import { readPivotCache } from './pivot-cache.js'
import { readPivotTable } from './pivot-table.js'
import { drawingPlaces, readSlicerCache, readSlicers, readTimelineCache, readTimelines } from './slicers.js'

/** The parts of the workbook's PivotTables, read once for its sheets: the caches by id, number formats, differential formats. */
export interface PivotBook {
  caches: PivotCache[]
  /** Excel 2010's id of each cache (`x14:pivotCacheDefinition pivotCacheId`), which slicers and timelines name. */
  x14: Map<string, number>
  format: (id: string | undefined) => NumberFormat | undefined
  dxfs: Dxf[]
  dateSystem: DateSystem
}

/** The ids of the connections that are the workbook's data model. */
async function modelConnections(archive: Archive, rels: Rel[]): Promise<Set<number>> {
  const r = rels.find((x) => x.kind === 'connections' && !x.external)
  const xml = r ? await archive.text(r.target) : undefined
  const out = new Set<number>()
  for (const e of elements(xml ?? '', 'connection')) {
    const a = attributes(e.attrs)
    if (a.name === 'ThisWorkbookDataModel' || a.model === '1') out.add(Number(a.id))
  }
  return out
}

/** The workbook's pivot caches with their records, by the ids `<pivotCaches>` gives them. */
export async function readPivotCaches(archive: Archive, wb: string, rels: Rel[], styles: StyleTable, dxfs: Dxf[], dateSystem: DateSystem, issues: Issues): Promise<PivotBook> {
  const byId = new Map(rels.map((r) => [r.id, r]))
  const format = (id: string | undefined) => (id === undefined ? undefined : numberFormatOf(styles.formats, Number(id)))
  const out: PivotBook = { caches: [], x14: new Map(), format, dxfs, dateSystem }
  const listed = within(wb, 'pivotCaches', 'pivotCache')
  if (!listed.length) return out
  const model = await modelConnections(archive, rels)
  for (const e of listed) {
    const a = attributes(e.attrs)
    const r = byId.get(attr(a, 'id') ?? '')
    const xml = r && !r.external ? await archive.text(r.target) : undefined
    if (xml === undefined) {
      issues.add('pivotTable', { key: 'pivotCacheMissing' })
      continue
    }
    const own = await relsOf(archive, r!.target)
    const records = own.find((x) => x.kind === 'pivotCacheRecords' && !x.external)
    const linked = new Map(own.map((x) => [x.id, x.target]))
    const id = Number(a.cacheId)
    const cache = readPivotCache(id, xml, records ? await archive.text(records.target) : undefined, { format, book: (rid) => (rid === undefined ? undefined : linked.get(rid)), model: (c) => model.has(c) })
    const held = cache.source.type
    if (held === 'external' || held === 'dataModel' || held === 'consolidation' || held === 'scenario') issues.add('pivotTable', { key: 'pivotSourceHeld', source: held })
    if (has(xml, 'cacheHierarchies')) issues.add('pivotTable', { key: 'pivotOlap' })
    const ext = /<[\w.-]+:pivotCacheDefinition\s[^>]*?pivotCacheId="(\d+)"/.exec(xml)
    if (ext) out.x14.set(ext[1], id)
    out.caches.push(cache)
  }
  return out
}

/** The texts of a sheet's PivotTable parts, undefined for one missing. */
export const pivotTableParts = (archive: Archive, sheetRels: Rel[]): Promise<(string | undefined)[]> => Promise.all(sheetRels.filter((x) => x.kind === 'pivotTable' && !x.external).map((r) => archive.text(r.target)))

/** A sheet's PivotTables (from the texts of its parts), slicers and timelines, set on `sheet`. */
export async function readSheetPivots(archive: Archive, sheet: SheetInput, sheetRels: Rel[], parts: (string | undefined)[], drawing: string | undefined, book: PivotBook, issues: Issues) {
  const where = { sheet: sheet.name }
  for (const xml of parts) {
    const id = Number(attrsAt(xml ?? '', 'pivotTableDefinition').cacheId)
    if (xml === undefined || !book.caches.some((c) => c.id === id)) {
      issues.add('pivotTable', { key: 'pivotCacheMissing' }, where)
      continue
    }
    const { table, charts } = readPivotTable(xml, id, book)
    if (charts) issues.add('pivotTable', { key: 'pivotChartFormats' }, { ...where, name: table.name })
    ;(sheet.pivotTables ??= []).push(table)
  }
  const slicers = sheetRels.filter((x) => (x.kind === 'slicer' || x.kind === 'timeline') && !x.external)
  if (!slicers.length) return
  const places = drawingPlaces(drawing)
  for (const r of slicers) {
    const xml = (await archive.text(r.target)) ?? ''
    if (r.kind === 'slicer') (sheet.slicers ??= []).push(...readSlicers(xml, places))
    else (sheet.timelines ??= []).push(...readTimelines(xml, places))
  }
}

/** The workbook's caches of slicers and timelines, by its relationships; their PivotTables by the sheets' ids. */
export async function readSlicerCaches(archive: Archive, wb: string, rels: Rel[], pivots: PivotBook, book: BookInput, tables: Map<string, string>, issues: Issues) {
  const ids = new Map(within(wb, 'sheets', 'sheet').map((e) => [attributes(e.attrs).sheetId, attributes(e.attrs).name]))
  const named = new Map(book.sheets.map((s) => [caseKey(s.name), s.name]))
  const ctx = {
    sheetOf: (id: string | undefined) => named.get(caseKey(ids.get(id ?? '') ?? '')),
    cacheOf: (id: string | undefined) => (id === undefined ? undefined : pivots.x14.get(id)),
    tableOf: (id: string | undefined) => (id === undefined ? undefined : tables.get(id)),
  }
  for (const r of rels.filter((x) => (x.kind === 'slicerCache' || x.kind === 'timelineCache') && !x.external)) {
    const xml = await archive.text(r.target)
    if (xml === undefined) continue
    if (r.kind === 'slicerCache') (book.slicerCaches ??= []).push(readSlicerCache(xml, ctx))
    else {
      const c = readTimelineCache(xml, ctx)
      // a cache Excel 2010's id does not find takes its PivotTables' cache
      if (c.cache < 0) c.cache = book.sheets.find((s) => caseKey(s.name) === caseKey(c.pivots[0]?.sheet ?? ''))?.pivotTables?.find((t) => caseKey(t.name) === caseKey(c.pivots[0].name))?.cache ?? -1
      if (c.cache < 0) issues.add('pivotTable', { key: 'pivotCacheMissing' }, { name: c.name })
      else (book.timelineCaches ??= []).push(c)
    }
  }
}

/** The book's PivotTables and the rest as the model takes them, or all left out and reported when it cannot (a file's ids or names that clash). */
export function checkedPivots(book: BookInput, issues: Issues) {
  try {
    checkPivotShapes(book)
    checkPivots(book as unknown as Record<string, unknown>, 'book')
  } catch (e) {
    if (!(e instanceof InputError)) throw e
    issues.add('pivotTable', { key: 'pivotUnread', detail: e.message })
    delete book.pivotCaches
    delete book.slicerCaches
    delete book.timelineCaches
    for (const s of book.sheets) {
      delete s.pivotTables
      delete s.slicers
      delete s.timelines
      const rules = s.conditionalFormats?.filter((c) => !c.pivot)
      if (rules?.length) s.conditionalFormats = rules
      else delete s.conditionalFormats
    }
  }
}
