/**
 * PivotTables and their caches over the file (`writeKept`): a PivotTable or a cache the book did not change stays as the
 * file has it, byte for byte (a PivotTable moved with its cells has its place written back by its carrier); one the book
 * changed or refreshed is written again from the model as `writeXlsx` writes it, through the file's relationships; a
 * new one gets parts of its own, and one the book deleted goes with its parts and its entry in the workbook.
 */
import type { BookInput, PivotCache, PivotTable, SlicerCache, TimelineCache } from '../../contract.js'
import { caseKey } from '../../collate.js'
import type { Rel } from '../read/package.js'
import { relsOf } from '../read/package.js'
import { pivotRules } from '../write/pivot-parts.js'
import { pivotCacheXml, pivotRecordsXml } from '../write/pivot-cache.js'
import { pivotTableXml, type PivotWriting } from '../write/pivot.js'
import { slicerCacheXml, slicersPartXml, timelineCacheXml, timelinesPartXml } from '../write/slicers.js'
import { KeptRefusal } from './refusals.js'
import { sheetKey } from '../../syntax/print.js'
import type { CellFormats } from '../write/styles.js'
import { attr, attributes, element } from '../xml/scan.js'
import { readText } from '../xml/text.js'
import { relsPath, type Source, type SourceSheet } from './source.js'
import { attrsOf, childrenOf, Rels, relType, rootOf, TYPE, type ContentTypes, type Placed } from './xml.js'
import type { Edit } from './carriers.js'


/** A cache of the file: its id, its workbook relationship and part, and the part of its rows. */
interface FileCache {
  id: number
  rel: string
  part: string
  xml: string
  rels: Rel[]
  relsXml?: string
}

/** The file's PivotTables' parts by name, its caches by id, its slicers' and timelines' caches by name with their parts. */
export interface PivotFiles {
  tables: Map<string, string>
  caches: FileCache[]
  slicerCaches: Map<string, string>
}

export async function pivotFiles(src: Source): Promise<PivotFiles> {
  const tables = new Map<string, string>()
  for (const s of src.sheets) for (const r of s.rels) if (r.kind === 'pivotTable' && !r.external) tables.set(r.target, (await src.archive.text(r.target)) ?? '')
  const caches: FileCache[] = []
  const list = childrenOf(src.wb, rootOf(src.wb)!).find((c) => c.local === 'pivotCaches')
  for (const c of list ? childrenOf(src.wb, list).filter((x) => x.local === 'pivotCache') : []) {
    const a = attrsOf(src.wb, c)
    const rel = attr(a, 'id') ?? ''
    const r = src.wbRels.find((x) => x.id === rel && !x.external)
    if (r) caches.push({ id: Number(a.cacheId), rel, part: r.target, xml: (await src.archive.text(r.target)) ?? '', rels: await relsOf(src.archive, r.target), relsXml: await src.archive.text(relsPath(r.target)) })
  }
  const slicerCaches = new Map<string, string>()
  for (const r of src.wbRels)
    if ((r.kind === 'slicerCache' || r.kind === 'timelineCache') && !r.external) {
      const xml = (await src.archive.text(r.target)) ?? ''
      const name = /<(?:[\w.-]+:)?(?:slicer|timeline)CacheDefinition\b[^>]*?\sname="([^"]*)"/.exec(xml)?.[1]
      if (name !== undefined) slicerCaches.set(caseKey(readText(name)), r.target)
    }
  return { tables, caches, slicerCaches }
}

/** What the writing of the book over the file gives the PivotTables. */
export interface PivotHost {
  src: Source
  book0: BookInput
  book: BookInput
  parts: { has(part: string): boolean; set(part: string, text: string): void; remove(part: string): void; fresh(prefix: string, ext: string): string }
  types: ContentTypes
  wbRels: Rels
  /** The book's styles as the file writes them, and the differential formats a part uses marked used. */
  styles: CellFormats
  useDxf(i: number): void
  /** The part's text with the moved references written back. */
  edited(part: string, text: string): string
  relative(from: string, to: string): string
  change(): void
  /** The file's sheet of a sheet of the book, by its name. */
  fileSheet(name: string): number | undefined
}

/** A PivotTable without its place, which moves with its cells (written back by its carrier). */
const unplaced = (p: PivotTable) => JSON.stringify({ ...p, location: undefined })

/** A cache without its source, which its carrier writes back where it moves with its cells. */
const unsourced = (c: PivotCache | undefined) => JSON.stringify(c && { ...c, source: undefined })

/** A cache's source as its part gives it, its references written back: a range by sheet and reference, a table or name by its name. */
function fileSource(xml: string): string {
  const a = attributes(element(xml, 'worksheetSource')?.attrs ?? '')
  return a.name !== undefined ? `n${readText(a.name)}` : `r${readText(a.sheet ?? '')}!${a.ref ?? ''}`
}

const modelSource = (c: PivotCache) => (c.source.type === 'range' ? `r${c.source.sheet}!${c.source.range}` : c.source.type === 'table' ? `n${c.source.table}` : c.source.type === 'name' ? `n${c.source.name}` : '')

export class KeptPivots {
  /** The part of each cache of the book written, by id. */
  private readonly part = new Map<number, string>()
  /** The file's id of each cache among its slicers' sources (x14 `pivotCacheId`), by id. */
  private readonly x14 = new Map<number, number>()

  constructor(
    private readonly h: PivotHost,
    private readonly files: PivotFiles,
  ) {
    for (const f of files.caches) {
      const id = /<(?:[\w.-]+:)?pivotCacheDefinition\b[^>]*?\spivotCacheId="(\d+)"/.exec(f.xml.slice(f.xml.indexOf('<extLst')))?.[1]
      if (id !== undefined && f.xml.includes('<extLst')) this.x14.set(f.id, Number(id))
    }
  }

  /** A cache's id among the slicers' sources: the file's, else its own where a slicer or a timeline filters it. */
  private sourceId(id: number): number | undefined {
    const { book } = this.h
    return this.x14.get(id) ?? ([...(book.slicerCaches ?? []), ...(book.timelineCaches ?? [])].some((c) => c.cache === id) ? id : undefined)
  }

  private writing(s: number, p: PivotTable): PivotWriting {
    const { book } = this.h
    return { numFmt: (f) => this.h.styles.numFmt(f), dxf: (x) => this.h.styles.dxf(x), date1904: book.dateSystem === 1904, rules: pivotRules(book.sheets[s], p.name), locale: book.locale }
  }

  /** The caches: each the book changed written again, a new one added, one the book no longer has removed with its rows. */
  caches(): void {
    const { book0, book, parts, types } = this.h
    const was = new Map((book0.pivotCaches ?? []).map((c) => [c.id, c]))
    for (const c of book.pivotCaches ?? []) {
      const f = this.files.caches.find((x) => x.id === c.id)
      const c0 = was.get(c.id)
      // the same but for a source its carrier moved
      if (f && c0 && unsourced(c0) === unsourced(c) && (JSON.stringify(c0.source) === JSON.stringify(c.source) || fileSource(this.h.edited(f.part, f.xml)) === modelSource(c))) {
        this.part.set(c.id, f.part)
        continue
      }
      this.h.change()
      const part = f?.part ?? parts.fresh('xl/pivotCache/pivotCacheDefinition', '.xml')
      this.part.set(c.id, part)
      if (!f) types.override(part, `${TYPE}.pivotCacheDefinition+xml`)
      this.cache(c, part, f)
    }
    const kept = new Set((book.pivotCaches ?? []).map((c) => c.id))
    for (const f of this.files.caches) {
      if (kept.has(f.id) || !was.has(f.id)) continue
      this.h.change()
      for (const r of f.rels) if (r.kind === 'pivotCacheRecords' && !r.external) this.drop(r.target)
      this.drop(f.part)
      this.h.wbRels.remove(f.rel)
    }
  }

  /** A cache's definition written over `part` with its rows, through the file's relationships where it has them. */
  private cache(c: PivotCache, part: string, f: FileCache | undefined): void {
    const { parts, types } = this.h
    const rels = new Rels(f?.relsXml)
    const old = f?.rels.find((r) => r.kind === 'pivotCacheRecords' && !r.external)
    let records: string | undefined
    if (c.records !== undefined && c.saveData !== false) {
      const rp = old?.target ?? parts.fresh('xl/pivotCache/pivotCacheRecords', '.xml')
      parts.set(rp, pivotRecordsXml(c))
      if (!old) types.override(rp, `${TYPE}.pivotCacheRecords+xml`)
      records = old?.id ?? rels.add(relType('pivotCacheRecords'), this.h.relative(part, rp))
    } else if (old) {
      this.drop(old.target)
      rels.remove(old.id)
    }
    const link = 'book' in c.source && c.source.book !== undefined ? (f?.rels.find((r) => r.kind === 'externalLinkPath')?.id ?? rels.add(relType('externalLinkPath'), c.source.book, true)) : undefined
    parts.set(part, pivotCacheXml(c, (x) => this.h.styles.numFmt(x.numberFormat), records, link, this.sourceId(c.id)))
    if (rels.size) parts.set(relsPath(part), rels.xml())
    else parts.remove(relsPath(part))
  }

  /** A part and its relationships removed. */
  private drop(part: string): void {
    this.h.parts.remove(part)
    this.h.types.remove(part)
    this.h.parts.remove(relsPath(part))
  }

  /** The PivotTables of the file's sheet `k` (the book's `j`): each the book changed written again, a new one added, one gone removed. */
  sheet(ss: SourceSheet, k: number, j: number, rels: Rels): void {
    const { book0, book } = this.h
    const before = book0.sheets[k].pivotTables ?? []
    const left = [...(book.sheets[j].pivotTables ?? [])]
    for (const r of ss.rels) {
      if (r.kind !== 'pivotTable' || r.external) continue
      const xml = this.files.tables.get(r.target) ?? ''
      const name = readText(attributes(element(xml, 'pivotTableDefinition')?.attrs ?? '').name ?? '')
      const p0 = before.find((p) => caseKey(p.name) === caseKey(name))
      // a PivotTable the model left out stays as the file has it
      if (!p0) continue
      // by its name, else renamed: the one at its place that no PivotTable of the file has
      let i = left.findIndex((p) => caseKey(p.name) === caseKey(name))
      if (i < 0) i = left.findIndex((p) => p.location.range === p0.location.range && !before.some((x) => caseKey(x.name) === caseKey(p.name)))
      if (i < 0) {
        this.h.change()
        this.drop(r.target)
        rels.remove(r.id)
        continue
      }
      const p1 = left.splice(i, 1)[0]
      if (unplaced(p0) === unplaced(p1) && p0.cache === p1.cache && unsourced(book0.pivotCaches?.find((c) => c.id === p0.cache)) === unsourced(book.pivotCaches?.find((c) => c.id === p1.cache))) continue
      this.h.change()
      this.table(p1, j, r.target)
    }
    for (const p of left) {
      this.h.change()
      this.added(p, j, ss.part, rels)
    }
  }

  /** A PivotTable the file does not have, with parts of its own named in the sheet's relationships. */
  private added(p: PivotTable, j: number, sheetPart: string, rels: Rels): void {
    const part = this.h.parts.fresh('xl/pivotTables/pivotTable', '.xml')
    this.h.types.override(part, `${TYPE}.pivotTable+xml`)
    this.table(p, j, part)
    rels.add(relType('pivotTable'), this.h.relative(sheetPart, part))
  }

  /** A PivotTable written over `part`, its relationship pointing at its cache's part. */
  private table(p: PivotTable, j: number, part: string): void {
    const cache = this.h.book.pivotCaches!.find((c) => c.id === p.cache)!
    const xml = pivotTableXml(p, cache, this.writing(j, p))
    for (const m of xml.matchAll(/ dxfId="(\d+)"/g)) this.h.useDxf(Number(m[1]))
    this.h.parts.set(part, xml)
    const rels = new Rels(undefined)
    rels.add(relType('pivotCacheDefinition'), this.h.relative(part, this.part.get(p.cache)!))
    this.h.parts.set(relsPath(part), rels.xml())
  }

  /** The PivotTables of a sheet the file does not have: each with parts of its own. */
  newSheet(sheetPart: string, j: number, rels: Rels): void {
    for (const p of this.h.book.sheets[j].pivotTables ?? []) this.added(p, j, sheetPart, rels)
  }

  /** The workbook's `<pivotCaches>`: the entries of the caches the model did not read as the file has them, then the book's caches; as it was while the book has the file's caches. */
  workbook(wb: string, kids: Placed[], edits: Edit[], insertAt: (before: string[]) => number): void {
    const { book0, book, src } = this.h
    const ids = (b: BookInput) => (b.pivotCaches ?? []).map((c) => c.id).join()
    if (ids(book0) === ids(book)) return
    const list = kids.find((c) => c.local === 'pivotCaches')
    const read = new Set((book0.pivotCaches ?? []).map((c) => c.id))
    const own = list ? childrenOf(wb, list).filter((c) => c.local === 'pivotCache' && !read.has(Number(attrsOf(wb, c).cacheId))).map((c) => wb.slice(c.start, c.end)) : []
    const items = (book.pivotCaches ?? []).map((c) => {
      const f = this.files.caches.find((x) => x.id === c.id)
      return `<pivotCache cacheId="${c.id}" r:id="${f?.rel ?? this.h.wbRels.add(relType('pivotCacheDefinition'), this.h.relative(src.wbPath, this.part.get(c.id)!))}"/>`
    })
    const all = [...own, ...items]
    const text = all.length ? `<pivotCaches>${all.join('')}</pivotCaches>` : ''
    this.h.change()
    if (list) edits.push({ part: src.wbPath, start: list.start, end: list.end, text })
    else {
      const at = insertAt(['sheets', 'functionGroups', 'externalReferences', 'definedNames', 'calcPr', 'oleSize', 'customWorkbookViews'])
      edits.push({ part: src.wbPath, start: at, end: at, text })
    }
  }

  /**
   * The slicers and timelines: a cache the book changed (the items or dates chosen, the PivotTables it filters) written
   * again through its part, a sheet's slicers or timelines whose fields changed their part written again; one added or
   * deleted is refused (`keptSlicers`). Their frames move with their cells by their carriers.
   */
  slicers(): void {
    const { book0, book, src } = this.h
    const names = (b: BookInput) => JSON.stringify([...(b.slicerCaches ?? []), ...(b.timelineCaches ?? [])].map((c) => caseKey(c.name)).sort())
    const own = (b: BookInput) => JSON.stringify(b.sheets.flatMap((s) => [...(s.slicers ?? []), ...(s.timelines ?? [])].map((x) => caseKey(x.name))).sort())
    if (names(book0) !== names(book) || own(book0) !== own(book)) throw new KeptRefusal({ key: 'keptSlicers' }, 'slicers')
    // a PivotTable's sheet by its `sheetId` in the file, which a renamed sheet keeps
    const tabOf = (k: number | undefined) => (k === undefined ? 0 : Number(src.decls[src.sheets[k].decl].attrs.sheetId ?? 0))
    const tab = (sheet: string) => tabOf(this.h.fileSheet(sheet))
    const tab0 = (sheet: string) => tabOf(book0.sheets.findIndex((x) => sheetKey(x.name) === sheetKey(sheet)))
    const key = (c: { pivots: { sheet: string; name: string }[] } | undefined, t: (sheet: string) => number) => JSON.stringify(c && { ...c, pivots: c.pivots.map((r) => [t(r.sheet), caseKey(r.name)]) })
    for (const kind of ['slicerCaches', 'timelineCaches'] as const)
      for (const c of book[kind] ?? []) {
        const c0 = (book0[kind] ?? []).find((x) => caseKey(x.name) === caseKey(c.name))
        const part = this.files.slicerCaches.get(caseKey(c.name))
        if (!part || key(c0, tab0) === key(c, tab)) continue
        this.h.change()
        const id = this.sourceId(c.cache!) ?? c.cache!
        this.h.parts.set(part, kind === 'slicerCaches' ? slicerCacheXml(book, c as SlicerCache, tab, id) : timelineCacheXml(book, c as TimelineCache, tab, id))
      }
    // a sheet's part where a slicer or timeline changed but for its place, which its frame's carrier writes back
    const unplaced = (list: { place?: unknown }[] | undefined) => JSON.stringify((list ?? []).map((x) => ({ ...x, place: undefined })))
    for (const [k, ss] of src.sheets.entries()) {
      const [s0, s1] = [book0.sheets[k], book.sheets.find((x) => sheetKey(x.name) === sheetKey(book0.sheets[k].name))]
      if (!s1) continue
      for (const [kind, rel, xml] of [['slicers', 'slicer', slicersPartXml], ['timelines', 'timeline', timelinesPartXml]] as const) {
        if (unplaced(s0[kind]) === unplaced(s1[kind])) continue
        const r = ss.rels.find((x) => x.kind === rel && !x.external)
        if (!r) continue
        this.h.change()
        this.h.parts.set(r.target, xml(s1[kind] as never))
      }
    }
  }
}
