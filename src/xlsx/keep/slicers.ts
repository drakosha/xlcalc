/**
 * Slicers and timelines over an original package: caches, sheet lists and drawing frames kept while unchanged,
 * written through their existing parts when changed, added with new relationships, removed with their last owner.
 * Other drawing objects and extensions stay as the file has them.
 */
import type { BookInput, DrawingPlace, Slicer, SlicerCache, Timeline, TimelineCache } from '../../contract.js'
import { caseKey } from '../../collate.js'
import { isCacheNameValue } from '../../slicer-cache-name.js'
import { localeBooleans } from '../../locale/registry.js'
import { nameProblem } from '../../syntax/check.js'
import { sheetKey } from '../../syntax/print.js'
import { attrsAt } from '../xml/scan.js'
import { escapeAttr, readText } from '../xml/text.js'
import {
  frameXml,
  type FrameKind,
  SLICER_RELS,
  sheetSlicerExt,
  slicerCacheXml,
  slicersPartXml,
  tableSlicerCacheXml,
  timelineCacheXml,
  timelinesPartXml,
  workbookSlicerExt,
} from '../write/slicers.js'
import { HEAD } from '../write/xml.js'
import type { PivotFiles, PivotHost } from './pivots.js'
import { KeptRefusal } from './refusals.js'
import { relsPath } from './source.js'
import { attrsOf, childrenOf, opening, Rels, relType, rootOf, setAttrs } from './xml.js'

type SlicerObject = Slicer | Timeline
type Cache = SlicerCache | TimelineCache
const WB_EXT = ['{BBE1A952-AA13-448E-AADC-164F8A28A991}', '{46BE6895-7355-4A93-B00E-2C351335B9C9}', '{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}']
const SHEET_EXT = ['{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}', '{3A4CF648-6AED-40F4-86FF-DC5316D8AED3}', '{7E03D99C-DC04-49D9-9315-930204A7B6E9}']
const TYPE = 'application/vnd.ms-excel.'
const DRAWING = 'application/vnd.openxmlformats-officedocument.drawing+xml'
const tableCache = (c: Cache): c is SlicerCache & { table: NonNullable<SlicerCache['table']> } => 'table' in c && c.table !== undefined
const key = (c: { name: string }) => caseKey(c.name)
const caches = (b: BookInput): Cache[] => [...(b.slicerCaches ?? []), ...(b.timelineCaches ?? [])]
const objects = (s: BookInput['sheets'][number]): SlicerObject[] => [...(s.slicers ?? []), ...(s.timelines ?? [])]
const unplaced = (list: SlicerObject[]) => JSON.stringify(list.map((x) => ({ ...x, place: undefined })))

/** One root child replaced, removed, or inserted before the first of `before`, else at the end; unrelated text stays. */
function child(xml: string, name: string, text: string, before: string[] = []): string {
  const root = rootOf(xml)!
  const kids = childrenOf(xml, root)
  const old = kids.find((c) => c.local === name)
  if (old) return xml.slice(0, old.start) + text + xml.slice(old.end)
  if (!text) return xml
  if (root.end === root.open) return xml.slice(0, root.start) + opening(xml.slice(root.start, root.open)) + text + `</${root.qname}>` + xml.slice(root.end)
  const at = kids.find((c) => before.includes(c.local))?.start ?? xml.lastIndexOf('</', root.end)
  return xml.slice(0, at) + text + xml.slice(at)
}

/** Only the named extensions replaced; foreign extensions keep their bytes and order. */
function extensions(xml: string, uris: string[], text: string): string {
  const list = childrenOf(xml, rootOf(xml)!).find((c) => c.local === 'extLst')
  const held = list
    ? childrenOf(xml, list)
        .filter((c) => !uris.includes((attrsOf(xml, c).uri ?? '').toUpperCase()))
        .map((c) => xml.slice(c.start, c.end))
        .join('')
    : ''
  return child(xml, 'extLst', held || text ? `<extLst>${held}${text}</extLst>` : '')
}

export interface SlicerHost extends PivotHost {
  parts: PivotHost['parts'] & { text: Map<string, string> }
  /** The parts of the sheets written, in the book's order, and their ids in the file. */
  sheets: string[]
  tab(sheet: string): number
  sourceId(id: number): number | undefined
  /** The table and column ids actually written, including those of added tables. */
  tables: Map<string, { id: number; columns: Map<string, number> }>
}

export class KeptSlicers {
  private cacheRefs: { kind: 'slicerCache' | 'timelineCache'; id: string; table?: true }[] = []
  private cacheListChanged = false

  constructor(
    private readonly h: SlicerHost,
    private readonly files: PivotFiles,
  ) {}

  private text(part: string): string {
    return this.h.parts.text.get(part.toLowerCase()) ?? this.files.texts.get(part.toLowerCase()) ?? ''
  }

  private drop(part: string): void {
    this.h.parts.remove(part)
    this.h.parts.remove(relsPath(part))
    this.h.types.remove(part)
  }

  /** Caches written after tables and sheets, whose ids a new table slicer or PivotTable reference needs. */
  write(): void {
    const { book, book0, src, parts, wbRels } = this.h
    const before = caches(book0)
    const after = caches(book)
    const kept = new Set(after.map(key))
    const timelines = book.timelineCaches ?? []
    for (const [i, c] of timelines.entries())
      if (
        timelines
          .slice(0, i)
          .some(
            (p) =>
              p.cache === c.cache &&
              caseKey(p.field) === caseKey(c.field) &&
              p.pivots.some((r) => c.pivots.some((t) => sheetKey(r.sheet) === sheetKey(t.sheet) && caseKey(r.name) === caseKey(t.name))),
          )
      )
        throw new KeptRefusal({ key: 'keptSlicers' }, 'slicers')
    const tab0 = (name: string) => {
      const k = book0.sheets.findIndex((s) => sheetKey(s.name) === sheetKey(name))
      return k < 0 ? 0 : Number(src.decls[src.sheets[k].decl].attrs.sheetId)
    }
    const cacheKey = (c: Cache | undefined, tab: (s: string) => number) =>
      JSON.stringify(c && { ...c, table: undefined, pivots: c.pivots.map((p) => [tab(p.sheet), caseKey(p.name)]) })
    for (const kind of ['slicerCache', 'timelineCache'] as const) {
      const list = kind === 'slicerCache' ? (book.slicerCaches ?? []) : (book.timelineCaches ?? [])
      for (const c of list) {
        const c0 = before.find((x) => key(x) === key(c))
        const old = this.files.slicerCaches.get(key(c))
        const part = old ?? parts.fresh(`xl/${kind}s/${kind}`, '.xml')
        const table = tableCache(c) ? this.h.tables.get(caseKey(c.table.name)) : undefined
        const column = tableCache(c) ? table?.columns.get(caseKey(c.table.column)) : undefined
        const ids = table && column !== undefined ? { tableId: table.id, column } : undefined
        const oldIds = this.files.tableSlicers.get(key(c))
        const sameIds = !ids || (String(ids.tableId) === oldIds?.tableId && String(ids.column) === oldIds?.column)
        const name = book.names?.find((n) => n.scope === undefined && caseKey(n.name) === key(c))
        if (!old && (nameProblem(c.name, localeBooleans(book.locale)) !== null || (name && !isCacheNameValue(name.ref))))
          throw new KeptRefusal({ key: 'keptSlicers' }, 'slicers')
        if (!old || cacheKey(c0, tab0) !== cacheKey(c, this.h.tab) || !sameIds) {
          if ((tableCache(c) && !ids) || (!ids && c.cache === undefined)) throw new KeptRefusal({ key: 'keptSlicerSource', name: c.name }, 'slicers')
          const id = c.cache === undefined ? undefined : (this.h.sourceId(c.cache) ?? c.cache)
          parts.set(
            part,
            ids
              ? tableSlicerCacheXml(c as SlicerCache, ids)
              : kind === 'slicerCache'
                ? slicerCacheXml(book, c as SlicerCache, this.h.tab, id)
                : timelineCacheXml(book, c as TimelineCache, this.h.tab, id),
          )
          this.h.change()
        }
        const rel = src.wbRels.find((r) => !r.external && r.kind === kind && r.target === old)
        const id = rel?.id ?? wbRels.add(SLICER_RELS[kind], this.h.relative(src.wbPath, part))
        this.cacheRefs.push({ kind, id, ...(ids || oldIds ? { table: true as const } : {}) })
        if (!old) {
          this.h.types.override(part, `${TYPE}${kind}+xml`)
          this.cacheListChanged = true
        }
      }
    }
    for (const r of src.wbRels) {
      if (r.external || (r.kind !== 'slicerCache' && r.kind !== 'timelineCache')) continue
      const c0 = before.find((c) => this.files.slicerCaches.get(key(c)) === r.target)
      if (!c0)
        this.cacheRefs.push({
          kind: r.kind,
          id: r.id,
          ...(this.files.tableSlicers.has(caseKey(readText(attrsAt(this.text(r.target), 'slicerCacheDefinition').name ?? ''))) ? { table: true as const } : {}),
        })
      else if (!kept.has(key(c0))) {
        this.drop(r.target)
        wbRels.remove(r.id)
        this.cacheListChanged = true
        this.h.change()
      }
    }
    for (const [j, part] of this.h.sheets.entries()) this.sheet(j, part)
  }

  /** A sheet's separate lists for PivotTable slicers, table slicers and timelines, and their shared drawing. */
  private sheet(j: number, part: string): void {
    const { src, book, book0, parts } = this.h
    const k = this.h.fileSheet(book.sheets[j].name)
    const ss = k === undefined ? undefined : src.sheets[k]
    const s0 = k === undefined ? undefined : book0.sheets[k]
    const s1 = book.sheets[j]
    let xml = parts.text.get(part.toLowerCase()) ?? ss?.xml ?? ''
    const rels = new Rels(parts.text.get(relsPath(part).toLowerCase()) ?? ss?.relsXml)
    const table = (x: SlicerObject, b: BookInput) => !!b.slicerCaches?.find((c) => key(c) === caseKey(x.cache))?.table || this.files.tableSlicers.has(caseKey(x.cache))
    const lists = (s: BookInput['sheets'][number] | undefined, b: BookInput) => [
      s?.slicers?.filter((x) => !table(x, b)) ?? [],
      s?.slicers?.filter((x) => table(x, b)) ?? [],
      s?.timelines ?? [],
    ]
    const [oldLists, newLists] = [lists(s0, book0), lists(s1, book)]
    const list = childrenOf(xml, rootOf(xml)!).find((c) => c.local === 'extLst')
    const ids: (string | undefined)[] = SHEET_EXT.map((uri) => {
      const ext = list && childrenOf(xml, list).find((c) => (attrsOf(xml, c).uri ?? '').toUpperCase() === uri)
      return ext ? /\b(?:\w+:)?id="([^"]+)"/.exec(xml.slice(ext.start, ext.end))?.[1] : undefined
    })
    let listChanged = false
    for (let n = 0; n < 3; n++) {
      const [old, next] = [oldLists[n], newLists[n]]
      if (unplaced(old) === unplaced(next)) continue
      listChanged = true
      const rel = ss?.rels.find((r) => !r.external && r.id === ids[n])
      if (next.length) {
        const kind = n === 2 ? 'timeline' : 'slicer'
        const target = rel?.target ?? parts.fresh(`xl/${kind}s/${kind}`, '.xml')
        const text = n === 2 ? timelinesPartXml(next as Timeline[]) : slicersPartXml(next as Slicer[], n === 1 ? 257175 : 241300)
        parts.set(target, text)
        if (!rel) {
          this.h.types.override(target, `${TYPE}${kind}+xml`)
          ids[n] = rels.add(SLICER_RELS[kind], this.h.relative(part, target))
        }
      } else if (rel) {
        this.drop(rel.target)
        rels.remove(rel.id)
        ids[n] = undefined
      }
    }
    if (listChanged) xml = extensions(xml, SHEET_EXT, sheetSlicerExt(ids[0], ids[2], ids[1]))
    const old = s0 ? objects(s0) : []
    const next = objects(s1)
    const framesChanged = JSON.stringify(old) !== JSON.stringify(next)
    if (framesChanged) xml = this.frames(xml, part, rels, ss?.rels.find((r) => r.kind === 'drawing' && !r.external)?.target, old, next)
    if (listChanged || framesChanged) {
      parts.set(part, xml)
      if (rels.size) parts.set(relsPath(part), rels.xml())
      else parts.remove(relsPath(part))
      this.h.change()
    }
  }

  /** Named frames matched by name, then a unique rename over the same cache; untouched anchors keep their bytes. */
  private frames(xml: string, sheet: string, rels: Rels, original: string | undefined, old: SlicerObject[], next: SlicerObject[]): string {
    const { parts, book } = this.h
    const drawing = original && parts.has(original) ? original : undefined
    const part = drawing ?? parts.fresh('xl/drawings/drawing', '.xml')
    const file = drawing
      ? (parts.text.get(drawing.toLowerCase()) ?? this.h.edited(drawing, this.text(drawing)))
      : `${HEAD}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>`
    const root = rootOf(file)!
    const remaining = new Set(next)
    const take = (x0: SlicerObject) => {
      let x = next.find((y) => remaining.has(y) && key(y) === key(x0))
      if (!x) {
        const candidates = next.filter((y) => remaining.has(y) && caseKey(y.cache) === caseKey(x0.cache) && !old.some((o) => key(o) === key(y)))
        if (candidates.length === 1 && old.filter((o) => caseKey(o.cache) === caseKey(x0.cache) && !next.some((y) => key(y) === key(o))).length === 1) x = candidates[0]
      }
      if (x) remaining.delete(x)
      return x
    }
    const kind = (x: SlicerObject): FrameKind => {
      const c = book.slicerCaches?.find((c) => key(c) === caseKey(x.cache))
      return c ? c.table || this.files.tableSlicers.has(key(c)) ? 'tableSlicer' : 'slicer' : 'timeline'
    }
    const frame = (x: SlicerObject, id: number) => {
      if (!x.place) throw new KeptRefusal({ key: 'keptSlicers' }, 'slicers')
      return frameXml(x as SlicerObject & { place: DrawingPlace }, id, kind(x))
    }
    const anchors: string[] = []
    let id = Math.max(0, ...[...file.matchAll(/<(?:[\w.-]+:)?cNvPr\b[^>]*?\sid="(\d+)"/g)].map((m) => Number(m[1])))
    for (const anchor of childrenOf(file, root)) {
      const text = file.slice(anchor.start, anchor.end)
      const name = /<(?:[\w.-]+:)?(?:slicer|timeslicer)\b[^>]*?\sname="([^"]*)"/.exec(text)?.[1]
      const x0 = name === undefined ? undefined : old.find((x) => key(x) === caseKey(readText(name)))
      if (!x0) {
        anchors.push(text)
        continue
      }
      const x = take(x0)
      if (!x) continue
      const visible = (y: SlicerObject) => JSON.stringify([y.name, y.cache, y.place, y.locked !== false, 'positionLocked' in y && !!y.positionLocked])
      anchors.push(visible(x0) === visible(x) ? text : frame(x, Number(/<(?:[\w.-]+:)?cNvPr\b[^>]*?\sid="(\d+)"/.exec(text)?.[1] ?? ++id)))
    }
    for (const x of remaining) anchors.push(frame(x, ++id))
    const rel = this.h.src.sheets.flatMap((s) => (s.part === sheet ? s.rels : [])).find((r) => !r.external && r.kind === 'drawing')
    if (!anchors.length) {
      if (drawing) this.drop(drawing)
      if (rel) rels.remove(rel.id)
      return child(xml, 'drawing', '')
    }
    const declarations = attrsOf(file, root)
    let tag = opening(file.slice(root.start, root.open))
    if (!declarations['xmlns:xdr']) tag = setAttrs(tag, { 'xmlns:xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing' })
    if (!declarations['xmlns:a']) tag = setAttrs(tag, { 'xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main' })
    parts.set(part, `${file.slice(0, root.start)}${tag}${anchors.join('')}</${root.qname}>`)
    if (!drawing) {
      if (rel) rels.remove(rel.id)
      this.h.types.override(part, DRAWING)
      const id = rels.add(relType('drawing'), this.h.relative(sheet, part))
      xml = child(xml, 'drawing', `<drawing r:id="${id}"/>`, ['legacyDrawing', 'legacyDrawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'])
    }
    return xml
  }

  /** Cache extensions and their defined names, after the workbook's ordinary names have been written. */
  workbook(xml: string): string {
    if (!this.cacheListChanged) return xml
    xml = extensions(xml, WB_EXT, workbookSlicerExt(this.cacheRefs))
    const before = caches(this.h.book0)
    const after = caches(this.h.book)
    const list = childrenOf(xml, rootOf(xml)!).find((c) => c.local === 'definedNames')
    const held: string[] = []
    const seen = new Set<string>()
    for (const e of list ? childrenOf(xml, list) : []) {
      const a = attrsOf(xml, e)
      const name = caseKey(readText(a.name ?? ''))
      const text = xml.slice(e.start, e.end)
      const ref = xml.slice(e.open, xml.lastIndexOf('</', e.end - 1))
      if (a.localSheetId === undefined && before.some((c) => key(c) === name) && !after.some((c) => key(c) === name) && isCacheNameValue(ref)) continue
      if (a.localSheetId === undefined) seen.add(name)
      held.push(text)
    }
    for (const c of after) if (!seen.has(key(c))) held.push(`<definedName name="${escapeAttr(c.name)}">#N/A</definedName>`)
    return child(xml, 'definedNames', held.length ? `<definedNames>${held.join('')}</definedNames>` : '', [
      'calcPr',
      'oleSize',
      'customWorkbookViews',
      'pivotCaches',
      'smartTagPr',
      'smartTagTypes',
      'webPublishing',
      'fileRecoveryPr',
      'webPublishObjects',
      'extLst',
    ])
  }
}
