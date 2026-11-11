/**
 * A workbook's slicers and timelines over PivotTables and tables as parts of a package, as Excel 365 writes them: each
 * slicer's cache (x14 `slicerCacheDefinition`, a table's with its table's and column's ids) and timeline's cache (x15
 * `timelineCacheDefinition`), named by the workbook's extensions; each sheet's slicers of PivotTables, of tables and its
 * timelines in a part of each kind, named by the sheet's extensions, and their frames in the sheet's drawing, at their
 * places; each cache's name for the workbook's defined names. A slicer without a place is not written (reported), nor one
 * whose cache's name the workbook holds for anything else, or whose table or column is not in the workbook, nor a cache
 * with nothing left to show it.
 */
import type { BookInput, DrawingPlace, Slicer, SlicerCache, Timeline, TimelineCache, TimelineLevel } from '../../contract.js'
import { caseKey } from '../../collate.js'
import { localeBooleans } from '../../locale/registry.js'
import { nameProblem } from '../../syntax/check.js'
import { sheetKey } from '../../syntax/print.js'
import { parseAddress } from '../../model.js'
import type { Issues } from '../issues.js'
import { escapeAttr } from '../xml/text.js'
import { counted, give, HEAD, MAIN, X14, X15 } from './xml.js'

const MS = 'http://schemas.microsoft.com/office'
export const SLICER_RELS = { slicerCache: `${MS}/2007/relationships/slicerCache`, slicer: `${MS}/2007/relationships/slicer`, timelineCache: `${MS}/2011/relationships/timelineCache`, timeline: `${MS}/2011/relationships/timeline` }
const TYPES = { slicerCache: 'application/vnd.ms-excel.slicerCache+xml', slicer: 'application/vnd.ms-excel.slicer+xml', timelineCache: 'application/vnd.ms-excel.timelineCache+xml', timeline: 'application/vnd.ms-excel.timeline+xml' }
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
/** What a part of Excel 2010's or 2013's namespace declares beside it: SpreadsheetML as `x`, which older readers may ignore. */
const IGNORABLE = ` xmlns:mc="${MC}" mc:Ignorable="x" xmlns:x="${MAIN}"`
const DRAWING_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml'
const LEVEL: Record<TimelineLevel, number> = { years: 0, quarters: 1, months: 2, days: 3 }
/** A slicer's buttons' height as Excel makes one, in EMU: over a PivotTable, over a table (the set `table-slicer`). */
const ROW_HEIGHT = 241300
const TABLE_ROW_HEIGHT = 257175

/** What the slicers and timelines add to a package. */
export interface SlicerPackage {
  files: { name: string; text: string }[]
  contentTypes: [string, string][]
  /** The workbook's relationships to the caches, in order: slicer caches, then timeline caches; `table`, a table's slicer cache. */
  caches: { kind: 'slicerCache' | 'timelineCache'; target: string; table?: true }[]
  /** Each sheet's parts, by the relationships it adds: its drawing, its slicers' part, its tables' slicers' part, its timelines' part. */
  sheets: { drawing?: string; slicers?: string; tableSlicers?: string; timelines?: string }[]
  /** The pivot caches a slicer or a timeline filters, by id: their `pivotCacheId`. */
  pivotCacheIds: Set<number>
  /** The caches written, by name: Excel refuses a file whose framed slicer's cache no defined name holds (the set `links/slicer-names`). */
  cacheNames: string[]
}

/** The value of the name of a cache Excel writes: a name of the workbook of any other costs the file. */
const NO_VALUE = /^=?\s*#N\/A\s*$/i

/** A sheet's `sheetId` by its name: as `writeXlsx` numbers them, by place, or the file's. */
export type TabId = (sheet: string) => number

const tabOf = (book: BookInput): TabId => (sheet) => book.sheets.findIndex((s) => sheetKey(s.name) === sheetKey(sheet)) + 1

/** The PivotTables a cache filters, by their sheet's tab (`tabId`) and name. */
function pivotTables(tab: TabId, c: { pivots: { sheet: string; name: string }[] }): string {
  const tabs = c.pivots.map((r) => `<pivotTable tabId="${tab(r.sheet)}" name="${escapeAttr(r.name)}"/>`)
  return tabs.length ? `<pivotTables>${tabs.join('')}</pivotTables>` : '<pivotTables/>'
}

/** A slicer cache's opening tag. */
const cacheHead = (c: SlicerCache) => `${HEAD}<slicerCacheDefinition xmlns="${X14}"${IGNORABLE} name="${escapeAttr(c.name)}" sourceName="${escapeAttr(c.field)}">`

/** The attributes of a cache's order and display of its items, of those `spec` names. */
function displayOf(c: SlicerCache, spec: string): string {
  const nd = c.noData ?? {}
  return `${c.sort === 'descending' ? ' sortOrder="descending"' : ''}${give(c, spec)}${nd.mark === false ? ' crossFilter="none"' : nd.last === false ? ' crossFilter="showItemsWithNoData"' : ''}`
}

/** The extension of a cache that hides its items with no data. */
const hideExt = (c: SlicerCache) => (c.noData?.hide ? `<x:ext uri="{470722E0-AACD-4C17-9CDC-17EF765DBC7E}" xmlns:x15="${X15}"><x15:slicerCacheHideItemsWithNoData/></x:ext>` : '')

export function slicerCacheXml(book: BookInput, c: SlicerCache, tab: TabId = tabOf(book), cacheId = c.cache!): string {
  // Excel opens no file whose slicer has no items (the set `links/file-refusals`): those of its field, each chosen
  const field = book.pivotCaches?.find((x) => x.id === c.cache)?.fields.find((f) => caseKey(f.name) === caseKey(c.field))
  const given = c.items?.length ? c.items : (field?.items ?? []).map((_, item) => ({ item }))
  const items = given.map((x: NonNullable<SlicerCache['items']>[number]) => `<i x="${x.item}"${x.selected !== false ? ' s="1"' : ''}${x.noData ? ' nd="1"' : ''}/>`)
  const hide = hideExt(c)
  return `${cacheHead(c)}${pivotTables(tab, c)}<data><tabular pivotCacheId="${cacheId}"${displayOf(c, 'customListSort:t showMissing:t')}>${counted('items', items)}</tabular></data>${hide && `<extLst>${hide}</extLst>`}</slicerCacheDefinition>`
}

/** A table's slicer's cache, as Excel writes it: the ids of its table and its column in the file, which a rename leaves. */
export function tableSlicerCacheXml(c: SlicerCache, ids: { tableId: string | number; column: string | number }): string {
  return `${cacheHead(c)}<extLst><x:ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" xmlns:x15="${X15}"><x15:tableSlicerCache tableId="${ids.tableId}" column="${ids.column}"${displayOf(c, 'customListSort:t')}/></x:ext>${hideExt(c)}</extLst></slicerCacheDefinition>`
}

/** A date of a timeline, ISO 8601 as the file writes it: with its time. */
const isoTime = (d: string) => (d.length === 10 ? `${d}T00:00:00` : d)

/** January 1 of a date's year, or of the next year's. */
const yearOf = (d: string, next: boolean) => `${Number(d.slice(0, 4)) + (next ? 1 : 0)}-01-01T00:00:00`

export function timelineCacheXml(book: BookInput, c: TimelineCache, tab: TabId = tabOf(book), cacheId = c.cache): string {
  const field = book.pivotCaches?.find((x) => x.id === c.cache)?.fields.find((f) => caseKey(f.name) === caseKey(c.field))
  const min = typeof field?.facts?.min === 'object' ? field.facts.min.date : undefined
  const max = typeof field?.facts?.max === 'object' ? field.facts.max.date : undefined
  // the dates it spans: the model's, else from January 1 of its first date's year to January 1 after its last's
  const bounds = c.bounds ?? (min && max ? { start: yearOf(min, false), end: yearOf(max, true) } : undefined)
  const sel = c.selection
  // Excel opens no file whose timeline without a selection lacks a filter type
  const type = ` filterType="${!sel ? 'unknown' : isoTime(sel.start) === isoTime(sel.end) ? 'dateEqual' : 'dateBetween'}"`
  const state = `<state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="${cacheId}"${type}>${sel ? `<selection startDate="${isoTime(sel.start)}" endDate="${isoTime(sel.end)}"/>` : ''}${bounds ? `<bounds startDate="${isoTime(bounds.start)}" endDate="${isoTime(bounds.end)}"/>` : ''}</state>`
  return `${HEAD}<timelineCacheDefinition xmlns="${X15}"${IGNORABLE} name="${escapeAttr(c.name)}" sourceName="${escapeAttr(c.field)}">${pivotTables(tab, c)}${state}</timelineCacheDefinition>`
}

/** A slicer's or a timeline's style where it has one: none for an empty name. */
const styleOf = (style: string | undefined) => (style ? ` style="${escapeAttr(style)}"` : '')

export function slicerXml(s: Slicer, rowHeight = ROW_HEIGHT): string {
  return `<slicer${give({ ...s, caption: s.caption ?? s.name }, 'name cache caption startItem:p columnCount=columns:p showCaption:t')}${styleOf(s.style)}${give(s, 'lockedPosition=positionLocked:f')} rowHeight="${s.rowHeight ?? rowHeight}"/>`
}

export function timelineXml(t: Timeline): string {
  const a =
    give({ ...t, caption: t.caption ?? t.name }, 'name cache caption showHeader:t showSelectionLabel:t showTimeLevel=showLevel:t showHorizontalScrollbar=showScrollbar:f') +
    ` level="${LEVEL[t.level ?? 'months']}" selectionLevel="${LEVEL[t.selectionLevel ?? 'years']}"` +
    (t.scrollPosition ? ` scrollPosition="${isoTime(t.scrollPosition)}"` : '') +
    styleOf(t.style)
  return `<timeline${a}/>`
}

/** A sheet's slicers' part, and its timelines'. */
export const slicersPartXml = (list: Slicer[], rowHeight = ROW_HEIGHT) => `${HEAD}<slicers xmlns="${X14}"${IGNORABLE}>${list.map((x) => slicerXml(x, rowHeight)).join('')}</slicers>`
export const timelinesPartXml = (list: Timeline[]) => `${HEAD}<timelines xmlns="${X15}"${IGNORABLE}>${list.map(timelineXml).join('')}</timelines>`

/** A cell of an anchor with its offset. */
const marker = (tag: string, cell: string, off: [number, number] | undefined) => {
  const { row, col } = parseAddress(cell)
  return `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>${off?.[0] ?? 0}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${off?.[1] ?? 0}</xdr:rowOff></xdr:${tag}>`
}

const DRAWING = `${MS}/drawing`
const SLICER = `${DRAWING}/2010/slicer`
const TIMESLICER = `${DRAWING}/2012/timeslicer`

/** What a frame shows: a table's slicer asks a later reader of Excel 2013's drawings than a PivotTable's does. */
export type FrameKind = 'slicer' | 'tableSlicer' | 'timeline'

/** A slicer's or timeline's frame in a drawing, at its place, its locks as Excel writes them, with the shape older readers show. */
export function frameXml(x: (Slicer | Timeline) & { place: DrawingPlace }, id: number, kind: FrameKind): string {
  const { name, place } = x
  const [requires, ns, uri, el] =
    kind === 'timeline'
      ? ['tsle', TIMESLICER, TIMESLICER, 'tsle:timeslicer']
      : kind === 'tableSlicer'
        ? ['sle15', `${DRAWING}/2012/slicer`, SLICER, 'sle:slicer']
        : ['a14', `${DRAWING}/2010/main`, SLICER, 'sle:slicer']
  const sle = kind === 'timeline' ? '' : ` xmlns:sle="${SLICER}"`
  const frame = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${escapeAttr(name)}"/>${'positionLocked' in x && x.positionLocked ? '<xdr:cNvGraphicFramePr><a:graphicFrameLocks noMove="1" noResize="1"/></xdr:cNvGraphicFramePr>' : '<xdr:cNvGraphicFramePr/>'}</xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${uri}"><${el}${sle} name="${escapeAttr(name)}"/></a:graphicData></a:graphic></xdr:graphicFrame>`
  const fallback = `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>`
  return `<xdr:twoCellAnchor editAs="oneCell">${marker('from', place.from, place.fromOffset)}${marker('to', place.to, place.toOffset)}<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:${requires}="${ns}" Requires="${requires}">${frame}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent><xdr:clientData${x.locked === false ? ' fLocksWithSheet="0"' : ''}/></xdr:twoCellAnchor>`
}

/**
 * What of a sheet's slicers or timelines the file holds: the framed. Excel opens a file with one that has no frame and drops it
 * when it saves, so a place is what the writer asks of them, here alone; the others go, reported.
 */
function framed<X extends Slicer | Timeline>(list: X[], sheet: string, issues: Issues): (X & { place: DrawingPlace })[] {
  return list.filter((x): x is X & { place: DrawingPlace } => {
    if (x.place) return true
    issues.add('pivotTable', { key: 'pivotNotWritten' }, { sheet, name: x.name })
    return false
  })
}

/** A table of the file by its name: its id and its columns' ids by name; undefined for none. */
export type FileTableIds = (table: string) => { id: number; column(name: string): number | undefined } | undefined

/** The parts of the book's slicers and timelines; `issues` takes what it leaves out (`slicerPackage`'s comment). `tables` gives the ids of the file's tables. */
export function slicerPackage(book: BookInput, issues: Issues, tables: FileTableIds = () => undefined): SlicerPackage {
  const out: SlicerPackage = { files: [], contentTypes: [], caches: [], sheets: book.sheets.map(() => ({})), pivotCacheIds: new Set(), cacheNames: [] }
  // a part `xl/${path}${n}.xml` of its content type; its target from a sheet's part
  const add = (path: string, n: number, text: string, type: string) => {
    const name = `xl/${path}${n}.xml`
    out.files.push({ name, text })
    out.contentTypes.push([`/${name}`, type])
    return `../${path}${n}.xml`
  }
  // a table's cache by its table's and column's ids in the file
  const ids = (c: SlicerCache) => {
    const t = c.table && tables(c.table.name)
    const column = t && t.column(c.table!.column)
    return t && column !== undefined ? { tableId: t.id, column } : undefined
  }
  const caches = <C extends { name: string; cache?: number }>(kind: 'slicerCache' | 'timelineCache', list: C[], xml: (book: BookInput, c: C) => string) =>
    list.forEach((c, k) => {
      const table = kind === 'slicerCache' ? ids(c as unknown as SlicerCache) : undefined
      out.caches.push({ kind, target: add(`${kind}s/${kind}`, k + 1, table ? tableSlicerCacheXml(c as unknown as SlicerCache, table) : xml(book, c), TYPES[kind]).slice(3), ...(table ? { table: true as const } : {}) })
      if (c.cache !== undefined) out.pivotCacheIds.add(c.cache)
      out.cacheNames.push(c.name)
    })
  // Excel opens no file whose slicer cache's name is no name (the set `links/file-refusals`), or a name the workbook holds for anything else (`links/slicer-names`)
  const nameFree = (name: string) => !book.names?.some((n) => n.scope === undefined && caseKey(n.name) === caseKey(name) && !NO_VALUE.test(n.ref))
  const fit = <C extends { name: string }>(c: C, ok: boolean) => {
    if (!ok || !nameFree(c.name)) issues.add('pivotTable', { key: 'pivotNotWritten' }, { name: c.name })
    return ok && nameFree(c.name)
  }
  // each cache's name becomes a defined name of the workbook, so it must be one
  const named = (name: string) => nameProblem(name, localeBooleans(book.locale)) === null
  const pivotCaches = (book.slicerCaches ?? []).filter((c) => fit(c, (c.cache !== undefined || ids(c) !== undefined) && named(c.name)))
  const ofTable = new Set(pivotCaches.filter((c) => c.cache === undefined).map((c) => caseKey(c.name)))
  const timelineCaches = (book.timelineCaches ?? []).filter((c) => fit(c, named(c.name)))
  const taken = new Set([...pivotCaches, ...timelineCaches].map((c) => caseKey(c.name)))
  const own = <X extends Slicer | Timeline>(list: X[] | undefined, sheet: string) => framed((list ?? []).filter((x) => taken.has(caseKey(x.cache))), sheet, issues)
  const shown = book.sheets.map((s) => ({ slicers: own(s.slicers, s.name), timelines: own(s.timelines, s.name) }))
  const seen = new Set(shown.flatMap((s) => [...s.slicers, ...s.timelines]).map((x) => caseKey(x.cache)))
  const asked = new Set(book.sheets.flatMap((s) => [...(s.slicers ?? []), ...(s.timelines ?? [])]).map((x) => caseKey(x.cache)))
  // a cache goes with what showed it, reported above; one that showed nothing is reported here
  const needed = <C extends { name: string }>(c: C) => {
    if (!asked.has(caseKey(c.name))) issues.add('pivotTable', { key: 'pivotNotWritten' }, { name: c.name })
    return seen.has(caseKey(c.name))
  }
  caches('slicerCache', pivotCaches.filter(needed), slicerCacheXml)
  caches('timelineCache', timelineCaches.filter(needed), timelineCacheXml)
  let drawings = 0
  let [slicerParts, timelineParts] = [0, 0]
  book.sheets.forEach((s, i) => {
    const { slicers, timelines } = shown[i]
    if (!slicers.length && !timelines.length) return
    const sheet = out.sheets[i]
    let id = 1
    const table = (x: Slicer) => ofTable.has(caseKey(x.cache))
    const frames = [...timelines.map((t) => frameXml(t, ++id, 'timeline')), ...slicers.map((x) => frameXml(x, ++id, table(x) ? 'tableSlicer' : 'slicer'))].join('')
    sheet.drawing = add('drawings/drawing', ++drawings, `${HEAD}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${frames}</xdr:wsDr>`, DRAWING_TYPE)
    // a table's slicers in a part of their own, which the sheet's extension of Excel 2013 names
    const [ofPivots, ofTables] = [slicers.filter((x) => !table(x)), slicers.filter(table)]
    if (ofPivots.length) sheet.slicers = add('slicers/slicer', ++slicerParts, slicersPartXml(ofPivots), TYPES.slicer)
    if (ofTables.length) sheet.tableSlicers = add('slicers/slicer', ++slicerParts, slicersPartXml(ofTables, TABLE_ROW_HEIGHT), TYPES.slicer)
    if (timelines.length) sheet.timelines = add('timelines/timeline', ++timelineParts, timelinesPartXml(timelines), TYPES.timeline)
  })
  return out
}

/**
 * The workbook's extensions naming the caches by their relationships' ids, in order: slicer caches of PivotTables, of
 * tables (Excel 2013's extension), then timeline caches.
 */
export function workbookSlicerExt(ids: { kind: 'slicerCache' | 'timelineCache'; id: string; table?: true }[]): string {
  const refs = (table: boolean) => ids.filter((x) => x.kind === 'slicerCache' && !x.table === !table).map((x) => `<x14:slicerCache r:id="${x.id}"/>`)
  const [s, st] = [refs(false), refs(true)]
  const t = ids.filter((x) => x.kind === 'timelineCache').map((x) => `<x15:timelineCacheRef r:id="${x.id}"/>`)
  return (
    (s.length ? `<ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x14="${X14}"><x14:slicerCaches>${s.join('')}</x14:slicerCaches></ext>` : '') +
    (st.length ? `<ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}" xmlns:x15="${X15}"><x15:slicerCaches xmlns:x14="${X14}">${st.join('')}</x15:slicerCaches></ext>` : '') +
    (t.length ? `<ext uri="{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}" xmlns:x15="${X15}"><x15:timelineCacheRefs>${t.join('')}</x15:timelineCacheRefs></ext>` : '')
  )
}

/** A sheet's extensions naming its slicers' parts, of PivotTables and of tables, and its timelines' by their relationships' ids. */
const slicerList = (id: string, ns: string) => `<x14:slicerList${ns}><x14:slicer r:id="${id}"/></x14:slicerList></ext>`
export const sheetSlicerExt = (slicers: string | undefined, timelines: string | undefined, tableSlicers?: string) =>
  (slicers ? `<ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="${X14}">${slicerList(slicers, '')}` : '') +
  (tableSlicers ? `<ext uri="{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}" xmlns:x15="${X15}">${slicerList(tableSlicers, ` xmlns:x14="${X14}"`)}` : '') +
  (timelines ? `<ext uri="{7E03D99C-DC04-49d9-9315-930204A7B6E9}" xmlns:x15="${X15}"><x15:timelineRefs><x15:timelineRef r:id="${timelines}"/></x15:timelineRefs></ext>` : '')
