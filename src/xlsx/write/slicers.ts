/**
 * A workbook's slicers and timelines over PivotTables as parts of a package, as Excel 365 writes them: each slicer's
 * cache (x14 `slicerCacheDefinition`) and timeline's cache (x15 `timelineCacheDefinition`), named by the workbook's
 * extensions; each sheet's slicers and timelines in a part of each kind, named by the sheet's extensions, and their
 * frames in the sheet's drawing, at their places. A slicer of a table is not written (reported).
 */
import type { BookInput, DrawingPlace, Slicer, SlicerCache, Timeline, TimelineCache, TimelineLevel } from '../../contract.js'
import { caseKey } from '../../collate.js'
import { sheetKey } from '../../syntax/print.js'
import { parseAddress } from '../../model.js'
import type { Issues } from '../issues.js'
import { escapeAttr } from '../xml/text.js'
import { HEAD, MAIN } from './xml.js'
import { X14 } from './pivot-cache.js'

const X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'
const MS = 'http://schemas.microsoft.com/office'
export const SLICER_RELS = { slicerCache: `${MS}/2007/relationships/slicerCache`, slicer: `${MS}/2007/relationships/slicer`, timelineCache: `${MS}/2011/relationships/timelineCache`, timeline: `${MS}/2011/relationships/timeline` }
const TYPES = { slicerCache: 'application/vnd.ms-excel.slicerCache+xml', slicer: 'application/vnd.ms-excel.slicer+xml', timelineCache: 'application/vnd.ms-excel.timelineCache+xml', timeline: 'application/vnd.ms-excel.timeline+xml' }
const DRAWING_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml'
const LEVEL: Record<TimelineLevel, number> = { years: 0, quarters: 1, months: 2, days: 3 }
/** A slicer's buttons' height as Excel makes one, in EMU. */
const ROW_HEIGHT = 241300

/** What the slicers and timelines add to a package. */
export interface SlicerPackage {
  files: { name: string; text: string }[]
  contentTypes: [string, string][]
  /** The workbook's relationships to the caches, in order: slicer caches, then timeline caches. */
  caches: { kind: 'slicerCache' | 'timelineCache'; target: string }[]
  /** Each sheet's parts, by the relationships it adds: its drawing, its slicers' part, its timelines' part. */
  sheets: { drawing?: string; slicers?: string; timelines?: string }[]
  /** The pivot caches a slicer or a timeline filters, by id: their `pivotCacheId`. */
  pivotCacheIds: Set<number>
}

/** The tab of a PivotTable's sheet (`tabId`, its `sheetId` as `writeXlsx` numbers them) and its name. */
function pivotTables(book: BookInput, c: { pivots: { sheet: string; name: string }[] }): string {
  const tabs = c.pivots.map((r) => `<pivotTable tabId="${book.sheets.findIndex((s) => sheetKey(s.name) === sheetKey(r.sheet)) + 1}" name="${escapeAttr(r.name)}"/>`)
  return tabs.length ? `<pivotTables>${tabs.join('')}</pivotTables>` : '<pivotTables/>'
}

function slicerCacheXml(book: BookInput, c: SlicerCache): string {
  const items = (c.items ?? []).map((x) => `<i x="${x.item}"${x.selected !== false ? ' s="1"' : ''}${x.noData ? ' nd="1"' : ''}/>`)
  const nd = c.noData ?? {}
  const cross = nd.mark === false ? ' crossFilter="none"' : nd.last === false ? ' crossFilter="showItemsWithNoData"' : ''
  const a = `pivotCacheId="${c.cache}"${c.sort === 'descending' ? ' sortOrder="descending"' : ''}${c.customListSort === false ? ' customListSort="0"' : ''}${c.showMissing === false ? ' showMissing="0"' : ''}${cross}`
  const hide = nd.hide ? `<extLst><x:ext uri="{470722E0-AACD-4C17-9CDC-17EF765DBC7E}" xmlns:x15="${X15}"><x15:slicerCacheHideItemsWithNoData/></x:ext></extLst>` : ''
  return `${HEAD}<slicerCacheDefinition xmlns="${X14}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" xmlns:x="${MAIN}" name="${escapeAttr(c.name)}" sourceName="${escapeAttr(c.field)}">${pivotTables(book, c)}<data><tabular ${a}><items count="${items.length}">${items.join('')}</items></tabular></data>${hide}</slicerCacheDefinition>`
}

/** A date of a timeline, ISO 8601 as the file writes it: with its time. */
const isoTime = (d: string) => (d.length === 10 ? `${d}T00:00:00` : d)

/** January 1 of a date's year, or of the next year's. */
const yearOf = (d: string, next: boolean) => `${Number(d.slice(0, 4)) + (next ? 1 : 0)}-01-01T00:00:00`

function timelineCacheXml(book: BookInput, c: TimelineCache): string {
  const field = book.pivotCaches?.find((x) => x.id === c.cache)?.fields.find((f) => caseKey(f.name) === caseKey(c.field))
  const min = typeof field?.facts?.min === 'object' ? field.facts.min.date : undefined
  const max = typeof field?.facts?.max === 'object' ? field.facts.max.date : undefined
  // the dates it spans: the model's, else from January 1 of its first date's year to January 1 after its last's
  const bounds = c.bounds ?? (min && max ? { start: yearOf(min, false), end: yearOf(max, true) } : undefined)
  const sel = c.selection
  const type = !sel ? '' : ` filterType="${isoTime(sel.start) === isoTime(sel.end) ? 'dateEqual' : 'dateBetween'}"`
  const state = `<state minimalRefreshVersion="6" lastRefreshVersion="6" pivotCacheId="${c.cache}"${type}>${sel ? `<selection startDate="${isoTime(sel.start)}" endDate="${isoTime(sel.end)}"/>` : ''}${bounds ? `<bounds startDate="${isoTime(bounds.start)}" endDate="${isoTime(bounds.end)}"/>` : ''}</state>`
  return `${HEAD}<timelineCacheDefinition xmlns="${X15}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" xmlns:x="${MAIN}" name="${escapeAttr(c.name)}" sourceName="${escapeAttr(c.field)}">${pivotTables(book, c)}${state}</timelineCacheDefinition>`
}

const bool = (name: string, v: boolean | undefined, dflt: boolean) => (v === undefined || v === dflt ? '' : ` ${name}="${v ? 1 : 0}"`)

function slicerXml(s: Slicer): string {
  return `<slicer name="${escapeAttr(s.name)}" cache="${escapeAttr(s.cache)}" caption="${escapeAttr(s.caption ?? s.name)}"${s.startItem ? ` startItem="${s.startItem}"` : ''}${s.columns ? ` columnCount="${s.columns}"` : ''}${bool('showCaption', s.showCaption, true)}${s.style ? ` style="${escapeAttr(s.style)}"` : ''}${bool('lockedPosition', s.locked, false)} rowHeight="${s.rowHeight ?? ROW_HEIGHT}"/>`
}

function timelineXml(t: Timeline): string {
  const a =
    ` name="${escapeAttr(t.name)}" cache="${escapeAttr(t.cache)}" caption="${escapeAttr(t.caption ?? t.name)}"` +
    bool('showHeader', t.showHeader, true) +
    bool('showSelectionLabel', t.showSelectionLabel, true) +
    bool('showTimeLevel', t.showLevel, true) +
    bool('showHorizontalScrollbar', t.showScrollbar, false) +
    ` level="${LEVEL[t.level ?? 'months']}" selectionLevel="${LEVEL[t.selectionLevel ?? 'years']}"` +
    (t.scrollPosition ? ` scrollPosition="${isoTime(t.scrollPosition)}"` : '') +
    (t.style ? ` style="${escapeAttr(t.style)}"` : '')
  return `<timeline${a}/>`
}

/** A cell of an anchor with its offset. */
const marker = (tag: string, cell: string, off: [number, number] | undefined) => {
  const { row, col } = parseAddress(cell)
  return `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>${off?.[0] ?? 0}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${off?.[1] ?? 0}</xdr:rowOff></xdr:${tag}>`
}

/** A slicer's or timeline's frame in a drawing, at its place, with the shape older readers show. */
function frameXml(name: string, place: DrawingPlace, id: number, timeline: boolean): string {
  const [requires, ns, uri, el] = timeline
    ? ['tsle', 'http://schemas.microsoft.com/office/drawing/2012/timeslicer', 'http://schemas.microsoft.com/office/drawing/2012/timeslicer', 'tsle:timeslicer']
    : ['a14', 'http://schemas.microsoft.com/office/drawing/2010/main', 'http://schemas.microsoft.com/office/drawing/2010/slicer', 'sle:slicer']
  const sle = timeline ? '' : ' xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer"'
  const frame = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${escapeAttr(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${uri}"><${el}${sle} name="${escapeAttr(name)}"/></a:graphicData></a:graphic></xdr:graphicFrame>`
  const fallback = `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>`
  return `<xdr:twoCellAnchor editAs="oneCell">${marker('from', place.from, place.fromOffset)}${marker('to', place.to, place.toOffset)}<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:${requires}="${ns}" Requires="${requires}">${frame}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`
}

/** The parts of the book's slicers and timelines over PivotTables; `issues` takes the slicers of tables, left out. */
export function slicerPackage(book: BookInput, issues: Issues): SlicerPackage {
  const out: SlicerPackage = { files: [], contentTypes: [], caches: [], sheets: book.sheets.map(() => ({})), pivotCacheIds: new Set() }
  const pivotCaches = (book.slicerCaches ?? []).filter((c) => {
    if (c.cache !== undefined) return true
    issues.add('pivotTable', { key: 'pivotNotWritten' })
    return false
  })
  const written = new Set(pivotCaches.map((c) => caseKey(c.name)))
  pivotCaches.forEach((c, k) => {
    const name = `xl/slicerCaches/slicerCache${k + 1}.xml`
    out.files.push({ name, text: slicerCacheXml(book, c) })
    out.contentTypes.push([`/${name}`, TYPES.slicerCache])
    out.caches.push({ kind: 'slicerCache', target: `slicerCaches/slicerCache${k + 1}.xml` })
    out.pivotCacheIds.add(c.cache!)
  })
  ;(book.timelineCaches ?? []).forEach((c, k) => {
    const name = `xl/timelineCaches/timelineCache${k + 1}.xml`
    out.files.push({ name, text: timelineCacheXml(book, c) })
    out.contentTypes.push([`/${name}`, TYPES.timelineCache])
    out.caches.push({ kind: 'timelineCache', target: `timelineCaches/timelineCache${k + 1}.xml` })
    out.pivotCacheIds.add(c.cache)
  })
  let drawings = 0
  let [slicerParts, timelineParts] = [0, 0]
  book.sheets.forEach((s, i) => {
    const slicers = (s.slicers ?? []).filter((x) => written.has(caseKey(x.cache)))
    const timelines = s.timelines ?? []
    if (!slicers.length && !timelines.length) return
    const sheet = out.sheets[i]
    let id = 1
    const frames = [...timelines.map((t) => (t.place ? frameXml(t.name, t.place, ++id, true) : '')), ...slicers.map((x) => (x.place ? frameXml(x.name, x.place, ++id, false) : ''))].join('')
    if (frames) {
      const name = `xl/drawings/drawing${++drawings}.xml`
      out.files.push({ name, text: `${HEAD}<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${frames}</xdr:wsDr>` })
      out.contentTypes.push([`/${name}`, DRAWING_TYPE])
      sheet.drawing = `../drawings/drawing${drawings}.xml`
    }
    if (slicers.length) {
      const name = `xl/slicers/slicer${++slicerParts}.xml`
      out.files.push({ name, text: `${HEAD}<slicers xmlns="${X14}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" xmlns:x="${MAIN}">${slicers.map(slicerXml).join('')}</slicers>` })
      out.contentTypes.push([`/${name}`, TYPES.slicer])
      sheet.slicers = `../slicers/slicer${slicerParts}.xml`
    }
    if (timelines.length) {
      const name = `xl/timelines/timeline${++timelineParts}.xml`
      out.files.push({ name, text: `${HEAD}<timelines xmlns="${X15}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" xmlns:x="${MAIN}">${timelines.map(timelineXml).join('')}</timelines>` })
      out.contentTypes.push([`/${name}`, TYPES.timeline])
      sheet.timelines = `../timelines/timeline${timelineParts}.xml`
    }
  })
  return out
}

/** The workbook's extensions naming the caches by their relationships' ids, in order: slicer caches, then timeline caches. */
export function workbookSlicerExt(ids: { kind: 'slicerCache' | 'timelineCache'; id: string }[]): string {
  const s = ids.filter((x) => x.kind === 'slicerCache').map((x) => `<x14:slicerCache r:id="${x.id}"/>`)
  const t = ids.filter((x) => x.kind === 'timelineCache').map((x) => `<x15:timelineCacheRef r:id="${x.id}"/>`)
  return (
    (s.length ? `<ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x14="${X14}"><x14:slicerCaches>${s.join('')}</x14:slicerCaches></ext>` : '') +
    (t.length ? `<ext uri="{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}" xmlns:x15="${X15}"><x15:timelineCacheRefs>${t.join('')}</x15:timelineCacheRefs></ext>` : '')
  )
}

/** A sheet's extensions naming its slicers' and timelines' parts by their relationships' ids. */
export const sheetSlicerExt = (slicers: string | undefined, timelines: string | undefined) =>
  (slicers ? `<ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="${X14}"><x14:slicerList><x14:slicer r:id="${slicers}"/></x14:slicerList></ext>` : '') +
  (timelines ? `<ext uri="{7E03D99C-DC04-49d9-9315-930204A7B6E9}" xmlns:x15="${X15}"><x15:timelineRefs><x15:timelineRef r:id="${timelines}"/></x15:timelineRefs></ext>` : '')
