/**
 * Slicers (Excel 2010's x14 parts) and timelines (Excel 2013's x15 parts): the workbook's caches with the field, the
 * PivotTables they filter by sheet and name, the items chosen and the dates; a sheet's slicers and timelines with their
 * captions, styles and places, which the sheet's drawing holds by name.
 */
import type { DrawingPlace, PivotRef, Slicer, SlicerCache, Timeline, TimelineCache, TimelineLevel } from '../../contract.js'
import { formatAddress } from '../../model.js'
import { attributes, attrsAt, element, elements, flag, within } from '../xml/scan.js'
import { attrText, take } from './pivot-cache.js'

/** The PivotTables a cache names by the sheet's id (`tabId`, the workbook's `sheetId`) and name. */
function pivotsOf(xml: string, sheetOf: (id: string | undefined) => string | undefined): PivotRef[] {
  return within(xml, 'pivotTables', 'pivotTable').flatMap((e) => {
    const a = attributes(e.attrs)
    const sheet = sheetOf(a.tabId)
    return sheet === undefined ? [] : [{ sheet, name: attrText(a.name) }]
  })
}

/** What the reader of the caches needs of the workbook: sheets by id, a pivot cache by its Excel 2010 id, a table by its id. */
export interface SlicerContext {
  sheetOf: (id: string | undefined) => string | undefined
  cacheOf: (id: string | undefined) => number | undefined
  tableOf: (id: string | undefined) => string | undefined
}

export function readSlicerCache(xml: string, ctx: SlicerContext): SlicerCache {
  const a = attrsAt(xml, 'slicerCacheDefinition')
  const c: SlicerCache = { name: attrText(a.name), field: attrText(a.sourceName), pivots: pivotsOf(xml, ctx.sheetOf) }
  const tab = element(xml, 'tabular')
  if (tab) {
    const ta = attributes(tab.attrs)
    const cache = ctx.cacheOf(ta.pivotCacheId)
    if (cache !== undefined) c.cache = cache
    c.items = within(tab.body ?? '', 'items', 'i').map((e) => {
      const ia = attributes(e.attrs)
      return { item: Number(ia.x ?? 0), ...(flag(ia.s) ? {} : { selected: false }), ...(flag(ia.nd) ? { noData: true } : {}) }
    })
    if (ta.sortOrder === 'descending') c.sort = 'descending'
    take(c, ta, 'customListSort:t showMissing:t')
    const cross = ta.crossFilter ?? 'showItemsWithDataAtTop'
    const noData: NonNullable<SlicerCache['noData']> = {}
    if (element(xml, 'slicerCacheHideItemsWithNoData')) noData.hide = true
    if (cross === 'none') noData.mark = false
    if (cross === 'showItemsWithNoData') noData.last = false
    if (Object.keys(noData).length) c.noData = noData
  }
  const table = element(xml, 'tableSlicerCache')
  if (table) {
    const t = attributes(table.attrs)
    c.table = { name: ctx.tableOf(t.tableId) ?? '', column: attrText(t.column) }
  }
  return c
}

const LEVELS: TimelineLevel[] = ['years', 'quarters', 'months', 'days']

export function readTimelineCache(xml: string, ctx: SlicerContext): TimelineCache {
  const a = attrsAt(xml, 'timelineCacheDefinition')
  const state = attrsAt(xml, 'state')
  const c: TimelineCache = { name: attrText(a.name), field: attrText(a.sourceName), pivots: pivotsOf(xml, ctx.sheetOf), cache: ctx.cacheOf(state.pivotCacheId) ?? -1 }
  for (const k of ['selection', 'bounds'] as const) {
    const s = attrsAt(xml, k)
    if (s.startDate !== undefined && s.endDate !== undefined) c[k] = { start: s.startDate, end: s.endDate }
  }
  return c
}

/** A frame of a slicer (`sle:slicer`) or a timeline (`tsle:timeslicer`) in an anchor's body, as Excel puts each in `mc:AlternateContent`. */
const frameOf = (body: string) => element(body, 'slicer') ?? element(body, 'timeslicer')

/** A drawing that shows only slicers and timelines: every anchor a frame of one, which the model holds; false without the drawing. */
export function framesOnly(xml: string | undefined): boolean {
  if (xml === undefined) return false
  return ['twoCellAnchor', 'oneCellAnchor', 'absoluteAnchor'].every((local) => elements(xml, local).every((a) => frameOf(a.body ?? '') !== undefined))
}

/** The places of a drawing's frames by the name of the slicer or timeline each shows. */
export function drawingPlaces(xml: string | undefined): Map<string, DrawingPlace> {
  const out = new Map<string, DrawingPlace>()
  for (const anchor of elements(xml ?? '', 'twoCellAnchor')) {
    const body = anchor.body ?? ''
    const shown = frameOf(body)
    if (!shown) continue
    const corner = (local: string): [string, [number, number]] => {
      const b = element(body, local)?.body ?? ''
      const n = (k: string) => Number(element(b, k)?.body ?? 0)
      return [formatAddress(n('row'), n('col')), [n('colOff'), n('rowOff')]]
    }
    const [from, fromOffset] = corner('from')
    const [to, toOffset] = corner('to')
    const place: DrawingPlace = { from, to }
    if (fromOffset[0] || fromOffset[1]) place.fromOffset = fromOffset
    if (toOffset[0] || toOffset[1]) place.toOffset = toOffset
    out.set(attrText(attributes(shown.attrs).name), place)
  }
  return out
}

export function readSlicers(xml: string, places: Map<string, DrawingPlace>): Slicer[] {
  return elements(xml, 'slicer').map((e) => {
    const a = attributes(e.attrs)
    const s = take<Slicer>({ name: '', cache: '' }, a, 'name cache caption showCaption:t style columnCount=columns:n rowHeight:n startItem:p lockedPosition=locked:f')
    const place = places.get(s.name)
    return place ? { ...s, place } : s
  })
}

export function readTimelines(xml: string, places: Map<string, DrawingPlace>): Timeline[] {
  return elements(xml, 'timeline').map((e) => {
    const a = attributes(e.attrs)
    const t = take<Timeline>({ name: '', cache: '' }, a, 'name cache caption showHeader:t showSelectionLabel:t showTimeLevel=showLevel:t showHorizontalScrollbar=showScrollbar:t scrollPosition style')
    for (const k of ['level', 'selectionLevel'] as const) if (a[k] !== undefined) t[k] = LEVELS[Number(a[k])] ?? 'months'
    const place = places.get(t.name)
    return place ? { ...t, place } : t
  })
}
