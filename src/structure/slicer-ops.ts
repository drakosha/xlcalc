/**
 * Operations on slicers and timelines (`editStructure`): one added over a cache of the workbook or a new cache of a
 * field over PivotTables of one pivot cache, changed with its cache (the items or dates chosen, the PivotTables it
 * filters, its look and place), deleted with its cache when it is the last over it. A choice is the filter of every
 * PivotTable the cache filters (`slicer-choice.ts`); they are laid out again from their cache by `xlcalc/pivot`, as
 * Excel does without a refresh, or their cache is marked `stale`.
 */
import type { BookInput, DrawingPlace, PivotRef, Slicer, SlicerCache, Timeline, TimelineCache } from '../contract.js'
import { caseKey } from '../collate.js'
import { InputError } from '../input-error.js'
import { checkSheetOf } from '../input.js'
import { array, checkBoolean, checkString, integer, object, onlyKeys } from '../input/basic.js'
import { checkIsoDate } from '../input/pivot.js'
import { say } from '../locale/messages.js'
import { withField } from '../looks/sheet.js'
import { formatAddress } from '../model.js'
import { pivotArea } from '../pivot/area.js'
import { pivotRefresh } from '../pivot/hook.js'
import { sheetKey } from '../syntax.js'
import type { Applied } from './geometry.js'
import { checkedPivots, patched, relaid } from './pivots.js'
import { sheetIndex } from './rewrite.js'
import { fieldIndex, itemsShown, selectionLevel, slicerField, timelineBounds, withSlicerChoice, withTimelineChoice } from './slicer-choice.js'
import type { ValuesIn } from './values.js'

type Kind = 'slicers' | 'timelines'
type Cache = SlicerCache | TimelineCache
const CACHES = { slicers: 'slicerCaches', timelines: 'timelineCaches' } as const
const KEY = { slicers: 'slicer', timelines: 'timeline' } as const

/** The fields of a cache an edit sets, by kind. */
const CACHE_FIELDS: Record<Kind, string[]> = { slicers: ['pivots', 'items', 'sort', 'customListSort', 'noData', 'showMissing'], timelines: ['pivots', 'selection'] }

/** The operations' fields for `editStructure`'s check. */
export const SLICER_FIELDS: Record<string, string[]> = {
  addSlicer: ['sheet', 'slicer', 'cache'],
  editSlicer: ['slicer', 'set', 'cache'],
  deleteSlicer: ['slicer'],
  addTimeline: ['sheet', 'timeline', 'cache'],
  editTimeline: ['timeline', 'set', 'cache'],
  deleteTimeline: ['timeline'],
}

const refused = (book: BookInput, m: Parameters<typeof say>[1]): Applied => ({ ok: false, code: 'pivot', message: say(book.locale, m) })

const sameRef = (a: PivotRef, b: PivotRef) => sheetKey(a.sheet) === sheetKey(b.sheet) && caseKey(a.name) === caseKey(b.name)

const cachesOf = (book: BookInput, kind: Kind): Cache[] => book[CACHES[kind]] ?? []

/** The book with its caches of `kind`, none left as none. */
function withCaches(book: BookInput, kind: Kind, caches: Cache[]): BookInput {
  const out = { ...book }
  if (caches.length) (out as Record<string, unknown>)[CACHES[kind]] = caches
  else delete out[CACHES[kind]]
  return out
}

/** Every slicer's and timeline's name, unique together in the workbook. */
const names = (book: BookInput) => new Set(book.sheets.flatMap((s) => [...(s.slicers ?? []), ...(s.timelines ?? [])].map((x) => caseKey(x.name))))

/** The first of `base`, `base 1`, `base 2`… (`sep` between) not in `taken`. */
function free(taken: Set<string>, base: string, sep: string): string {
  let name = base
  for (let n = 1; taken.has(caseKey(name)); n++) name = `${base}${sep}${n}`
  return name
}

/** A slicer or timeline by name: its sheet's index, its index there, itself. */
function find(book: BookInput, kind: Kind, v: unknown, path: string): { s: number; k: number; x: Slicer | Timeline } {
  const name = checkString(v, path)
  for (const [s, sheet] of book.sheets.entries()) {
    const k = (sheet[kind] ?? []).findIndex((x) => caseKey(x.name) === caseKey(name))
    if (k >= 0) return { s, k, x: sheet[kind]![k] }
  }
  throw new InputError(path, `unknown ${KEY[kind]} ${JSON.stringify(name)}`)
}

/** PivotTables checked: of the workbook, all over one pivot cache, whose id it gives. */
function pivotsOf(book: BookInput, v: unknown, path: string): { refs: PivotRef[]; cache: number } {
  const refs = array(v, path, 'an array of PivotTables such as [{ sheet: "P", name: "PivotTable1" }]').map((r, i) => {
    const x = object(r, `${path}[${i}]`, 'a PivotTable such as { sheet: "P", name: "PivotTable1" }')
    onlyKeys(x, `${path}[${i}]`, ['sheet', 'name'])
    const sheet = checkSheetOf(book, x.sheet, `${path}[${i}].sheet`)
    const name = checkString(x.name, `${path}[${i}].name`)
    const p = sheet.pivotTables?.find((q) => caseKey(q.name) === caseKey(name))
    if (!p) throw new InputError(`${path}[${i}].name`, `unknown PivotTable ${JSON.stringify(name)} on ${JSON.stringify(sheet.name)}`)
    return { ref: { sheet: sheet.name, name: p.name }, cache: p.cache }
  })
  if (!refs.length) throw new InputError(path, 'expected at least one PivotTable')
  if (refs.some((r) => r.cache !== refs[0].cache)) throw new InputError(path, 'expected PivotTables of one pivot cache')
  const out: PivotRef[] = []
  for (const r of refs) if (!out.some((q) => sameRef(q, r.ref))) out.push(r.ref)
  return { refs: out, cache: refs[0].cache }
}

/** A slicer cache's items as an edit gives them: each an item of the field once, at least one chosen, as Excel keeps one. */
function checkItems(v: unknown, path: string, n: number): NonNullable<SlicerCache['items']> {
  const seen = new Set<number>()
  const items = array(v, path, 'an array of items such as [{ item: 0, selected: false }]').map((x, i) => {
    const o = object(x, `${path}[${i}]`, 'an item such as { item: 0, selected: false }')
    onlyKeys(o, `${path}[${i}]`, ['item', 'selected', 'noData'])
    const item = integer(o.item, `${path}[${i}].item`, 0, Math.max(0, n - 1))
    if (seen.has(item)) throw new InputError(`${path}[${i}].item`, `item ${item} is given twice`)
    seen.add(item)
    if (o.selected !== undefined) checkBoolean(o.selected, `${path}[${i}].selected`)
    if (o.noData !== undefined) checkBoolean(o.noData, `${path}[${i}].noData`)
    return o as NonNullable<SlicerCache['items']>[number]
  })
  if (n && items.filter((x) => x.selected === false).length >= n) throw new InputError(path, 'expected at least one item chosen')
  return items
}

/** A new slicer's or timeline's place: right of the first PivotTable it filters, three columns and fourteen rows, as Excel puts one beside it. */
function placeBeside(book: BookInput, r: PivotRef): DrawingPlace | undefined {
  const p = book.sheets[sheetIndex(book, r.sheet)]?.pivotTables?.find((q) => caseKey(q.name) === caseKey(r.name))
  const x = p && pivotArea(p)
  return x ? { from: formatAddress(x.r1, x.c2 + 2), to: formatAddress(x.r1 + 14, x.c2 + 5) } : undefined
}

/** The PivotTables laid out again from their cache, or marked `stale` without `xlcalc/pivot` when `changed`; as they are when nothing changed and the subpath is not loaded. */
const laidOut = (book: BookInput, refs: PivotRef[], values: ValuesIn, changed: boolean): Applied =>
  refs.length && (changed || pivotRefresh()) ? relaid(book, refs, values, false) : { ok: true, book, geometry: [] }

/** The book with a cache's choice as its PivotTables' filter, on `refs`. */
const chosen = (book: BookInput, kind: Kind, c: Cache, refs?: PivotRef[]) => (kind === 'slicers' ? withSlicerChoice(book, c as SlicerCache, refs) : withTimelineChoice(book, c as TimelineCache, refs))

/** A new cache of `field` over the PivotTables `pivots`, its choice the filter they hold, or the refusal. */
function newCache(book: BookInput, kind: Kind, v: unknown): { book: BookInput; cache: Cache } | Applied {
  const n = object(v, 'op.cache', 'a cache such as { field: "Region", pivots: [{ sheet: "P", name: "PivotTable1" }] }')
  onlyKeys(n, 'op.cache', ['field', 'pivots'])
  const field = checkString(n.field, 'op.cache.field')
  const { refs, cache: id } = pivotsOf(book, n.pivots, 'op.cache.pivots')
  const pc = book.pivotCaches?.find((c) => c.id === id)
  const f = pc ? fieldIndex(pc, field) : -1
  if (!pc || f < 0) return refused(book, { key: 'slicerField', field })
  const fieldName = pc.fields[f].name
  const taken = new Set([...cachesOf(book, kind), ...(book.names ?? [])].map((c) => caseKey(c.name)))
  const name = free(taken, `${kind === 'slicers' ? 'Slicer' : 'NativeTimeline'}_${fieldName.replace(/[^\p{L}\p{N}_]/gu, '_')}`, '')
  if (kind === 'timelines') {
    if (!pc.fields[f].facts?.types?.includes('date') && !pc.fields[f].items?.some((x) => x !== null && typeof x === 'object' && 'date' in x)) return refused(book, { key: 'timelineField', field: fieldName })
    const bounds = timelineBounds(pc, f)
    return { book, cache: { name, field: fieldName, pivots: refs, cache: id, ...(bounds ? { bounds } : {}) } }
  }
  const at = slicerField(book, { field: fieldName, cache: id })!
  const items = itemsShown(at.book, refs[0], at.field, at.cache.fields[at.field].items?.length ?? 0)
  const cache: SlicerCache = { name, field: fieldName, pivots: refs, cache: id, ...(items.length ? { items } : {}) }
  return { book: at.book, cache }
}

function add(book: BookInput, kind: Kind, o: Record<string, unknown>, values: ValuesIn): Applied {
  const s = book.sheets.indexOf(checkSheetOf(book, o.sheet, 'op.sheet'))
  const key = KEY[kind]
  const given = object(o[key], `op.${key}`, `a ${key} such as { caption: "Region" }`)
  if ('cache' in given) throw new InputError(`op.${key}.cache`, `a new ${key} takes its cache by op.cache`)
  let out = book
  let cache: Cache
  let fresh = false
  if (typeof o.cache === 'string') {
    const c = cachesOf(book, kind).find((x) => caseKey(x.name) === caseKey(o.cache as string))
    if (!c) throw new InputError('op.cache', `unknown ${key} cache ${JSON.stringify(o.cache)}`)
    cache = c
  } else {
    const made = newCache(book, kind, o.cache)
    if ('ok' in made) return made
    ;({ book: out, cache } = made)
    fresh = true
  }
  const taken = names(out)
  const name = given.name === undefined ? free(taken, cache.field, ' ') : checkString(given.name, `op.${key}.name`)
  if (!name.trim()) throw new InputError(`op.${key}.name`, `expected a ${key}'s name, got an empty string`)
  if (taken.has(caseKey(name))) return refused(out, { key: 'slicerNameTaken', name })
  const place = given.place === undefined && cache.pivots.length ? placeBeside(out, cache.pivots[0]) : undefined
  // Excel captions it by its field
  const item = { caption: cache.field, ...given, name, cache: cache.name, ...(place ? { place } : {}) } as Slicer & Timeline
  const sheets = [...out.sheets]
  sheets[s] = withField(sheets[s], kind, [...((sheets[s][kind] ?? []) as (Slicer | Timeline)[]), item] as never)
  out = { ...out, sheets }
  if (fresh) out = chosen(withCaches(out, kind, [...cachesOf(out, kind), cache]), kind, cache)
  // a new slicer filters nothing yet: its PivotTables are laid out again only to learn its items' data
  return laidOut(checkedPivots(out), cache.pivots, values, false)
}

function edit(book: BookInput, kind: Kind, o: Record<string, unknown>, values: ValuesIn): Applied {
  const key = KEY[kind]
  const { s, k, x } = find(book, kind, o[key], `op.${key}`)
  let out = book
  if (o.set !== undefined) {
    const set = object(o.set, 'op.set', `a patch of a ${key} such as { caption: "Region" }`)
    if ('name' in set || 'cache' in set) throw new InputError('op.set', `a ${key}'s name and cache are not edited`)
    const list = [...((book.sheets[s][kind] ?? []) as (Slicer | Timeline)[])]
    list[k] = patched(list[k], set)
    const sheets = [...book.sheets]
    sheets[s] = withField(sheets[s], kind, list as never)
    out = { ...out, sheets }
  }
  if (o.cache === undefined) return { ok: true, book: checkedPivots(out), geometry: [] }
  const patch = object(o.cache, 'op.cache', `a patch of a ${key} cache such as { items: [...] }`)
  onlyKeys(patch, 'op.cache', CACHE_FIELDS[kind])
  const caches = cachesOf(out, kind)
  const c = caches.find((y) => caseKey(y.name) === caseKey(x.cache))!
  // a table's slicer chooses by its table's AutoFilter, which the file keeps; its look and order change alone
  if (c.cache === undefined) {
    if (patch.items != null || patch.pivots != null) throw new InputError('op.cache', "a table's slicer chooses by the table's AutoFilter, not by its items or PivotTables")
    return { ok: true, book: checkedPivots(withCaches(out, kind, caches.map((y) => (y === c ? patched(c, patch) : y)))), geometry: [] }
  }
  if (patch.pivots === null) throw new InputError('op.cache.pivots', `a ${key} cache filters at least one PivotTable`)
  if (patch.pivots !== undefined) {
    const { refs, cache } = pivotsOf(out, patch.pivots, 'op.cache.pivots')
    if (cache !== c.cache) throw new InputError('op.cache.pivots', "expected PivotTables of the cache's pivot cache")
    patch.pivots = refs
  }
  if (kind === 'slicers' && patch.items != null) {
    const at = slicerField(out, c as SlicerCache)
    out = at?.book ?? out
    checkItems(patch.items, 'op.cache.items', at?.cache.fields[at.field].items?.length ?? 0)
  }
  if (kind === 'timelines' && patch.selection != null) {
    const sel = object(patch.selection, 'op.cache.selection', 'dates such as { start: "2024-01-01T00:00:00", end: "2024-12-31T00:00:00" }')
    onlyKeys(sel, 'op.cache.selection', ['start', 'end'])
    checkIsoDate(sel.start, 'op.cache.selection.start')
    checkIsoDate(sel.end, 'op.cache.selection.end')
    if (String(sel.start) > String(sel.end)) throw new InputError('op.cache.selection', 'expected a start on or before the end')
  }
  const next = patched(c, patch)
  // the PivotTables it filters now take its choice; one no longer connected keeps the filter it has, as after a deletion
  const choice = patch.items !== undefined || patch.selection !== undefined
  const added = next.pivots.filter((r) => !c.pivots.some((q) => sameRef(q, r)))
  out = withCaches(out, kind, caches.map((y) => (y === c ? next : y)))
  if (kind === 'timelines' && patch.selection !== undefined) out = withLevels(out, next as TimelineCache)
  if (choice || added.length) out = chosen(out, kind, next, choice ? next.pivots : added)
  const refs = [...c.pivots, ...added]
  return laidOut(checkedPivots(out), refs, values, choice || added.length > 0)
}

/** The book with the timelines of `tc` at the level its dates are chosen at, none without a selection. */
function withLevels(book: BookInput, tc: TimelineCache): BookInput {
  const level = tc.selection && selectionLevel(tc.selection)
  const sheets = book.sheets.map((sh) =>
    sh.timelines?.some((t) => caseKey(t.cache) === caseKey(tc.name)) ? { ...sh, timelines: sh.timelines.map((t) => (caseKey(t.cache) !== caseKey(tc.name) ? t : patched(t, { selectionLevel: level ?? null }))) } : sh,
  )
  return { ...book, sheets }
}

function remove(book: BookInput, kind: Kind, o: Record<string, unknown>): Applied {
  const { s, k, x } = find(book, kind, o[KEY[kind]], `op.${KEY[kind]}`)
  const list = ((book.sheets[s][kind] ?? []) as (Slicer | Timeline)[]).filter((_, i) => i !== k)
  const sheets = [...book.sheets]
  sheets[s] = withField(sheets[s], kind, (list.length ? list : undefined) as never)
  const out: BookInput = { ...book, sheets }
  const last = !sheets.some((sh) => (sh[kind] as (Slicer | Timeline)[] | undefined)?.some((y) => caseKey(y.cache) === caseKey(x.cache)))
  if (!last) return { ok: true, book: out, geometry: [] }
  // its cache goes with it; the PivotTables keep the filter it made, as in Excel
  return { ok: true, book: withCaches(out, kind, cachesOf(book, kind).filter((c) => caseKey(c.name) !== caseKey(x.cache))), geometry: [] }
}

export function slicerOp(book: BookInput, o: Record<string, unknown>, values: ValuesIn): Applied {
  const op = String(o.op)
  const kind: Kind = op.endsWith('Slicer') ? 'slicers' : 'timelines'
  return op.startsWith('add') ? add(book, kind, o, values) : op.startsWith('edit') ? edit(book, kind, o, values) : remove(book, kind, o)
}
