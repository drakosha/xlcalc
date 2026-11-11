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
import { sayPivot } from '../locale/pivot-messages.js'
import { withField } from '../looks/sheet.js'
import { formatAddress, MAX_COLS, MAX_ROWS } from '../model.js'
import { DEFAULT_ROW_HEIGHT } from '../column-width.js'
import { DEFAULT_SLICER_SIZE, DEFAULT_TIMELINE_SIZE } from '../drawing-size.js'
import { entryAt, lines } from '../layout/spans.js'
import { pivotArea } from '../pivot/area.js'
import { tableKey, tablesOf } from '../tables/model.js'
import type { Rect } from './area.js'
import { pivotRefresh } from '../pivot/hook.js'
import { groupedItems } from '../pivot/items.js'
import type { Applied } from './geometry.js'
import { checkedPivots, patched, relaid } from './pivots.js'
import { sheetIndex } from './rewrite.js'
import { fieldIndex, itemsShown, sameRef, selectionLevel, slicerField, timelineBounds, withSlicerChoice, withTimelineChoice } from './slicer-choice.js'
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

const refused = (book: BookInput, m: Parameters<typeof sayPivot>[1]): Applied => ({ ok: false, code: 'pivot', message: sayPivot(book.locale, m) })

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

/**
 * A name given to a slicer or timeline (`path`), checked as Excel takes `Slicer.Name`: any text but the empty one (a space,
 * `A1`, a cache's name, 256 characters); the refusal when it is empty or another slicer or timeline has it, in any case
 * (`own` is the one renamed), as a sheet's name is refused.
 */
function givenName(book: BookInput, v: unknown, path: string, own?: string): string | Parameters<typeof sayPivot>[1] {
  const name = checkString(v, path)
  if (!name) return { key: 'slicerNameEmpty' }
  const taken = names(book)
  if (own !== undefined) taken.delete(caseKey(own))
  return taken.has(caseKey(name)) ? { key: 'slicerNameTaken', name } : name
}

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

/** EMU in a pixel and in a point; a column's pixels are 7 to a stored width's digit, 64 for the default width, as Excel for Windows draws Calibri 11. */
const [PIXEL, POINT, COLUMN] = [9525, 12_700, 64]

/** The line and offset in EMU where a length from the start of line `at` ends, by the lines' sizes. */
function reach(at: number, length: number, size: (k: number) => number, last: number): [line: number, offset: number] {
  let [k, left] = [at, length]
  while (k < last && size(k) <= left) left -= size(k++)
  return [k, left]
}

/** The area a new slicer is placed beside: the first PivotTable its cache filters, or its table. */
function sourceArea(book: BookInput, c: Cache): Rect | undefined {
  if ('table' in c && c.table) return tablesOf(book).byName.get(tableKey(c.table.name))
  const r = c.pivots[0]
  const p = r && book.sheets[sheetIndex(book, r.sheet)]?.pivotTables?.find((q) => caseKey(q.name) === caseKey(r.name))
  return (p && pivotArea(p)) || undefined
}

/**
 * A new slicer's or timeline's place: from two columns right of the area of its source (`sourceArea`), at its top row,
 * and Excel's size for it (`DEFAULT_SLICER_SIZE`, `DEFAULT_TIMELINE_SIZE`) over the sheet's columns and rows. Excel's
 * object model puts one in the middle of the window, and its Insert Slicer one after another there; the workbook has no
 * window, so the place is the library's.
 */
function placeBeside(book: BookInput, kind: Kind, s: number, x: Rect | undefined): DrawingPlace | undefined {
  if (!x) return undefined
  const sheet = book.sheets[s]
  const [cols, rows] = [lines(sheet.columns, 'c'), lines(sheet.rows, 'r')]
  const col = (c: number) => {
    const e = entryAt(cols, c)
    return e?.hidden ? 0 : (e?.width === undefined ? COLUMN : Math.round(e.width * 7)) * PIXEL
  }
  const row = (k: number) => {
    const e = entryAt(rows, k)
    return e?.hidden ? 0 : (e?.height ?? DEFAULT_ROW_HEIGHT) * POINT
  }
  const size = kind === 'slicers' ? DEFAULT_SLICER_SIZE : DEFAULT_TIMELINE_SIZE
  const [c2, dx] = reach(x.c2 + 2, size.width, col, MAX_COLS - 1)
  const [r2, dy] = reach(x.r1, size.height, row, MAX_ROWS - 1)
  return { from: formatAddress(x.r1, x.c2 + 2), to: formatAddress(r2, c2), ...(dx || dy ? { toOffset: [dx, dy] as [number, number] } : {}) }
}

/** The PivotTables laid out again from their cache, or marked `stale` without `xlcalc/pivot` when `changed`; as they are when nothing changed and the subpath is not loaded. */
const laidOut = (book: BookInput, refs: PivotRef[], values: ValuesIn, changed: boolean): Applied =>
  refs.length && (changed || pivotRefresh()) ? relaid(book, refs, values, false) : { ok: true, book, geometry: [] }

/** The book with a cache's choice as its PivotTables' filter, on `refs`. */
const chosen = (book: BookInput, kind: Kind, c: Cache, refs?: PivotRef[]) => (kind === 'slicers' ? withSlicerChoice(book, c as SlicerCache, refs) : withTimelineChoice(book, c as TimelineCache, refs))

/** A free name of a new cache of `field`, among the caches and defined names of the workbook. */
function cacheName(book: BookInput, kind: Kind, field: string): string {
  const taken = new Set([...cachesOf(book, kind), ...(book.names ?? [])].map((c) => caseKey(c.name)))
  return free(taken, `${kind === 'slicers' ? 'Slicer' : 'NativeTimeline'}_${field.replace(/[^\p{L}\p{N}_]/gu, '_')}`, '')
}

/** A new slicer cache of a table's column, by their names, as Excel's Insert Slicer on a table makes it: its field the column's name, no PivotTables. */
function tableCache(book: BookInput, n: Record<string, unknown>): SlicerCache {
  onlyKeys(n, 'op.cache', ['table', 'column'])
  const name = checkString(n.table, 'op.cache.table')
  const t = tablesOf(book).byName.get(tableKey(name))
  if (!t) throw new InputError('op.cache.table', `unknown table ${JSON.stringify(name)}`)
  const want = checkString(n.column, 'op.cache.column')
  const column = t.input.columns.find((c) => caseKey(c.name) === caseKey(want))?.name
  if (column === undefined) throw new InputError('op.cache.column', `the table ${t.name} has no column ${JSON.stringify(want)}`)
  return { name: cacheName(book, 'slicers', column), field: column, pivots: [], table: { name: t.name, column } }
}

/** A new cache of `field` over the PivotTables `pivots`, its choice the filter they hold, or of a table's column; or the refusal. */
function newCache(book: BookInput, kind: Kind, v: unknown): { book: BookInput; cache: Cache } | Applied {
  const n = object(v, 'op.cache', 'a cache such as { field: "Region", pivots: [{ sheet: "P", name: "PivotTable1" }] }')
  if (kind === 'slicers' && 'table' in n) return { book, cache: tableCache(book, n) }
  onlyKeys(n, 'op.cache', ['field', 'pivots'])
  const field = checkString(n.field, 'op.cache.field')
  const { refs, cache: id } = pivotsOf(book, n.pivots, 'op.cache.pivots')
  const pc = book.pivotCaches?.find((c) => c.id === id)
  const f = pc ? fieldIndex(pc, field) : -1
  if (!pc || f < 0) return refused(book, { key: 'slicerField', field })
  const fieldName = pc.fields[f].name
  const name = cacheName(book, kind, fieldName)
  if (kind === 'timelines') {
    if (!pc.fields[f].facts?.types?.includes('date') && !pc.fields[f].items?.some((x) => x !== null && typeof x === 'object' && 'date' in x)) return refused(book, { key: 'timelineField', field: fieldName })
    const bounds = timelineBounds(pc, f)
    return { book, cache: { name, field: fieldName, pivots: refs, cache: id, ...(bounds ? { bounds } : {}) } }
  }
  const at = slicerField(book, { field: fieldName, cache: id })!
  const items = itemsShown(at.book, refs[0], at.field, groupedItems(at.cache.fields[at.field])?.length ?? 0)
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
  const name = given.name === undefined ? free(names(out), cache.field, ' ') : givenName(out, given.name, `op.${key}.name`)
  if (typeof name !== 'string') return refused(out, name)
  const place = given.place === undefined ? placeBeside(out, kind, s, sourceArea(out, cache)) : undefined
  // Excel captions it by its field
  const item = { caption: cache.field, ...given, name, cache: cache.name, ...(place ? { place } : {}) } as Slicer & Timeline
  const sheets = [...out.sheets]
  sheets[s] = withField(sheets[s], kind, [...((sheets[s][kind] ?? []) as (Slicer | Timeline)[]), item] as never)
  out = { ...out, sheets }
  // its choice is the filter its PivotTables hold already: a page field keeps its one item (the set slicer-held)
  if (fresh) out = withCaches(out, kind, [...cachesOf(out, kind), cache])
  // a new slicer filters nothing yet: its PivotTables are laid out again only to learn its items' data
  return laidOut(checkedPivots(out), cache.pivots, values, false)
}

function edit(book: BookInput, kind: Kind, o: Record<string, unknown>, values: ValuesIn): Applied {
  const key = KEY[kind]
  const { s, k, x } = find(book, kind, o[key], `op.${key}`)
  let out = book
  if (o.set !== undefined) {
    const set = object(o.set, 'op.set', `a patch of a ${key} such as { caption: "Region" }`)
    if ('cache' in set) throw new InputError('op.set.cache', `a ${key}'s cache is not edited`)
    // Slicer Settings' Name: unique among the workbook's slicers and timelines, its caption as it was
    const name = set.name === undefined ? undefined : givenName(book, set.name, 'op.set.name', x.name)
    if (typeof name === 'object') return refused(book, name)
    const list = [...((book.sheets[s][kind] ?? []) as (Slicer | Timeline)[])]
    // a rename keeps the caption, which a slicer without one took from its name
    list[k] = patched(set.name !== undefined && list[k].caption === undefined ? { ...list[k], caption: list[k].name } : list[k], set)
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
    checkItems(patch.items, 'op.cache.items', groupedItems(at?.cache.fields[at.field])?.length ?? 0)
  }
  if (kind === 'timelines' && patch.selection != null) {
    const sel = object(patch.selection, 'op.cache.selection', 'dates such as { start: "2024-01-01T00:00:00", end: "2024-12-31T00:00:00" }')
    onlyKeys(sel, 'op.cache.selection', ['start', 'end'])
    checkIsoDate(sel.start, 'op.cache.selection.start', out.dateSystem)
    checkIsoDate(sel.end, 'op.cache.selection.end', out.dateSystem)
    if (String(sel.start) > String(sel.end)) throw new InputError('op.cache.selection', 'expected a start on or before the end')
  }
  const next = patched(c, patch)
  // the PivotTables it filters now take its choice; one no longer connected shows every item again (the set slicer-held)
  const choice = patch.items !== undefined || patch.selection !== undefined
  const added = next.pivots.filter((r) => !c.pivots.some((q) => sameRef(q, r)))
  const removed = c.pivots.filter((r) => !next.pivots.some((q) => sameRef(q, r)))
  out = withCaches(out, kind, caches.map((y) => (y === c ? next : y)))
  if (kind === 'timelines' && patch.selection !== undefined) out = withLevels(out, next as TimelineCache)
  if (choice || added.length) out = chosen(out, kind, next, choice ? next.pivots : added)
  if (removed.length) out = chosen(out, kind, { ...next, items: undefined, selection: undefined } as Cache, removed)
  const refs = [...c.pivots, ...added]
  return laidOut(checkedPivots(out), refs, values, choice || added.length + removed.length > 0)
}

/** The book with the timelines of `tc` at the level its dates are chosen at, none without a selection. */
function withLevels(book: BookInput, tc: TimelineCache): BookInput {
  const level = tc.selection && selectionLevel(tc.selection)
  const sheets = book.sheets.map((sh) =>
    sh.timelines?.some((t) => caseKey(t.cache) === caseKey(tc.name)) ? { ...sh, timelines: sh.timelines.map((t) => (caseKey(t.cache) !== caseKey(tc.name) ? t : patched(t, { selectionLevel: level ?? null }))) } : sh,
  )
  return { ...book, sheets }
}

function remove(book: BookInput, kind: Kind, o: Record<string, unknown>, values: ValuesIn): Applied {
  const { s, k, x } = find(book, kind, o[KEY[kind]], `op.${KEY[kind]}`)
  const list = ((book.sheets[s][kind] ?? []) as (Slicer | Timeline)[]).filter((_, i) => i !== k)
  const sheets = [...book.sheets]
  sheets[s] = withField(sheets[s], kind, (list.length ? list : undefined) as never)
  const out: BookInput = { ...book, sheets }
  const last = !sheets.some((sh) => (sh[kind] as (Slicer | Timeline)[] | undefined)?.some((y) => caseKey(y.cache) === caseKey(x.cache)))
  if (!last) return { ok: true, book: out, geometry: [] }
  // its cache goes with it: the PivotTables keep a slicer's hidden items, and lose a timeline's dates (the set slicer-delete)
  const gone = cachesOf(book, kind).find((c) => caseKey(c.name) === caseKey(x.cache))!
  const left = withCaches(out, kind, cachesOf(book, kind).filter((c) => c !== gone))
  if (kind === 'slicers' || !gone.pivots.length) return { ok: true, book: left, geometry: [] }
  return laidOut(checkedPivots(withTimelineChoice(left, { ...(gone as TimelineCache), selection: undefined })), gone.pivots, values, true)
}

export function slicerOp(book: BookInput, o: Record<string, unknown>, values: ValuesIn): Applied {
  const op = String(o.op)
  const kind: Kind = op.endsWith('Slicer') ? 'slicers' : 'timelines'
  return op.startsWith('add') ? add(book, kind, o, values) : op.startsWith('edit') ? edit(book, kind, o, values) : remove(book, kind, o, values)
}

/**
 * The book without the table slicer caches whose table or column is gone, and their slicers, as Excel saves the book
 * after deleting the table or the column, or converting the table to a range (the set `table-slicer`).
 */
export function withoutLostTableSlicers(book: BookInput): BookInput {
  if (!book.slicerCaches?.some((c) => c.table)) return book
  const tables = tablesOf(book)
  const lost = new Set(
    book.slicerCaches.flatMap((c) => {
      if (!c.table) return []
      const t = tables.byName.get(tableKey(c.table.name))
      return t && t.input.columns.some((x) => caseKey(x.name) === caseKey(c.table!.column)) ? [] : [caseKey(c.name)]
    }),
  )
  if (!lost.size) return book
  const sheets = book.sheets.map((sh) => (sh.slicers?.some((x) => lost.has(caseKey(x.cache))) ? withField(sh, 'slicers', sh.slicers.filter((x) => !lost.has(caseKey(x.cache))) as never) : sh))
  return withCaches({ ...book, sheets }, 'slicers', book.slicerCaches.filter((c) => !lost.has(caseKey(c.name))))
}
