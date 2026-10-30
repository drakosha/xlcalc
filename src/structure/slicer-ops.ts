/**
 * Operations on slicers and timelines (`editStructure`): one added over a cache of the workbook or a new cache of a
 * field over PivotTables of one pivot cache, changed with its cache (the items or dates chosen, the PivotTables it
 * filters, its look and place), deleted with its cache when it is the last over it. A choice changes what the
 * PivotTables show: they are laid out again from their cache, or marked `stale` (`relaid`).
 */
import type { BookInput, DrawingPlace, PivotRef, Slicer, SlicerCache, Timeline, TimelineCache } from '../contract.js'
import { caseKey } from '../collate.js'
import { InputError } from '../input-error.js'
import { checkSheetOf } from '../input.js'
import { array, checkString, object, onlyKeys } from '../input/basic.js'
import { say } from '../locale/messages.js'
import { withField } from '../looks/sheet.js'
import { formatAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import type { Applied } from './geometry.js'
import { checkedPivots, patched, pivotArea, relaid } from './pivots.js'
import { sheetIndex } from './rewrite.js'
import type { ValuesIn } from './values.js'

type Kind = 'slicers' | 'timelines'
type Caches = 'slicerCaches' | 'timelineCaches'
const CACHES: Record<Kind, Caches> = { slicers: 'slicerCaches', timelines: 'timelineCaches' }

/** The fields of a cache an edit sets, by kind. */
const CACHE_FIELDS: Record<Kind, string[]> = { slicers: ['pivots', 'items', 'sort', 'customListSort', 'noData', 'showMissing'], timelines: ['pivots', 'selection'] }


/** Every slicer's and timeline's name, which share the workbook's names. */
const names = (book: BookInput) => new Set(book.sheets.flatMap((s) => [...(s.slicers ?? []), ...(s.timelines ?? [])].map((x) => caseKey(x.name))))

/** The first of `base`, `base 1`, `base 2`… (`sep` between) not in `taken`. */
function free(taken: Set<string>, base: string, sep: string): string {
  let name = base
  for (let n = 1; taken.has(caseKey(name)); n++) name = `${base}${sep}${n}`
  return name
}

/** A slicer or timeline by name: its sheet's index, its kind and index there. */
function find(book: BookInput, kind: Kind, v: unknown, path: string): { s: number; k: number; x: Slicer | Timeline } {
  const name = checkString(v, path)
  for (const [s, sheet] of book.sheets.entries()) {
    const k = (sheet[kind] ?? []).findIndex((x) => caseKey(x.name) === caseKey(name))
    if (k >= 0) return { s, k, x: sheet[kind]![k] }
  }
  throw new InputError(path, `unknown ${kind === 'slicers' ? 'slicer' : 'timeline'} ${JSON.stringify(name)}`)
}

/** A cache's PivotTables checked: of the workbook, all over one pivot cache, whose id it gives. */
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
  return { refs: refs.map((r) => r.ref), cache: refs[0].cache }
}

/** A new slicer's or timeline's place: right of the first PivotTable it filters, three columns and fourteen rows, as Excel puts one beside it. */
function placeBeside(book: BookInput, r: PivotRef): DrawingPlace {
  const p = book.sheets[sheetIndex(book, r.sheet)].pivotTables!.find((q) => caseKey(q.name) === caseKey(r.name))!
  const x = pivotArea(p)
  return { from: formatAddress(x.r1, x.c2 + 2), to: formatAddress(x.r1 + 14, x.c2 + 5) }
}

function add(book: BookInput, kind: Kind, o: Record<string, unknown>): Applied {
  const s = book.sheets.indexOf(checkSheetOf(book, o.sheet, 'op.sheet'))
  const key = kind === 'slicers' ? 'slicer' : 'timeline'
  const given = object(o[key], `op.${key}`, `a ${key} such as { caption: "Region" }`)
  const ck = CACHES[kind]
  let caches = (book[ck] ?? []) as (SlicerCache | TimelineCache)[]
  let cache: SlicerCache | TimelineCache
  if (typeof o.cache === 'string') {
    const c = caches.find((x) => caseKey(x.name) === caseKey(o.cache as string))
    if (!c) throw new InputError('op.cache', `unknown ${key} cache ${JSON.stringify(o.cache)}`)
    cache = c
  } else {
    const n = object(o.cache, 'op.cache', 'a cache such as { field: "Region", pivots: [{ sheet: "P", name: "PivotTable1" }] }')
    onlyKeys(n, 'op.cache', ['field', 'pivots'])
    const field = checkString(n.field, 'op.cache.field')
    const { refs, cache: id } = pivotsOf(book, n.pivots, 'op.cache.pivots')
    const pc = book.pivotCaches!.find((c) => c.id === id)!
    const f = pc.fields.findIndex((x) => caseKey(x.name) === caseKey(field))
    if (f < 0) return { ok: false, code: 'pivot', message: say(book.locale, { key: 'slicerField', field }) }
    const fieldName = pc.fields[f].name
    const taken = new Set(caches.map((c) => caseKey(c.name)))
    const name = free(taken, `${kind === 'slicers' ? 'Slicer' : 'NativeTimeline'}_${fieldName.replace(/[^\p{L}\p{N}_]/gu, '_')}`, '')
    cache =
      kind === 'slicers'
        ? { name, field: fieldName, pivots: refs, cache: id, ...(pc.fields[f].items ? { items: pc.fields[f].items!.map((_, i) => ({ item: i })) } : {}) }
        : { name, field: fieldName, pivots: refs, cache: id }
    caches = [...caches, cache]
  }
  const taken = names(book)
  const name = given.name === undefined ? free(taken, cache.field, ' ') : checkString(given.name, `op.${key}.name`)
  if (taken.has(caseKey(name))) return { ok: false, code: 'pivot', message: say(book.locale, { key: 'slicerNameTaken', name }) }
  const item = { ...given, name, cache: cache.name, ...(given.place === undefined ? { place: placeBeside(book, cache.pivots[0]) } : {}) } as Slicer & Timeline
  const sheets = [...book.sheets]
  sheets[s] = withField(sheets[s], kind, [...((sheets[s][kind] ?? []) as (Slicer | Timeline)[]), item] as never)
  return { ok: true, book: checkedPivots({ ...book, sheets, [ck]: caches }), geometry: [] }
}

function edit(book: BookInput, kind: Kind, o: Record<string, unknown>, values: ValuesIn): Applied {
  const key = kind === 'slicers' ? 'slicer' : 'timeline'
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
  let refs: PivotRef[] = []
  if (o.cache !== undefined) {
    const patch = object(o.cache, 'op.cache', `a patch of a ${key} cache such as { items: [...] }`)
    onlyKeys(patch, 'op.cache', CACHE_FIELDS[kind])
    const ck = CACHES[kind]
    const caches = (out[ck] ?? []) as (SlicerCache | TimelineCache)[]
    const c = caches.find((y) => caseKey(y.name) === caseKey(x.cache))!
    if (patch.pivots !== undefined && patch.pivots !== null) {
      const { cache } = pivotsOf(out, patch.pivots, 'op.cache.pivots')
      if (c.cache !== undefined && cache !== c.cache) throw new InputError('op.cache.pivots', "expected PivotTables of the cache's pivot cache")
    }
    const next = patched(c, patch)
    // the PivotTables it filtered and those it filters now are laid out again
    refs = [...c.pivots, ...next.pivots.filter((r) => !c.pivots.some((q) => sheetKey(q.sheet) === sheetKey(r.sheet) && caseKey(q.name) === caseKey(r.name)))]
    out = { ...out, [ck]: caches.map((y) => (y === c ? next : y)) }
  }
  out = checkedPivots(out)
  return refs.length ? relaid(out, refs, values, false) : { ok: true, book: out, geometry: [] }
}

function remove(book: BookInput, kind: Kind, o: Record<string, unknown>, values: ValuesIn): Applied {
  const key = kind === 'slicers' ? 'slicer' : 'timeline'
  const { s, k, x } = find(book, kind, o[key], `op.${key}`)
  const list = ((book.sheets[s][kind] ?? []) as (Slicer | Timeline)[]).filter((_, i) => i !== k)
  const sheets = [...book.sheets]
  sheets[s] = withField(sheets[s], kind, (list.length ? list : undefined) as never)
  const out: BookInput = { ...book, sheets }
  const ck = CACHES[kind]
  const last = !sheets.some((sh) => (sh[kind] as (Slicer | Timeline)[] | undefined)?.some((y) => caseKey(y.cache) === caseKey(x.cache)))
  if (!last) return { ok: true, book: out, geometry: [] }
  // its cache goes with it, and the PivotTables it filtered show every item again
  const caches = ((book[ck] ?? []) as (SlicerCache | TimelineCache)[]).filter((c) => caseKey(c.name) !== caseKey(x.cache))
  const gone = ((book[ck] ?? []) as (SlicerCache | TimelineCache)[]).find((c) => caseKey(c.name) === caseKey(x.cache))
  if (caches.length) out[ck] = caches as never
  else delete out[ck]
  return gone?.pivots.length ? relaid(out, gone.pivots, values, false) : { ok: true, book: out, geometry: [] }
}


export function slicerOp(book: BookInput, o: Record<string, unknown>, values: ValuesIn): Applied {
  const kind: Kind = String(o.op).endsWith('Slicer') ? 'slicers' : 'timelines'
  return o.op === 'addSlicer' || o.op === 'addTimeline' ? add(book, kind, o) : o.op === 'editSlicer' || o.op === 'editTimeline' ? edit(book, kind, o, values) : remove(book, kind, o, values)
}
