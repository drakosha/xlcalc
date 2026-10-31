/**
 * The workbook's PivotTables through operations of the structure: a sheet renamed gives its new name to the pivot
 * caches whose source is on it and to the slicers' and timelines' PivotTables on it; PivotTables, sources and anchors
 * move with their cells (`carryPivots`).
 */
import type { BookInput, NameInput, PivotCache, PivotRef, PivotRefusalReason, PivotSource, PivotTable, Slicer, StructureOp, Timeline } from '../contract.js'
import { parseAddress } from '../model.js'
import { caseKey } from '../collate.js'
import { sayPivot } from '../locale/pivot-messages.js'
import { withField } from '../looks/sheet.js'
import { quoteSheet, sheetKey } from '../syntax.js'
import { areaText, contains, meets, parseRange, type Rect } from './area.js'
import { pivotArea as areaOf } from '../pivot/area.js'
import type { Applied, Geometry } from './geometry.js'
import type { ValuesIn } from './values.js'
import { checkBook } from '../input.js'
import { pivotCheck, pivotRefresh } from '../pivot/hook.js'
import { sheetIndex } from './rewrite.js'

/** A source's sheet renamed; a source in another workbook keeps its sheet. */
function renamedSource(s: PivotSource, rename: (name: string) => string): PivotSource {
  if (s.type === 'range' && s.book === undefined && rename(s.sheet) !== s.sheet) return { ...s, sheet: rename(s.sheet) }
  if (s.type === 'consolidation' && s.ranges.some((r) => r.book === undefined && r.sheet !== undefined && rename(r.sheet) !== r.sheet))
    return { ...s, ranges: s.ranges.map((r) => (r.book === undefined && r.sheet !== undefined ? { ...r, sheet: rename(r.sheet) } : r)) }
  return s
}

const renamedRefs = (refs: PivotRef[], rename: (name: string) => string): PivotRef[] => (refs.some((r) => rename(r.sheet) !== r.sheet) ? refs.map((r) => ({ ...r, sheet: rename(r.sheet) })) : refs)

/** The book's caches, slicer and timeline caches with a sheet's new name, `rename` giving it; the book itself when none names it. */
export function renamePivotSheets(book: BookInput, rename: (name: string) => string): BookInput {
  if (!book.pivotCaches && !book.slicerCaches && !book.timelineCaches) return book
  const out = { ...book }
  if (book.pivotCaches) out.pivotCaches = book.pivotCaches.map((c): PivotCache => {
    const source = renamedSource(c.source, rename)
    return source === c.source ? c : { ...c, source }
  })
  if (book.slicerCaches) out.slicerCaches = book.slicerCaches.map((c) => ({ ...c, pivots: renamedRefs(c.pivots, rename) }))
  if (book.timelineCaches) out.timelineCaches = book.timelineCaches.map((c) => ({ ...c, pivots: renamedRefs(c.pivots, rename) }))
  return out
}

/** The prefix of the hidden names that carry PivotTables' places, caches' sources and slicers' and timelines' anchors through an operation. */
const CARRIER = '_xlcalc.pivot.'

const absolute = (sheet: string, x: Rect) => `${quoteSheet(sheet)}!${areaText(x).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`

const CARRIED = /^(?:'((?:[^']|'')+)'|([^'!]+))!([A-Z$0-9:]+)$/

/** A carrier's reference after the operation: its sheet and rectangle, or null when it was lost. */
function carried(ref: string): { sheet: string; x: Rect } | null {
  const m = CARRIED.exec(ref)
  const x = m && parseRange(m[3].replace(/\$/g, ''))
  return m && x ? { sheet: m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2], x } : null
}

/** The same size. */
const sameSize = (a: Rect, b: Rect) => a.r2 - a.r1 === b.r2 - b.r1 && a.c2 - a.c1 === b.c2 - b.c1

/** A rectangle moved by the rows and columns `d`. */
const moved = (x: Rect, dr: number, dc: number): Rect => ({ r1: x.r1 + dr, r2: x.r2 + dr, c1: x.c1 + dc, c2: x.c2 + dc })

/** The PivotTable at `p` moved by `dr` rows and `dc` columns. */
export const movedPivot = (p: PivotTable, dr: number, dc: number): PivotTable => ({ ...p, location: { ...p.location, range: areaText(moved(parseRange(p.location.range)!, dr, dc)) } })

/**
 * The book with what moves with cells carried by hidden names, and `after`, which reads them back from the book an
 * operation gave and strips them: a PivotTable moves with its cells, goes with all of them deleted, and refuses an
 * operation that cuts it (`pivot`); a cache's source on a sheet follows its cells as a reference does, and stays as it
 * was when they are deleted; a slicer's and a timeline's anchor cells follow theirs.
 */
export function carryPivots(book: BookInput): { book: BookInput; after(out: Applied & { ok: true }): Applied } | null {
  if (!book.pivotCaches?.length && !book.slicerCaches?.length && !book.timelineCaches?.length) return null
  const names: NameInput[] = []
  const add = (key: string, sheet: string, x: Rect) => names.push({ name: CARRIER + key, ref: absolute(sheet, x), hidden: true })
  book.sheets.forEach((s, i) => {
    s.pivotTables?.forEach((p, k) => add(`a${i}_${k}`, s.name, pivotArea(p)))
    for (const kind of ['slicers', 'timelines'] as const)
      s[kind]?.forEach((d, k) => {
        if (!d.place) return
        add(`${kind[0]}${i}_${k}f`, s.name, parseRange(d.place.from)!)
        add(`${kind[0]}${i}_${k}t`, s.name, parseRange(d.place.to)!)
      })
  })
  book.pivotCaches?.forEach((c, j) => {
    if (c.source.type === 'range' && c.source.book === undefined && sheetIndex(book, c.source.sheet) >= 0) add(`c${j}`, c.source.sheet, parseRange(c.source.range)!)
  })
  return {
    book: { ...book, names: [...(book.names ?? []), ...names] },
    after(out) {
      const at = new Map<string, string>()
      const rest = (out.book.names ?? []).filter((n) => !(n.name.startsWith(CARRIER) && n.hidden && at.set(n.name.slice(CARRIER.length), n.ref)))
      let b: BookInput = { ...out.book, names: rest.length ? rest : undefined }
      if (!book.names) delete b.names
      // a reference to a deleted sheet is lost too (`#REF!A1`)
      const now = (key: string) => {
        const n = carried(at.get(key) ?? '')
        return n && sheetIndex(b, n.sheet) >= 0 ? n : null
      }
      // the PivotTables: moved, gone with their cells, or cut
      const gone: (PivotRef & { cache: number })[] = []
      const arrivals = new Map<number, PivotTable[]>()
      const renamed = new Map<string, PivotRef>()
      for (const [i, s] of book.sheets.entries())
        for (const [k, p] of (s.pivotTables ?? []).entries()) {
          const was = pivotArea(p)
          const n = now(`a${i}_${k}`)
          if (!n) {
            gone.push({ sheet: s.name, name: p.name, cache: p.cache })
            continue
          }
          if (!sameSize(was, n.x)) return { ok: false, code: 'pivot', message: sayPivot(book.locale, { key: 'pivotCut', sheet: s.name, pivot: p.name }), reason: 'affect' }
          const to = sheetIndex(b, n.sheet)
          const list = arrivals.get(to) ?? arrivals.set(to, []).get(to)!
          list.push(movedPivot(p, n.x.r1 - was.r1, n.x.c1 - was.c1))
          if (sheetKey(n.sheet) !== sheetKey(b.sheets[sheetIndex(b, s.name)]?.name ?? s.name)) renamed.set(`${sheetKey(s.name)}\0${caseKey(p.name)}`, { sheet: n.sheet, name: p.name })
        }
      b = {
        ...b,
        sheets: b.sheets.map((s, i) => {
          const list = arrivals.get(i)
          return list || s.pivotTables ? withField(s, 'pivotTables', list?.length ? list : undefined) : s
        }),
      }
      // the sources and the anchors
      if (b.pivotCaches)
        b.pivotCaches = b.pivotCaches.map((c, j) => {
          const n = now(`c${j}`)
          return n && c.source.type === 'range' ? { ...c, source: { ...c.source, sheet: n.sheet, range: areaText(n.x) } } : c
        })
      b = {
        ...b,
        sheets: b.sheets.map((s) => {
          const i = book.sheets.findIndex((o) => sheetKey(o.name) === sheetKey(s.name))
          if (i < 0) return s
          let out = s
          for (const kind of ['slicers', 'timelines'] as const)
            if (s[kind])
              out = withField(out, kind, (s[kind] as (Slicer | Timeline)[]).map((d, k) => {
                const [f, t] = [now(`${kind[0]}${i}_${k}f`), now(`${kind[0]}${i}_${k}t`)]
                return d.place && f && t ? { ...d, place: { ...d.place, from: areaText(f.x), to: areaText(t.x) } } : d
              }) as never)
          return out
        }),
      }
      const refs = (list: PivotRef[]) => list.map((r) => renamed.get(`${sheetKey(r.sheet)}\0${caseKey(r.name)}`) ?? r)
      if (b.slicerCaches) b.slicerCaches = b.slicerCaches.map((c) => ({ ...c, pivots: refs(c.pivots) }))
      if (b.timelineCaches) b.timelineCaches = b.timelineCaches.map((c) => ({ ...c, pivots: refs(c.pivots) }))
      for (const g of gone) b = withoutPivot(b, g, g.cache)
      return { ...out, book: b }
    },
  }
}

const isRef = (r: PivotRef) => (x: PivotRef) => sheetKey(x.sheet) === sheetKey(r.sheet) && caseKey(x.name) === caseKey(r.name)

/**
 * The book without the PivotTable `r` (its cells are the caller's): its rules go, and the slicers' and timelines' caches
 * no longer name it, each kept with its slicers, its choice and its pivot cache even when it names no PivotTable, as
 * Excel keeps them (the set `slicer-delete`); its pivot cache goes with the last PivotTable and cache over it.
 */
export function withoutPivot(book: BookInput, r: PivotRef, of?: number): BookInput {
  const is = isRef(r)
  const left = <C extends { pivots: PivotRef[] }>(list: C[] | undefined) => list?.map((c) => (c.pivots.some(is) ? { ...c, pivots: c.pivots.filter((x) => !is(x)) } : c))
  const slicerCaches = left(book.slicerCaches)
  const timelineCaches = left(book.timelineCaches)
  let cache = of
  const sheets = book.sheets.map((s) => {
    if (sheetKey(s.name) !== sheetKey(r.sheet)) return s
    cache ??= s.pivotTables?.find((p) => caseKey(p.name) === caseKey(r.name))?.cache
    const list = s.pivotTables?.filter((p) => caseKey(p.name) !== caseKey(r.name))
    let out = withField(s, 'pivotTables', list?.length ? list : undefined)
    const rules = s.conditionalFormats?.filter((f) => !(f.pivot && caseKey(f.pivot.pivot) === caseKey(r.name)))
    if (rules && rules.length !== s.conditionalFormats!.length) out = withField(out, 'conditionalFormats', rules.length ? rules : undefined)
    return out
  })
  const out: BookInput = { ...book, sheets }
  if (slicerCaches) out.slicerCaches = slicerCaches
  if (timelineCaches) out.timelineCaches = timelineCaches
  const used = cache !== undefined && (sheets.some((s) => s.pivotTables?.some((p) => p.cache === cache)) || slicerCaches?.some((c) => c.cache === cache) || timelineCaches?.some((c) => c.cache === cache))
  if (cache !== undefined && !used) {
    const caches = book.pivotCaches?.filter((c) => c.id !== cache)
    if (caches?.length) out.pivotCaches = caches
    else delete out.pivotCaches
  }
  return out
}

/** The PivotTables of a sheet with their areas. */
export const pivotsOn = (book: BookInput, s: number) => (book.sheets[s]?.pivotTables ?? []).map((p) => ({ p, x: pivotArea(p) }))

/** The refusal of an area of sheet `s` that meets a PivotTable, but those `but` names, with Excel's alert; null when none does. */
export function pivotMet(book: BookInput, s: number, x: Rect | null, reason: PivotRefusalReason | undefined, but?: (p: PivotTable) => boolean): Exclude<Applied, { ok: true }> | null {
  const hit = x && pivotsOn(book, s).find((t) => meets(t.x, x) && !but?.(t.p))
  return hit ? { ok: false, code: 'pivot', message: sayPivot(book.locale, { key: 'pivotCut', sheet: book.sheets[s].name, pivot: hit.p.name }), ...(reason ? { reason } : {}) } : null
}

/**
 * What Excel refuses of an operation over PivotTables before it runs: cells written onto one (a paste, a fill, a merge,
 * a table), a move that takes part of one or lands on one it does not take whole. Moving every cell of one moves it.
 */
export function pivotRefusal(book: BookInput, op: StructureOp): Exclude<Applied, { ok: true }> | null {
  if (!book.sheets.some((s) => s.pivotTables?.length)) return null
  const at = (sheet: string) => sheetIndex(book, sheet)
  switch (op.op) {
    // a paste, a fill or a merge onto one is Excel's alert of a part changed; a table over one has an alert of its own
    case 'copyRange':
      return pivotMet(book, at(op.to.sheet), parseRange(op.to.range), 'part')
    case 'fillRange':
    case 'mergeCells':
      return pivotMet(book, at(op.sheet), parseRange(op.op === 'fillRange' ? op.to : op.range), 'part')
    case 'addTable':
      return pivotMet(book, at(op.sheet), typeof op.table?.range === 'string' ? parseRange(op.table.range) : null, undefined)
    case 'moveRange': {
      const src = parseRange(op.from.range)
      const to = parseAddress(op.to.cell)
      if (!src) return null
      const s = at(op.from.sheet)
      // a PivotTable the move takes whole moves with it; one it takes part of is cut
      const whole = (p: PivotTable) => contains(src, pivotArea(p))
      const cut = pivotsOn(book, s).find((t) => meets(t.x, src) && !contains(src, t.x))
      // one the move takes part of is affected; one it lands on has a part changed
      if (cut) return pivotMet(book, s, cut.x, 'affect')
      const dest = { r1: to.row, c1: to.col, r2: to.row + src.r2 - src.r1, c2: to.col + src.c2 - src.c1 }
      return pivotMet(book, at(op.to.sheet), dest, 'part', (p) => at(op.to.sheet) === s && whole(p))
    }
    default:
      return null
  }
}

/** The patch's fields set on `target`: a value replaces, null removes. */
export function patched<T extends object>(target: T, patch: Record<string, unknown>): T {
  const out: Record<string, unknown> = { ...(target as Record<string, unknown>) }
  for (const [k, v] of Object.entries(patch)) {
    if (v === null) delete out[k]
    else if (v !== undefined) out[k] = v
  }
  return out as T
}

/**
 * The PivotTables `refs` laid out again by `xlcalc/pivot` (`fromSource`: their caches refreshed from the sources); without
 * it, their caches marked `stale`, the cells left.
 */
export function relaid(book: BookInput, refs: PivotRef[], values: ValuesIn, fromSource: boolean, geometry: Geometry = [], blocked: 'spill' | 'refuse' = 'spill', labelSort?: readonly number[]): Applied {
  const refresh = pivotRefresh()
  const r = refresh?.(book, refs, values, values.today?.() ?? 0, fromSource, blocked, labelSort)
  if (r && !r.ok) return r
  // laid out, unless the cache keeps no rows to lay it out from
  if (r && r.book !== book) return { ...r, geometry: [...geometry, ...r.geometry] }
  const caches = new Set(refs.flatMap((r) => book.sheets[sheetIndex(book, r.sheet)]?.pivotTables?.filter((p) => caseKey(p.name) === caseKey(r.name)).map((p) => p.cache) ?? []))
  const pivotCaches = book.pivotCaches?.map((c) => (caches.has(c.id) && !c.stale ? { ...c, stale: true } : c))
  return { ok: true, book: pivotCaches ? { ...book, pivotCaches } : book, geometry, stale: refs }
}

/** The book checked as input, the full check of PivotTables too when `xlcalc/pivot` registered it. */
export function checkedPivots(book: BookInput, path = 'op'): BookInput {
  const out = checkBook(book, path)
  pivotCheck()?.(out, path)
  return out
}

/** A PivotTable's cells with its Filters area (one the model holds, whose range parses). */
export const pivotArea = (p: PivotTable): Rect => areaOf(p)!
