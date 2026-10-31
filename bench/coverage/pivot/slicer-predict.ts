/**
 * The engine's prediction of the set `slicer`, written before Excel's answers are read: each case's PivotTables made
 * from a book of the set `basic` Excel saved (`price-sum`, `date-rows`), the second PivotTable, the page field and the
 * hidden items put in its model, then the slicers and timelines added and chosen through a session's `editStructure`
 * with `xlcalc/pivot` loaded. Each phase gives the cells of every PivotTable's area and each slicer's items as the
 * harness reads them (name, selected, has data, in the slicer's order) or a timeline's dates and level. Run:
 * `npx tsx bench/coverage/pivot/slicer-predict.ts <set> <out.json>`. `slicer.predicted.json` is the prediction made before the
 * answers were read: 31 of 38 held, the 7 others a blank region whose one record has no price, which Excel tells has no
 * data. `slicer-delete.predicted.json` likewise for that set's cases a session takes; `slicer-held.predicted.json`
 * for the held-out set over the rows of `stock`, its PivotTables played on the engine alone (`drive.ts`). The comparison with Excel:
 * `test/pivot-slicers-excel.test.ts`.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { createSession, registerPivot, type BookInput, type CellValue, type PivotRef, type StructureOp } from '../../../src/index.js'
import { dateToSerial, formatAddress } from '../../../src/model.js'
import { pivotArea } from '../../../src/pivot/area.js'
import { pivotEngine } from '../../../src/pivot/index.js'
import { refreshPivots } from '../../../src/pivot/engine.js'
import { itemText, labelOrder } from '../../../src/pivot/values.js'
import { withListed } from '../../../src/pivot/items.js'
import { areaText } from '../../../src/structure/area.js'
import { fieldIndex } from '../../../src/structure/slicer-choice.js'
import { ERROR_CODE } from '../../../src/syntax/errors.js'
import { readXlsx } from '../../../src/xlsx.js'
import type { PivotCase, PivotOp, PivotPhase, PivotSet, SlicerItem } from './cases.js'
import { drive } from './drive.js'

const ROOT = new URL('../../../', import.meta.url)
const LEVELS = ['years', 'quarters', 'months', 'days'] as const

/** A PivotTable's area as predicted: its range with the Filters area and the values of its cells. */
export interface PredictedArea {
  area: string
  values: CellValue[][]
}

export interface PredictedPhase {
  pivots: Record<string, PredictedArea>
  slicers: NonNullable<PivotPhase['slicers']>
}

const load = async (name: string): Promise<BookInput> => (await readXlsx(readFileSync(new URL(`test/excel/win/pivot/basic/${name}.xlsx`, ROOT)))).book

/** The values of a PivotTable's area as the harness reads them: dates as serials, errors as `{ error }`. */
function areaOf(book: BookInput, name: string): PredictedArea {
  const sheet = book.sheets.find((s) => s.name === 'P')!
  const x = pivotArea(sheet.pivotTables!.find((p) => p.name === name)!)!
  const values = Array.from({ length: x.r2 - x.r1 + 1 }, (_, r) =>
    Array.from({ length: x.c2 - x.c1 + 1 }, (_, c): CellValue => {
      const cell = sheet.cells[formatAddress(x.r1 + r, x.c1 + c)]
      if (cell?.formula?.startsWith('#')) return { error: ERROR_CODE[cell.formula] }
      const v = cell?.value
      return v === undefined ? null : typeof v === 'object' && v !== null && 'date' in v ? dateToSerial(v.date) : (v as CellValue)
    }),
  )
  return { area: areaText(x), values }
}

/** The PivotTable `name`'s model changed by `f`, its cache by `c`, and laid out again from the cache. */
function changed(book: BookInput, name: string, f: (p: NonNullable<BookInput['sheets'][number]['pivotTables']>[number], book: BookInput) => typeof p, add = false): BookInput {
  const sheets = book.sheets.map((s) => (s.name !== 'P' ? s : { ...s, pivotTables: add ? [...s.pivotTables!, f(s.pivotTables![0], book)] : s.pivotTables!.map((p) => (p.name === name ? f(p, book) : p)) }))
  const r = refreshPivots({ ...book, sheets }, [{ sheet: 'P', name }], () => [], 0, false)
  if (!r.ok) throw new Error(`${name}: ${r.message}`)
  return r.book
}

/** The operations a session plays after the PivotTables are made: slicers, timelines, their choices and settings, the phases' ends. */
const slicing = (o: PivotOp) => 'slicer' in o || 'select' in o || 'timeline' in o || 'read' in o || 'deleteSlicer' in o || 'slicerSettings' in o || 'link' in o

/**
 * The book of a case before its slicers: the PivotTables and fields its first operations make, over a book of `basic`
 * for the rows of `sales`, or played on the engine alone (`drive.ts`) for other rows; the page field, its item and the
 * hidden items put in the model.
 */
async function start(c: PivotCase): Promise<{ book: BookInput; rest: PivotOp[] }> {
  const dates = 'rows' in c.ops[0] && c.ops[0].rows[0] === 'Date'
  const own = (o: PivotOp) => 'pages' in o || 'page' in o || 'hide' in o
  let book = c.data ? drive({ ...c, ops: c.ops.filter((o) => !slicing(o) && !own(o)) }).at(-1)!.book : await load(dates ? 'date-rows' : 'price-sum')
  const cache = () => book.pivotCaches![0]
  const at = (field: string) => fieldIndex(cache(), field)
  /** The field listed in the cache and on PT1 in Excel's order of a field placed: ascending. */
  const listedOn = (f: number) => {
    book = { ...book, pivotCaches: [withListed(cache(), [f])] }
    const values = cache().fields[f].items!
    return values.map((_, k) => k).sort((a, b) => labelOrder(values[a], values[b], false, null))
  }
  const rest: PivotOp[] = []
  for (const o of c.ops) {
    if (slicing(o)) rest.push(o)
    else if ('hide' in o) {
      const f = at(o.hide)
      book = changed(book, 'PT1', (p) => ({ ...p, fields: p.fields.map((pf, i) => (i !== f ? pf : { ...pf, items: pf.items!.map((x) => (o.items.includes(itemText(cache().fields[f].items![x.item!])) ? { ...x, hidden: true } : x)) })) }))
    } else if ('pages' in o) {
      const fs = o.pages.map(at)
      const orders = fs.map(listedOn)
      book = changed(book, 'PT1', (p) => ({ ...p, pages: fs.map((field) => ({ field })), fields: p.fields.map((pf, i) => (fs.includes(i) ? { ...pf, items: orders[fs.indexOf(i)].map((k) => ({ item: k })) } : pf)) }))
    } else if ('page' in o) {
      const f = at(o.page)
      const order = listedOn(f)
      const item = order.findIndex((k) => itemText(cache().fields[f].items![k]) === (o.items as string[])[0])
      book = changed(book, 'PT1', (p) => ({ ...p, pages: [...(p.pages ?? []).filter((x) => x.field !== f), { field: f, item }], fields: p.fields.map((pf, i) => (i !== f ? pf : { ...pf, items: order.map((k) => ({ item: k })) })) }))
    } else if (c.data) continue
    else if ('create' in o)
      book = changed(book, o.create, (p, b) => ({ name: o.create, cache: p.cache, location: { range: `${o.at.split('!')[1]}:${o.at.split('!')[1]}`, firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 }, fields: b.pivotCaches![0].fields.map(() => ({})) }), true)
    else if ('rows' in o && o.pt) book = changed(book, o.pt, (p) => ({ ...p, rows: o.rows.map(at) }))
    else if ('value' in o && o.pt) book = changed(book, o.pt, (p) => ({ ...p, values: [{ field: at(o.value.field), name: `Sum of ${o.value.field}` }] }))
  }
  return { book, rest }
}

/** A slicer's items as the harness reads them, or a timeline's dates and level. */
function slicersOf(book: BookInput): PredictedPhase['slicers'] {
  const out: PredictedPhase['slicers'] = {}
  const sheet = book.sheets.find((s) => s.name === 'P')!
  for (const s of sheet.slicers ?? []) {
    const sc = book.slicerCaches!.find((c) => c.name === s.cache)!
    const values = book.pivotCaches!.find((c) => c.id === sc.cache)!.fields[fieldIndex(book.pivotCaches![0], sc.field)].items!
    // Excel's HasData is false for every item of a slicer without cross filtering (the set slicer-held)
    out[s.name] = (sc.items ?? []).map((x): SlicerItem => [itemText(values[x.item]), x.selected !== false, sc.noData?.mark !== false && !x.noData])
  }
  for (const t of sheet.timelines ?? []) {
    const tc = book.timelineCaches!.find((c) => c.name === t.cache)!
    const serial = (iso?: string) => (iso === undefined ? null : Math.round(Date.parse(`${iso}Z`) / 86_400_000) + 25_569)
    out[t.name] = { timeline: [serial(tc.selection?.start), serial(tc.selection?.end), LEVELS.indexOf(t.level ?? 'months')] }
  }
  return out
}

const isoOf = (serial: number) => new Date((serial - 25_569) * 86_400_000).toISOString().slice(0, 19)

export async function predict(c: PivotCase): Promise<PredictedPhase[]> {
  const { book, rest } = await start(c)
  const s = createSession(book)
  const phases: PredictedPhase[] = []
  const edit = (op: StructureOp) => {
    const r = s.editStructure(op)
    if (!r.ok) throw new Error(`${c.id}: ${r.message}`)
  }
  const read = () => {
    const b = s.getBook()
    const pivots = Object.fromEntries(b.sheets.find((x) => x.name === 'P')!.pivotTables!.map((p) => [p.name, areaOf(b, p.name)]))
    phases.push({ pivots, slicers: slicersOf(b) })
  }
  for (const o of rest) {
    if ('read' in o) read()
    else if ('slicer' in o || 'timeline' in o) {
      // Excel's SlicerCaches.Add2 over the operation's PivotTable, then the others named
      const pivots: PivotRef[] = [o.pt ?? 'PT1', ...('pivots' in o ? (o.pivots ?? []) : [])].map((name) => ({ sheet: 'P', name }))
      if ('slicer' in o) edit({ op: 'addSlicer', sheet: 'P', slicer: { name: o.name }, cache: { field: o.slicer, pivots } })
      else {
        edit({ op: 'addTimeline', sheet: 'P', timeline: { name: o.name, ...(o.level ? { level: o.level } : {}) }, cache: { field: o.timeline, pivots } })
        if (o.from !== undefined && o.to !== undefined) edit({ op: 'editTimeline', timeline: o.name, cache: { selection: { start: isoOf(o.from), end: isoOf(o.to) } } })
      }
    } else if ('deleteSlicer' in o) {
      // SlicerCache.Delete: every slicer or timeline of the cache
      const b = s.getBook()
      for (const sh of b.sheets) for (const kind of ['slicers', 'timelines'] as const) for (const x of sh[kind] ?? []) if (x.cache === o.deleteSlicer || x.name === o.deleteSlicer) edit(kind === 'slicers' ? { op: 'deleteSlicer', slicer: x.name } : { op: 'deleteTimeline', timeline: x.name })
    } else if ('slicerSettings' in o) {
      // Excel's order of the data source has no word in the cache: taken as ascending
      const noData = { none: { mark: false }, last: {}, dimmed: { last: false }, hidden: { hide: true } } as const
      edit({ op: 'editSlicer', slicer: o.slicerSettings, cache: { ...(o.order && o.order !== 'source' ? { sort: o.order } : {}), ...(o.noData ? { noData: noData[o.noData] } : {}) } })
    } else if ('link' in o) {
      const b = s.getBook()
      const timeline = b.sheets.some((sh) => sh.timelines?.some((t) => t.name === o.link))
      const own = timeline ? b.sheets.flatMap((sh) => sh.timelines ?? []).find((t) => t.name === o.link)! : b.sheets.flatMap((sh) => sh.slicers ?? []).find((t) => t.name === o.link)!
      const was = (timeline ? b.timelineCaches : b.slicerCaches)!.find((x) => x.name === own.cache)!.pivots
      const ref = { sheet: 'P', name: o.pivot }
      const pivots = o.off ? was.filter((r) => r.name !== o.pivot) : [...was, ref]
      edit(timeline ? { op: 'editTimeline', timeline: o.link, cache: { pivots } } : { op: 'editSlicer', slicer: o.link, cache: { pivots } })
    } else if ('select' in o) {
      const b = s.getBook()
      const sc = b.slicerCaches!.find((x) => x.name === o.select)!
      const values = b.pivotCaches![0].fields[fieldIndex(b.pivotCaches![0], sc.field)].items!
      edit({ op: 'editSlicer', slicer: o.select, cache: { items: sc.items!.map((x) => ({ item: x.item, ...(o.items.includes(itemText(values[x.item])) ? {} : { selected: false }) })) } })
    }
  }
  read()
  return phases
}

/** The cases of a set this script predicts: those whose operations a session takes (no sheet added or deleted). */
export const predicted = (cases: PivotCase[]) => cases.filter((c) => !c.ops.some((o) => 'addSheet' in o || 'deleteSheet' in o))

if (import.meta.url === `file://${process.argv[1]}`) {
  registerPivot(pivotEngine)
  const out: Record<string, PredictedPhase[]> = {}
  const set = (await import(`./${process.argv[2]}.js`)) as PivotSet
  for (const c of predicted(set.buildPivotCases())) out[c.id] = await predict(c)
  writeFileSync(process.argv[3], JSON.stringify(out, null, 1) + '\n')
  console.log(Object.keys(out).length, 'cases predicted')
}
