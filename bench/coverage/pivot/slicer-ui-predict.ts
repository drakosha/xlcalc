/**
 * The engine's prediction of the sets `slicer-labels`, `slicer-objects` and `slicer-protect`, written before Excel's
 * answers are read. Each case's PivotTable made as `slicer-predict.ts` makes it (played on the engine alone over the rows
 * of `labels`, or a book of `basic`), its slicers and timelines added through a session's `editStructure`; then:
 * - `slicer-labels`: each `slicerGet`, the slicer's items in its order, each its item's text and its label as
 *   `getLayout` gives it in the locale (`en` or `ru`);
 * - `slicer-objects`: each rename taken or refused by the rules of `editSlicer`'s name (not empty, at most 255
 *   characters, unique among the workbook's slicers and timelines in any case), the caption kept; `Locked` and
 *   `DisableMoveResizeUI` read back as set; each new slicer's or timeline's place as `addSlicer` gives it;
 * - `slicer-protect`: Microsoft's rule taken by the host so far: a choice needs Use PivotTable reports, moving or
 *   resizing a locked slicer or timeline needs Edit objects, an unlocked one moves without it.
 * Run: `npx tsx bench/coverage/pivot/slicer-ui-predict.ts <set> [en|ru]`, which writes
 * `bench/coverage/pivot/<set>[.ru-RU].predicted.json`.
 */
import { writeFileSync } from 'node:fs'
import { createSession, registerLocale, registerPivot, type BookInput, type PivotRef, type Session, type StructureOp } from '../../../src/index.js'
import { ruRU } from '../../../src/locales/ru-RU.js'
import { pivotEngine } from '../../../src/pivot/index.js'
import { itemText } from '../../../src/pivot/values.js'
import { groupedItems } from '../../../src/pivot/items.js'
import { fieldIndex } from '../../../src/structure/slicer-choice.js'
import type { PivotCase, PivotSet } from './cases.js'
import { start } from './slicer-predict.js'

type Answer = Record<string, unknown>

/** The cache field of a slicer over a PivotTable field as Excel names it (`Months (Long)` is the field `Long` grouped). */
function cacheField(book: BookInput, name: string): string {
  const pc = book.pivotCaches![0]
  if (fieldIndex(pc, name) >= 0) return name
  const base = /^\w+ \((.+)\)$/.exec(name)?.[1]
  return base ?? name
}

/** A slicer's items in its order, [text, label] each. */
function items(s: Session, name: string): [string, string][] {
  const layout = s.getLayout()
  const own = layout.sheets.flatMap((x) => x.slicers ?? []).find((x) => x.name === name)!
  const sc = layout.slicerCaches!.find((c) => c.name === own.cache)!
  const pc = layout.pivotCaches!.find((c) => c.id === sc.cache)!
  const values = groupedItems(pc.fields[fieldIndex(pc, sc.field)]) ?? []
  return (sc.items ?? []).map((x) => [x.item < values.length ? itemText(values[x.item]) : `#${x.item}`, x.label])
}

/** The rules of a new name: not empty, at most 255 characters, not another slicer's or timeline's in any case. */
function nameRefused(book: BookInput, from: string, to: string): string | null {
  if (!to.trim()) return 'empty'
  if (to.length > 255) return 'long'
  const others = book.sheets.flatMap((s) => [...(s.slicers ?? []), ...(s.timelines ?? [])]).filter((x) => x.name !== from)
  return others.some((x) => x.name.toLowerCase() === to.toLowerCase()) ? 'taken' : null
}

async function predictCase(c: PivotCase, locale: string): Promise<Answer> {
  const { book, rest } = await start(c)
  let s = createSession(locale === 'ru' ? { ...book, locale: 'ru-RU' } : book)
  const out: Answer = {}
  const edit = (k: number, op: StructureOp) => {
    const r = s.editStructure(op)
    if (!r.ok) out[`op ${k}`] = { refused: r.message }
  }
  // the slicers' and timelines' own state the session does not hold yet: names given, Locked, DisableMoveResizeUI, protection
  const own = new Map<string, Record<string, unknown>>()
  let guard = { objects: true, pivots: true, on: false }
  const kindOf = (name: string) => (s.getBook().sheets.some((x) => x.timelines?.some((t) => t.name === name)) ? 'timeline' : 'slicer')
  // each answer by the operation's index in the case, as Excel's answers have it
  rest.forEach((o) => {
    const k = c.ops.indexOf(o)
    if ('slicer' in o || 'timeline' in o) {
      const pivots: PivotRef[] = [{ sheet: 'P', name: o.pt ?? 'PT1' }]
      const sheet = o.on ?? 'P'
      const field = cacheField(s.getBook(), 'slicer' in o ? o.slicer : o.timeline)
      if ('slicer' in o) edit(k, { op: 'addSlicer', sheet, slicer: { name: o.name }, cache: { field, pivots } })
      else edit(k, { op: 'addTimeline', sheet, timeline: { name: o.name }, cache: { field, pivots } })
      own.set(o.name, { Locked: true, DisableMoveResizeUI: false })
    } else if ('slicerGet' in o) {
      const b = s.getBook()
      const x = b.sheets.flatMap((sh) => [...(sh.slicers ?? []), ...(sh.timelines ?? [])]).find((y) => y.name === o.slicerGet)
      if (!x) return void (out[`op ${k}`] = { refused: 'unknown' })
      const got: Answer = {}
      for (const p of o.get ?? []) got[p] = p === 'Name' ? x.name : p === 'Caption' ? (x.caption ?? null) : p === 'SlicerCache.Name' ? x.cache : p in own.get(x.name)! ? own.get(x.name)![p] : null
      if ((o.get ?? []).includes('Top')) got.place = x.place ?? null
      if (o.items?.length) got.items = items(s, x.name)
      out[`op ${k}`] = got
    } else if ('slicerSet' in o) {
      const b = s.getBook()
      const locked = own.get(o.slicerSet)?.Locked !== false
      if (o.prop === 'Name') {
        const why = nameRefused(b, o.slicerSet, String(o.to))
        if (why) return void (out[`op ${k}`] = { refused: why })
        own.set(String(o.to), own.get(o.slicerSet)!)
        // the rename the session takes once editSlicer and editTimeline take a name
        const sheets = b.sheets.map((sh) => ({ ...sh, ...(sh.slicers ? { slicers: sh.slicers.map((x) => (x.name === o.slicerSet ? { ...x, name: String(o.to) } : x)) } : {}), ...(sh.timelines ? { timelines: sh.timelines.map((x) => (x.name === o.slicerSet ? { ...x, name: String(o.to) } : x)) } : {}) }))
        s.close()
        s = createSession({ ...b, sheets })
      } else if (o.prop === 'Caption') edit(k, kindOf(o.slicerSet) === 'slicer' ? { op: 'editSlicer', slicer: o.slicerSet, set: { caption: String(o.to) } } : { op: 'editTimeline', timeline: o.slicerSet, set: { caption: String(o.to) } })
      else if (guard.on && !guard.objects && locked && (o.prop === 'Top' || o.prop === 'Width')) out[`op ${k}`] = { refused: 'objects' }
      else own.get(o.slicerSet)![o.prop] = o.to
    } else if ('protect' in o) guard = { objects: !!o.objects, pivots: !!o.pivots, on: true }
    else if ('unprotect' in o) guard = { ...guard, on: false }
    else if (('select' in o || 'timelineSet' in o) && guard.on && !guard.pivots) out[`op ${k}`] = { refused: 'pivots' }
    else if ('select' in o) {
      const b = s.getBook()
      const sc = b.slicerCaches!.find((x) => x.name === o.select)!
      const values = b.pivotCaches![0].fields[fieldIndex(b.pivotCaches![0], sc.field)].items!
      edit(k, { op: 'editSlicer', slicer: o.select, cache: { items: sc.items!.map((x) => ({ item: x.item, ...(o.items.includes(itemText(values[x.item])) ? {} : { selected: false }) })) } })
    }
  })
  return out
}

if (import.meta.url === `file://${process.argv[1]}`) {
  registerPivot(pivotEngine)
  registerLocale(ruRU)
  const [set, locale = 'en'] = process.argv.slice(2)
  const { buildPivotCases } = (await import(`./${set}.js`)) as PivotSet
  const out: Record<string, Answer> = {}
  // a case the engine throws on is predicted as its error
  for (const c of buildPivotCases()) out[c.id] = await predictCase(c, locale).catch((e: Error) => ({ throws: e.message }))
  writeFileSync(new URL(`${set}${locale === 'ru' ? '.ru-RU' : ''}.predicted.json`, import.meta.url), JSON.stringify(out, null, 1) + '\n')
  console.log(Object.keys(out).length, 'cases predicted')
}
