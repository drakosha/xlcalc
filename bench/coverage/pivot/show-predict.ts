/**
 * The engine's prediction of a set of Show Values As (`show`, `show-items`), written before Excel's answers are read:
 * each case's PivotTable laid out by `xlcalc/pivot`, and the values of its area (`P`), dates as serials, errors by code,
 * numbers to 15 significant digits. A case Excel has saved a book of is laid out again from the saved cache and
 * refreshed from its source (the book's cells on `P` left out); a case without one is built from its operations over
 * the rows of `data.ts` (`modelOf`). Run: `npx tsx bench/coverage/pivot/show-predict.ts <set> [out.json]`, written to
 * `bench/coverage/pivot/<set>.predicted.json` by default.
 */
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import type { BookInput, CellValue, PivotDataField, PivotFilter, PivotFunction, PivotShowAs, PivotTable } from '../../../src/contract.js'
import { createState } from '../../../src/engine.js'
import { dateToSerial, formatAddress } from '../../../src/model.js'
import { pivotArea } from '../../../src/pivot/area.js'
import { wordsOf } from '../../../src/pivot/captions.js'
import { refreshPivots } from '../../../src/pivot/engine.js'
import { itemText } from '../../../src/pivot/values.js'
import { clockOptions } from '../../../src/result.js'
import { stateValues } from '../../../src/structure/values.js'
import { ERROR_CODE } from '../../../src/syntax/errors.js'
import { readXlsx } from '../../../src/xlsx.js'
import type { PivotCase, PivotSet, ShowAs } from './cases.js'
import { DATA, sourceBook, sourceRange } from './data.js'

const ROOT = new URL('../../../', import.meta.url)

/** The values of the area of each PivotTable on `P`, laid out again from the saved cache or refreshed from the source. */
function laidOut(book: BookInput, fromSource: boolean): Record<string, { area: string; values: CellValue[][] }> | string {
  const values = stateValues(createState(book, clockOptions({ now: '2026-10-30', seed: 0 })))
  const r = refreshPivots(book, undefined, values, values.today!(), fromSource)
  if (!r.ok) return r.message
  const sheet = r.book.sheets.find((s) => s.name === 'P')!
  const out: Record<string, { area: string; values: CellValue[][] }> = {}
  for (const t of sheet.pivotTables ?? []) {
    const x = pivotArea(t)!
    const grid = Array.from({ length: x.r2 - x.r1 + 1 }, (_, r) =>
      Array.from({ length: x.c2 - x.c1 + 1 }, (_, c): CellValue => {
        const cell = sheet.cells[formatAddress(x.r1 + r, x.c1 + c)]
        if (cell?.formula?.startsWith('#')) return { error: ERROR_CODE[cell.formula] }
        const v = cell?.value
        if (v === undefined || v === null) return null
        if (typeof v === 'object' && 'date' in v) return dateToSerial(v.date)
        return typeof v === 'number' ? Number(v.toPrecision(15)) : (v as CellValue)
      }),
    )
    out[t.name] = { area: `${formatAddress(x.r1, x.c1)}:${formatAddress(x.r2, x.c2)}`, values: grid }
  }
  return out
}

/** The calculations whose names the file and the model spell differently. */
const SHOW: Partial<Record<ShowAs, PivotShowAs>> = { percentOfCol: 'percentOfColumn', percentOfParentCol: 'percentOfParentColumn' }

/** The filters of the cases, as the model has them. */
const FILTERS: Record<string, PivotFilter['type']> = { TopCount: 'count', CaptionDoesNotEqual: 'captionNotEqual' }

/**
 * A case's PivotTable built from its operations over its rows (`sales` by default) as a range, as the harness has Excel build it:
 * the fields placed and the value fields added, refreshed once for the items, then the items hidden or shown on the
 * Filters area, the base items, filters and sorts set by the items' and value fields' names.
 */
function modelOf(c: PivotCase): BookInput | string {
  // grouping is not calculated by the engine yet: such a case is left unpredicted
  if (c.ops.some((o) => 'group' in o)) return 'not predicted: the engine does not group yet'
  const cols = DATA[c.data ?? 'sales']()
  const names = cols.map((x) => x.name)
  const at = (name: string) => names.indexOf(name)
  const table: PivotTable = { name: 'PT1', cache: 1, location: { range: 'A3', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 }, fields: names.map(() => ({})) }
  const bases: (string | undefined)[] = []
  for (const o of c.ops) {
    if ('rows' in o) table.rows = [...(table.rows ?? []), ...o.rows.map(at)]
    else if ('columns' in o) table.columns = [...(table.columns ?? []), ...o.columns.map(at)]
    else if ('pages' in o) table.pages = [...(table.pages ?? []), ...o.pages.map((f) => ({ field: at(f) }))]
    else if ('valuesOn' in o) table.valuesOnRows = o.valuesOn === 'row'
    else if ('layout' in o) table.layout = o.layout
    else if ('props' in o && o.props.showAllItems) table.fields[at(o.field)].showEmpty = true
    else if ('value' in o) {
      const v = o.value
      const d: PivotDataField = { field: at(v.field), name: v.caption ?? wordsOf('en-US').valueName(v.fn ?? 'sum', v.field) }
      if (v.fn && v.fn !== 'sum') d.function = v.fn as PivotFunction
      if (v.show && v.show !== 'normal') d.showAs = SHOW[v.show] ?? (v.show as PivotShowAs)
      if (v.base) d.baseField = at(v.base)
      table.values = [...(table.values ?? []), d]
      bases.push(v.item)
    }
  }
  const book = sourceBook(cols, 'range')
  book.pivotCaches = [{ id: 1, source: { type: 'range', sheet: 'D', range: sourceRange(cols) }, fields: names.map((name) => ({ name })) }]
  book.sheets = book.sheets.map((s) => (s.name === 'P' ? { ...s, pivotTables: [table] } : s))
  const first = refreshed(book, true)
  const t = first.sheets.find((s) => s.name === 'P')!.pivotTables![0]
  const cache = first.pivotCaches![0]
  // an item by its label, as Excel names it, in any case
  const item = (f: number, label: string) => t.fields[f].items?.findIndex((x) => itemText(cache.fields[f].items![x.item!] ?? null).toLowerCase() === label.toLowerCase()) ?? -1
  const value = (name: string) => t.values!.findIndex((d) => d.name === name)
  t.values?.forEach((d, k) => {
    const b = bases[k]
    // a base field on no axis lists no items: its item is left out
    if (b === '(previous)' || b === '(next)') d.baseItem = b === '(previous)' ? 'previous' : 'next'
    else if (b !== undefined && d.baseField !== undefined && item(d.baseField, b) >= 0) d.baseItem = item(d.baseField, b)
  })
  for (const o of c.ops) {
    if ('hide' in o) for (const name of o.items) t.fields[at(o.hide)].items![item(at(o.hide), name)].hidden = true
    else if ('page' in o && o.items !== '(All)') {
      const f = at(o.page)
      const kept = o.items.map((name) => item(f, name))
      if (kept.length === 1) t.pages = t.pages!.map((p) => (p.field === f ? { field: f, item: kept[0] } : p))
      else {
        t.fields[f].multipleItems = true
        t.fields[f].items!.forEach((x, p) => !kept.includes(p) && (x.hidden = true))
      }
    } else if ('filter' in o) {
      const flt: PivotFilter = { field: at(o.filter), type: FILTERS[o.type], args: [o.value1 as string | number] }
      if (o.by) flt.value = value(o.by)
      t.filters = [...(t.filters ?? []), flt]
    } else if ('sort' in o && o.order !== 'manual') t.fields[at(o.sort)].sort = { order: o.order, ...(o.by ? { by: { value: value(o.by) } } : {}) }
  }
  return first
}

/** The book with its PivotTables refreshed from the source or laid out again from the cache. */
function refreshed(book: BookInput, fromSource: boolean): BookInput {
  const values = stateValues(createState(book, clockOptions({ now: '2026-10-30', seed: 0 })))
  const r = refreshPivots(book, undefined, values, values.today!(), fromSource)
  if (!r.ok) throw new Error(r.message)
  return r.book
}

const set = process.argv[2] ?? 'show'
const { buildPivotCases } = (await import(`./${set}.js`)) as PivotSet
const prediction: Record<string, unknown> = {}
for (const c of buildPivotCases()) {
  const file = new URL(`test/excel/win/pivot/${set}/${c.id}.xlsx`, ROOT)
  if (!existsSync(file)) {
    const model = modelOf(c)
    prediction[c.id] = typeof model === 'string' ? model : laidOut(model, false)
    continue
  }
  const { book } = await readXlsx(readFileSync(file))
  // the cells Excel saved on P are its answer: left out, so that nothing of them reaches the prediction
  book.sheets = book.sheets.map((s) => (s.name === 'P' ? { ...s, cells: {} } : s))
  const saved = laidOut(book, false)
  const source = laidOut(book, true)
  prediction[c.id] = JSON.stringify(saved) === JSON.stringify(source) ? saved : { saved, source }
}
writeFileSync(process.argv[3] ?? new URL(`bench/coverage/pivot/${set}.predicted.json`, ROOT), `${JSON.stringify(prediction, null, 1)}\n`)
