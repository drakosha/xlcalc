/**
 * The core's check of PivotTables, their caches, slicers and timelines (`BookInput.pivotCaches`, `SheetInput.pivotTables`
 * and the rest): what the engine reads of them (`GETPIVOTDATA`: a PivotTable's place, fields, value fields and saved
 * lines, its cache's field names and items) and what refers to what, a PivotTable to its cache by id, a field by its
 * index, a slicer to its cache and its PivotTables by name. The rest of their shape is checked by `xlcalc/pivot` when it
 * is registered (`setPivotCheck`), and by the xlsx subpaths.
 */
import type { BookInput } from '../contract.js'
import { caseKey } from '../collate.js'
import { InputError, what } from '../input-error.js'
import { dateToSerial } from '../model.js'
import { pivotCheck } from '../pivot/hook.js'
import { ERROR_TEXT } from '../syntax.js'
import { array, checkBoolean, checkNumber, checkString, isObject, object } from './basic.js'
import { checkRange } from './range.js'

/** An ISO date, as a cache stores it (`2025-01-31T00:00:00`). */
export function checkIsoDate(v: unknown, p: string) {
  checkString(v, p)
  try {
    dateToSerial(v as string)
  } catch {
    throw new InputError(p, `invalid date ${JSON.stringify(v)}: expected an ISO date such as "2025-01-31T00:00:00"`)
  }
}

/** A value of a cache: blank, number, text, boolean, `{ error }` or `{ date }`. */
export function checkPivotValue(v: unknown, p: string) {
  if (v === null || typeof v === 'string' || typeof v === 'boolean' || (typeof v === 'number' && Number.isFinite(v))) return
  if (isObject(v) && Object.keys(v).length === 1) {
    if (typeof v.date === 'string') return checkIsoDate(v.date, `${p}.date`)
    if (Object.hasOwn(ERROR_TEXT, v.error as string)) return
  }
  throw new InputError(p, `expected a number, text, a boolean, null, { error: code } or { date: "2025-01-31T00:00:00" }, got ${what(v)}`)
}

const MAX = 2 ** 31

/** `o[k]` when given, checked by `check`. */
function given(o: Record<string, unknown>, k: string, p: string, check: (v: unknown, p: string) => void) {
  if (o[k] !== undefined) check(o[k], `${p}.${k}`)
}

/** The objects of an array, each given to `each` with its path. */
function objects(v: unknown, p: string, each: (o: Record<string, unknown>, p: string, i: number) => void) {
  array(v, p, 'an array').forEach((x, i) => each(object(x, `${p}[${i}]`, 'an object'), `${p}[${i}]`, i))
}

const index = (v: unknown, p: string, n = MAX) => checkNumber(v, p, 0, n - 1, true)
const text = (v: unknown, p: string) => void checkString(v, p)

/** A field of a PivotTable whose cache has `n`: its index, or `values` where `values` may be. */
function field(v: unknown, p: string, n: number, values = false) {
  if (values && v === 'values') return
  if (typeof v === 'number' && Number.isInteger(v) && v >= n) throw new InputError(p, `expected the index of a field of the PivotTable's cache, 0 to ${n - 1}, got ${v}`)
  index(v, p)
}

/** Names unique within `items`, compared as Excel compares them. */
function unique(items: unknown, path: string, kind: string) {
  const seen = new Set<string>()
  ;(items as { name: string }[] | undefined)?.forEach((x, i) => {
    if (seen.has(caseKey(x.name))) throw new InputError(`${path}[${i}].name`, `duplicate ${kind} name ${JSON.stringify(x.name)}`)
    seen.add(caseKey(x.name))
  })
}

const line = (o: Record<string, unknown>, p: string) => {
  array(o.items, `${p}.items`, 'an array of indexes').forEach((x, i) => index(x, `${p}.items[${i}]`))
  given(o, 'repeat', p, index)
  given(o, 'value', p, index)
  given(o, 'type', p, text)
}

/** A PivotTable over a cache of `n` fields: what the core reads of it. */
function pivotTable(t: Record<string, unknown>, p: string, n: number) {
  text(t.name, `${p}.name`)
  const at = object(t.location, `${p}.location`, 'a place such as { range: "A3:B8", firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 }')
  checkRange(at.range, `${p}.location.range`)
  for (const k of ['firstHeaderRow', 'firstDataRow', 'firstDataColumn']) index(at[k], `${p}.location.${k}`)
  for (const k of ['pageRows', 'pageColumns']) given(at, k, `${p}.location`, index)
  const fields = array(t.fields, `${p}.fields`, 'an array of fields')
  if (fields.length !== n) throw new InputError(`${p}.fields`, `expected a field for each of the cache's ${n} fields, got ${fields.length}`)
  objects(fields, `${p}.fields`, (f, fp) => {
    given(f, 'name', fp, text)
    if (f.items !== undefined) objects(f.items, `${fp}.items`, (x, xp) => (given(x, 'item', xp, index), given(x, 'caption', xp, text)))
  })
  for (const k of ['rows', 'columns']) if (t[k] !== undefined) array(t[k], `${p}.${k}`, 'an array of fields').forEach((x, i) => field(x, `${p}.${k}[${i}]`, n, true))
  if (t.pages !== undefined) objects(t.pages, `${p}.pages`, (x, xp) => (field(x.field, `${xp}.field`, n), given(x, 'item', xp, index)))
  if (t.values !== undefined) objects(t.values, `${p}.values`, (x, xp) => (field(x.field, `${xp}.field`, n), given(x, 'name', xp, text), given(x, 'function', xp, text)))
  if (t.filters !== undefined) objects(t.filters, `${p}.filters`, (x, xp) => field(x.field, `${xp}.field`, n))
  given(t, 'valuesOnRows', p, checkBoolean)
  given(t, 'valuesPosition', p, index)
  for (const k of ['rowLines', 'columnLines']) if (t[k] !== undefined) objects(t[k], `${p}.${k}`, line)
}

/**
 * The book's and its sheets' PivotTables, caches, slicers and timelines: what the core reads of them, a PivotTable's cache
 * by id with a field for each of its fields, a slicer's and a timeline's cache by name, a cache's PivotTables by sheet and
 * name; then the full check `xlcalc/pivot` registered.
 */
export function checkPivots(book: Record<string, unknown>, path: string) {
  const fields = new Map<number, number>()
  if (book.pivotCaches !== undefined)
    objects(book.pivotCaches, `${path}.pivotCaches`, (c, p) => {
      index(c.id, `${p}.id`)
      if (fields.has(c.id as number)) throw new InputError(`${p}.id`, `duplicate pivot cache id ${c.id}`)
      const list = array(c.fields, `${p}.fields`, 'an array of fields')
      fields.set(c.id as number, list.length)
      objects(list, `${p}.fields`, (f, fp) => {
        text(f.name, `${fp}.name`)
        const values = (v: unknown, vp: string) => array(v, vp, 'an array of values').forEach((x, i) => checkPivotValue(x, `${vp}[${i}]`))
        given(f, 'items', fp, values)
        if (f.group !== undefined) given(object(f.group, `${fp}.group`, 'a grouping'), 'items', `${fp}.group`, values)
      })
    })
  const pivots = new Set<string>()
  const b = book as unknown as BookInput
  b.sheets.forEach((s, i) => {
    const p = `${path}.sheets[${i}]`
    if (s.pivotTables === undefined) return
    objects(s.pivotTables, `${p}.pivotTables`, (t, tp) => {
      const n = fields.get(t.cache as number)
      if (n === undefined) throw new InputError(`${tp}.cache`, `expected the id of a pivot cache of the workbook, got ${what(t.cache)}`)
      pivotTable(t, tp, n)
      pivots.add(caseKey(s.name) + '\0' + caseKey(t.name as string))
    })
    unique(s.pivotTables, `${p}.pivotTables`, 'PivotTable')
  })
  for (const k of ['slicerCaches', 'timelineCaches'] as const) {
    if (book[k] === undefined) continue
    objects(book[k], `${path}.${k}`, (c, p) => {
      text(c.name, `${p}.name`)
      given(c, 'cache', p, (v, vp) => {
        if (!fields.has(v as number)) throw new InputError(vp, `expected the id of a pivot cache of the workbook, got ${what(v)}`)
      })
      objects(c.pivots, `${p}.pivots`, (r, rp) => {
        if (!pivots.has(caseKey(String(r.sheet)) + '\0' + caseKey(String(r.name)))) throw new InputError(rp, `expected a PivotTable of the workbook, got ${JSON.stringify(r)}`)
      })
    })
    unique(book[k], `${path}.${k}`, 'cache')
  }
  b.sheets.forEach((s, i) => {
    for (const [k, of] of [['slicers', b.slicerCaches], ['timelines', b.timelineCaches]] as const)
      if (s[k] !== undefined)
        objects(s[k], `${path}.sheets[${i}].${k}`, (x, p) => {
          text(x.name, `${p}.name`)
          if (!of?.some((c) => caseKey(c.name) === caseKey(String(x.cache)))) throw new InputError(`${p}.cache`, `expected the name of a cache of the workbook, got ${JSON.stringify(x.cache)}`)
        })
  })
  pivotCheck()?.(b, path)
}
