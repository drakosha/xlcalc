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
import { epochOf, type DateSystem } from '../dates.js'
import { beforeDayZero } from '../date-serial.js'
import { modelDate } from '../pivot-time.js'
import { pivotCheck } from '../pivot/hook.js'
import { array, checkBoolean, checkCellValue, checkNumber, checkString, given, isObject, object, type Check } from './basic.js'
import { checkRange } from './range.js'

/**
 * An ISO date, as a cache stores it (`2025-01-31T00:00:00`), of the workbook's date system: from its day 0 on, as the
 * layout reads it, a time of day on Excel's 1899-12-30 at day 0 (`pivot-time.ts`).
 */
export function checkIsoDate(v: unknown, p: string, dateSystem?: DateSystem) {
  checkString(v, p)
  let serial: number
  try {
    serial = dateToSerial(modelDate(v as string))
  } catch {
    throw new InputError(p, `invalid date ${JSON.stringify(v)}: expected an ISO date such as "2025-01-31T00:00:00"`)
  }
  // the day 0 the layout reads its dates from (`date-serial.ts`)
  const e = epochOf(dateSystem)
  if (serial < e) throw new InputError(p, beforeDayZero(v as string, e))
}

/** A value of a cache: a cell's (blank, number, text, boolean, `{ error }`) or `{ date }` of the workbook's date system. */
export function checkPivotValue(v: unknown, p: string, dateSystem?: DateSystem) {
  if (isObject(v) && Object.keys(v).length === 1 && typeof v.date === 'string') return checkIsoDate(v.date, `${p}.date`, dateSystem)
  try {
    checkCellValue(v, p)
  } catch {
    throw new InputError(p, `expected a number, text, a boolean, null, { error: code } or { date: "2025-01-31T00:00:00" }, got ${what(v)}`)
  }
}

type Fields = Record<string, unknown>

const MAX = 2 ** 31
const index: Check = (v, p) => checkNumber(v, p, 0, MAX - 1, true)
const text: Check = checkString

/** An array, each of its elements checked by `each`. */
const list = (each: Check, expected: string): Check => (v, p) => array(v, p, expected).forEach((x, i) => each(x, `${p}[${i}]`))

/** An array of objects, each checked by `each`. */
const objects = (each: (o: Fields, p: string) => void): Check => list((x, p) => each(object(x, p, 'an object'), p), 'an array')

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

const line = objects((o, p) => {
  list(index, 'an array of indexes')(o.items, `${p}.items`)
  given(o, p, { repeat: index, value: index, type: text })
})

/** A PivotTable over a cache of `n` fields: what the core reads of it. */
function pivotTable(t: Fields, p: string, n: number) {
  text(t.name, `${p}.name`)
  const at = object(t.location, `${p}.location`, 'a place such as { range: "A3:B8", firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 }')
  checkRange(at.range, `${p}.location.range`)
  for (const k of ['firstHeaderRow', 'firstDataRow', 'firstDataColumn']) index(at[k], `${p}.location.${k}`)
  given(at, `${p}.location`, { pageRows: index, pageColumns: index })
  const fields = array(t.fields, `${p}.fields`, 'an array of fields')
  if (fields.length !== n) throw new InputError(`${p}.fields`, `expected a field for each of the cache's ${n} fields, got ${fields.length}`)
  objects((f, fp) => given(f, fp, { name: text, items: objects((x, xp) => given(x, xp, { item: index, caption: text })) }))(fields, `${p}.fields`)
  const axis = list((v, q) => field(v, q, n, true), 'an array of fields')
  // an object of the PivotTable's areas: its field's index, then `more`
  const of = (more: Record<string, Check>) => objects((x, xp) => (field(x.field, `${xp}.field`, n), given(x, xp, more)))
  given(t, p, { rows: axis, columns: axis, pages: of({ item: index }), values: of({ name: text, function: text }), filters: of({}), valuesOnRows: checkBoolean, valuesPosition: index, rowLines: line, columnLines: line })
}

/**
 * The book's and its sheets' PivotTables, caches, slicers and timelines: what the core reads of them, a PivotTable's cache
 * by id with a field for each of its fields, a slicer's and a timeline's cache by name, a cache's PivotTables by sheet and
 * name; then the full check `xlcalc/pivot` registered.
 */
export function checkPivots(book: Fields, path: string) {
  const fields = new Map<number, number>()
  const cache: Check = (v, p) => {
    if (!fields.has(v as number)) throw new InputError(p, `expected the id of a pivot cache of the workbook, got ${what(v)}`)
  }
  const dateSystem = book.dateSystem as DateSystem | undefined
  const values = list((v, p) => checkPivotValue(v, p, dateSystem), 'an array of values')
  given(book, path, {
    pivotCaches: objects((c, p) => {
      index(c.id, `${p}.id`)
      if (fields.has(c.id as number)) throw new InputError(`${p}.id`, `duplicate pivot cache id ${c.id}`)
      const each = array(c.fields, `${p}.fields`, 'an array of fields')
      fields.set(c.id as number, each.length)
      objects((f, fp) => {
        text(f.name, `${fp}.name`)
        given(f, fp, { items: values })
        if (f.group !== undefined) given(object(f.group, `${fp}.group`, 'a grouping'), `${fp}.group`, { items: values })
      })(each, `${p}.fields`)
    }),
  })
  const pivots = new Set<string>()
  const b = book as unknown as BookInput
  b.sheets.forEach((s, i) => {
    const p = `${path}.sheets[${i}]`
    given(s as unknown as Fields, p, {
      pivotTables: objects((t, tp) => {
        cache(t.cache, `${tp}.cache`)
        pivotTable(t, tp, fields.get(t.cache as number)!)
        pivots.add(caseKey(s.name) + '\0' + caseKey(t.name as string))
      }),
    })
    unique(s.pivotTables, `${p}.pivotTables`, 'PivotTable')
  })
  for (const k of ['slicerCaches', 'timelineCaches'] as const) {
    given(book, path, {
      [k]: objects((c, p) => {
        text(c.name, `${p}.name`)
        given(c, p, { cache })
        objects((r, rp) => {
          if (!pivots.has(caseKey(String(r.sheet)) + '\0' + caseKey(String(r.name)))) throw new InputError(rp, `expected a PivotTable of the workbook, got ${JSON.stringify(r)}`)
        })(c.pivots, `${p}.pivots`)
      }),
    })
    unique(book[k], `${path}.${k}`, 'cache')
  }
  b.sheets.forEach((s, i) => {
    for (const [k, of] of [['slicers', b.slicerCaches], ['timelines', b.timelineCaches]] as const)
      given(s as unknown as Fields, `${path}.sheets[${i}]`, {
        [k]: objects((x, p) => {
          text(x.name, `${p}.name`)
          if (!of?.some((c) => caseKey(c.name) === caseKey(String(x.cache)))) throw new InputError(`${p}.cache`, `expected the name of a cache of the workbook, got ${JSON.stringify(x.cache)}`)
        }),
      })
  })
  pivotCheck()?.(b, path)
}
