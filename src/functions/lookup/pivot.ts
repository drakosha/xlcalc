/**
 * GETPIVOTDATA by a PivotTable's saved layout: the PivotTable whose cells hold the reference, its value field by name
 * (its caption or its source field's, in any case), each pair of a field and an item matched as Excel matches them,
 * then the row and the column of its saved lines (`rowLines`, `columnLines`) whose fields and items are exactly those
 * the pairs name, a subtotal or a grand total where the pairs name fewer; the answer is the value of that cell. What
 * the layout does not show is #REF!, as are a PivotTable laid out before its last change (`PivotCache.stale`) and one in
 * another workbook, which the formula's file keeps its value of (`held.ts`). The PivotTable's cells are learned as the
 * formula's references, so an edit of them computes it again; a change of the PivotTable builds the book again.
 */
import type { Arg, Cx, FnDef } from '../../args.js'
import { caseKey } from '../../collate.js'
import type { NumberFormat, PivotAxisField, PivotCache, PivotLine, PivotTable, PivotValue } from '../../contract.js'
import { dateToSerial } from '../../model.js'
import { itemFormat } from '../../pivot-format.js'
import { groupedItems } from '../../pivot/items.js'
import { parseRange } from '../../structure/area.js'
import { ERROR_TEXT } from '../../syntax.js'
import { E, isErr, Rng, type V } from '../../values.js'

/** A line of an axis as a cell's place: the fields it names with their items, the value field, and whether it holds values. */
interface Slot {
  items: Map<number, number>
  value: number
  /** A line of an outer field's item, which holds its subtotal where no line of the subtotal's own does (several functions, `subtotals-three-compact`). */
  header?: boolean
}

/** The PivotTable's form of a field: its own, else the PivotTable's. */
const formOf = (t: PivotTable, f: number) => t.fields[f]?.layout ?? t.layout ?? 'compact'

/**
 * An axis's lines as places, null for a line that holds no values of its own: a blank line, a line of a field's item
 * whose subtotal is not shown on it (subtotals off, at the bottom, or the tabular form) and which is not collapsed, a
 * line above the value fields where they lie across the axis.
 */
function slots(t: PivotTable, axis: PivotAxisField[], lines: PivotLine[] | undefined): (Slot | null)[] {
  const at = axis.indexOf('values')
  let prev: number[] = []
  return (lines ?? [{ items: [] }]).map((line) => {
    const items = [...prev.slice(0, line.repeat ?? 0), ...line.items]
    prev = items
    if (line.type === 'blank') return null
    const depth = line.type === 'grand' ? 0 : items.length
    const out: Slot = { items: new Map(), value: at >= 0 && at < depth ? items[at] : (line.value ?? 0) }
    for (let k = 0; k < depth; k++) if (axis[k] !== 'values') out.items.set(axis[k] as number, items[k])
    if (line.type !== undefined || depth >= axis.length) return out
    if (at >= depth) return null
    // a line of an outer field's item: its total where the item is collapsed (in any form), its subtotal where it is
    // shown at the top; a value field's above the items, none
    const f = axis[depth - 1]
    if (f === 'values') return null
    const field = t.fields[f]
    if (field?.items?.[items[depth - 1]]?.collapsed) return out
    return field?.subtotals?.length !== 0 && field?.subtotalTop !== false && formOf(t, f) !== 'tabular' ? { ...out, header: true } : null
  })
}

/** The value of a cache's item as the cell holds it: a date as the book's serial number. */
function cellOf(v: PivotValue, epoch: number): V {
  if (v === null || typeof v !== 'object') return v
  if ('date' in v) return dateToSerial(v.date) - epoch
  return E[v.error]
}

/** The values of a cache field's items, or of its groups' where it is grouped. */
const itemsOf = (c: PivotCache, f: number): PivotValue[] => groupedItems(c.fields[f]) ?? []

/**
 * Whether an argument names an item: text by the item's label alone (in any case; its caption or its value where it is
 * renamed; `(blank)` a blank item, a number or a date by its text in the field's format: `"1/1/2024"`, not `"2024-01-01"`, `"10"`, not `"10.0"`); a number by an equal
 * number or date; a boolean by an equal boolean.
 */
function names(cx: Cx, arg: V, item: V, label: string | undefined, format: NumberFormat | undefined): boolean {
  // a number by the comparison of `=`, to 15 significant digits: a date and time typed or calculated equals the item's
  if (typeof arg === 'number') return typeof item === 'number' && cx.rules.compareOp(arg, item) === 0
  if (typeof arg !== 'string') return typeof arg === typeof item && arg === item
  const key = caseKey(arg)
  // a renamed item by its caption or by its value (the corpus: a pair of an item's value, its caption another)
  if (label !== undefined && caseKey(label) === key) return true
  if (item === null) return key === caseKey('(blank)')
  if (typeof item === 'number') return caseKey(cx.rules.display(item, format, Infinity).text) === key
  if (typeof item === 'boolean') return caseKey(cx.rules.toText(item) as string) === key
  if (isErr(item)) return caseKey(ERROR_TEXT[item.code]) === key
  return caseKey(item) === key
}

/** The indexes of a field's items (its place in the PivotTable's field) an argument names: text its text items before others of that label. */
function itemsNamed(cx: Cx, t: PivotTable, c: PivotCache, f: number, arg: V): Set<number> {
  const values = itemsOf(c, f)
  const found: [k: number, text: boolean][] = []
  t.fields[f]?.items?.forEach((it, k) => {
    const v = it.item === undefined ? undefined : values[it.item]
    if (v !== undefined && names(cx, arg, cellOf(v, cx.rules.epoch), it.caption, itemFormat(c, f, v))) found.push([k, typeof v === 'string'])
  })
  const text = typeof arg === 'string' && found.some(([, x]) => x)
  const kept = found.filter(([, x]) => !text || x)
  // items whose labels differ in case alone (a cache that keeps both): the one spelled as the argument where there is one
  const spelled = kept.filter(([k]) => typeof arg === 'string' && values[t.fields[f].items![k].item!] === arg)
  return new Set((spelled.length ? spelled : kept).map(([k]) => k))
}

/** A field by its name in the PivotTable, the source's where it has none of its own, in any case (the set gpd-held). */
function fieldNamed(t: PivotTable, c: PivotCache, name: string): number {
  const key = caseKey(name)
  return c.fields.findIndex((f, k) => caseKey(t.fields[k]?.name ?? f.name) === key)
}

/** A value field by its caption, else by its source field's name, in any case. */
function valueNamed(t: PivotTable, c: PivotCache, name: string): number {
  const key = caseKey(name)
  const values = t.values ?? []
  const k = values.findIndex((d) => d.name !== undefined && caseKey(d.name) === key)
  return k >= 0 ? k : values.findIndex((d) => c.fields[d.field] !== undefined && caseKey(c.fields[d.field].name) === key)
}

/**
 * The slot of the pairs on an axis, of the value field `value` where the axis tells one: the line whose fields are
 * exactly those of `wanted` on the axis, each with one of its items, a line of an item's own subtotal before the item's
 * line above it; several such lines (subtotals of several functions) name none. Without one, a line that names more
 * fields, the fewest, where only one does: an inner field named without the one between (the set gpd-held).
 */
function slotOf(list: (Slot | null)[], wanted: Map<number, Set<number>>, own: (f: number) => boolean, value: number, tells: boolean): number {
  const fits = (x: Slot | null, exact: boolean): x is Slot => {
    if (!x || (tells && x.value !== value)) return false
    for (const [f, items] of wanted) if (own(f) && !(x.items.has(f) && items.has(x.items.get(f)!))) return false
    if (exact) for (const f of x.items.keys()) if (!wanted.has(f)) return false
    return true
  }
  const at = (keep: (x: Slot) => boolean) => list.flatMap((x, k) => (x && keep(x) ? [k] : []))
  const exact = at((x) => fits(x, true))
  const lines = exact.filter((k) => !list[k]!.header)
  // lines of the same items are a field's subtotals of several functions; of other items, the first is taken
  const same = (a: Slot, b: Slot) => a.items.size === b.items.size && [...a.items].every(([f, i]) => b.items.get(f) === i)
  if (exact.length) return lines.length > 1 && same(list[lines[0]]!, list[lines[1]]!) ? -1 : (lines[0] ?? exact[0])
  const more = at((x) => fits(x, false))
  const fewest = Math.min(...more.map((k) => list[k]!.items.size))
  const least = more.filter((k) => list[k]!.items.size === fewest)
  return least.length === 1 ? least[0] : -1
}

function getPivotData(cx: Cx, data: Arg, ref: Arg, pairs: V[]): Arg {
  if (isErr(data)) return data
  // the value field by a reference to a cell holding its name: #REF! (the set gpd-1)
  if (!(ref instanceof Rng) || data instanceof Rng) return isErr(ref) ? ref : E.REF
  if (ref.s >= cx.book.sheets) return E.REF
  const at = cx.book.pivotAt(ref.s, ref.r1, ref.c1)
  if (!at?.cache || at.cache.stale) return E.REF
  const { table: t, cache: c, area } = at
  cx.found(new Rng(ref.s, area.r1, area.c1, area.r2, area.c2))
  const name = cx.rules.toText(cx.scalar(data))
  if (isErr(name)) return name
  const value = valueNamed(t, c, name)
  if (value < 0 || pairs.length % 2) return E.REF
  const rows = t.rows ?? []
  const columns = t.columns ?? []
  const wanted = new Map<number, Set<number>>()
  for (let k = 0; k < pairs.length; k += 2) {
    const f = fieldNamed(t, c, cx.rules.toText(pairs[k]) as string)
    const items = f < 0 ? new Set<number>() : itemsNamed(cx, t, c, f, pairs[k + 1])
    if (!items.size) return E.REF
    const page = t.pages?.find((p) => p.field === f)
    if (page) {
      // a field of the Filters area: the item it shows
      if (page.item === undefined || !items.has(page.item)) return E.REF
      continue
    }
    if (!rows.includes(f) && !columns.includes(f)) return E.REF
    wanted.set(f, items)
  }
  const onRows = rows.includes('values') || (t.valuesOnRows === true && !columns.includes('values'))
  const r = slotOf(slots(t, rows, t.rowLines), wanted, (f) => rows.includes(f), value, onRows && (t.values?.length ?? 0) > 1)
  const col = slotOf(slots(t, columns, t.columnLines), wanted, (f) => columns.includes(f), value, !onRows && (t.values?.length ?? 0) > 1)
  if (r < 0 || col < 0) return E.REF
  const x = parseRange(t.location.range)!
  const row = x.r1 + t.location.firstDataRow + r
  const column = x.c1 + t.location.firstDataColumn + col
  return cx.scalar(new Rng(ref.s, row, column, row, column)) ?? 0
}

export const functions: Record<string, FnDef> = {
  GETPIVOTDATA: { args: 'R R v?', rep: 1, max: 254, cx: true, dynamic: true, fn: getPivotData },
}
