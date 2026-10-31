/**
 * Calculated fields and items in a PivotTable's calculation. A calculated field is its formula over the sums of the
 * fields it names at each cell of the cube, totals too, not over each record (the set calc: Amount/Qty over (100, 1) and
 * (100, 9) is 20). A calculated item is its formula over the values of the field's items it names at the same place of
 * the cube, by the value field's function. Formulas are calculated by the engine in batches (`formulas.ts`).
 */
import { caseKey } from '../collate.js'
import type { CellValue, PivotCache, PivotTable } from '../contract.js'
import { calculate, template, type Template } from './formulas.js'
import { groupedItems } from './group.js'
import type { Branch } from './tree.js'
import { itemText } from './values.js'

/** Thrown where a formula's value waits for the next batch. */
export const PENDING: unique symbol = Symbol('pending')

/** Formulas over their arguments' values, each calculated once: a value asked for first waits for `run`. */
export class Batch {
  private done = new Map<string, CellValue>()
  private asked = new Map<string, { text: string; values: CellValue[] }>()

  /** The value of `text` over `values`, or `PENDING` thrown until `run` calculates it. */
  get(text: string, values: CellValue[]): CellValue {
    const k = `${text}\u0000${JSON.stringify(values)}`
    const v = this.done.get(k)
    if (v !== undefined) return v
    this.asked.set(k, { text, values })
    throw PENDING
  }

  /** Calculates the formulas asked for; false when there were none. */
  run(): boolean {
    if (!this.asked.size) return false
    const asked = [...this.asked]
    this.asked = new Map()
    const got = calculate(asked.map(([, x]) => x))
    asked.forEach(([k], i) => this.done.set(k, got[i]))
    return true
  }
}

/** Calls `f` until it gives a value, calculating the batch's formulas it waits for. */
export function settled<T>(batch: Batch, f: () => T): T {
  for (;;) {
    try {
      return f()
    } catch (e) {
      if (e !== PENDING || !batch.run()) throw e
    }
  }
}

/** The calculated fields of a cache by index, each its formula read over the fields it names (by index). */
export function fieldFormulas(cache: PivotCache): Map<number, Template<number>> {
  const byName = new Map(cache.fields.map((f, i) => [caseKey(f.name), i]))
  const out = new Map<number, Template<number>>()
  cache.fields.forEach((f, i) => {
    if (f.formula !== undefined) out.set(i, template(f.formula, (field, name) => (field === null ? byName.get(caseKey(name)) : undefined)))
  })
  return out
}

/** The fields not calculated whose sums the calculated fields `fields` take, through calculated fields they name too. */
export function sumsOf(fields: Iterable<number>, formulas: ReadonlyMap<number, Template<number>>): number[] {
  const out = new Set<number>()
  const seen = new Set<number>()
  const visit = (f: number) => {
    if (seen.has(f)) return
    seen.add(f)
    const t = formulas.get(f)
    if (!t) out.add(f)
    else t.args.forEach(visit)
  }
  for (const f of fields) visit(f)
  return [...out]
}

/**
 * The calculated items of a PivotTable's fields, by field and the item's place on the PivotTable, each its formula read
 * over the places of the items it names: by the item's value or caption, in any case.
 */
export function itemFormulas(table: PivotTable, cache: PivotCache, epoch: number): Map<number, Map<number, Template<number>>> {
  const out = new Map<number, Map<number, Template<number>>>()
  for (const ci of cache.calculatedItems ?? []) {
    const pf = table.fields[ci.field]
    const values = groupedItems(cache.fields[ci.field]) ?? []
    const place = pf?.items?.findIndex((x) => x.item === ci.item) ?? -1
    if (!pf?.items || place < 0) continue
    const named = new Map<string, number>()
    pf.items.forEach((x, p) => {
      if (x.calculated || x.item === undefined) return
      for (const label of [itemText(values[x.item] ?? null, epoch), x.caption]) if (label !== undefined && !named.has(caseKey(label))) named.set(caseKey(label), p)
    })
    const own = caseKey(cache.fields[ci.field].name)
    const t = template(ci.formula, (field, name) => (field === null || caseKey(field) === own ? named.get(caseKey(name)) : undefined))
    ;(out.get(ci.field) ?? out.set(ci.field, new Map()).get(ci.field)!).set(place, t)
  }
  return out
}

/** A copy of branches merged into one: the items of each, under each the merge of theirs. */
function union(bs: readonly Branch[]): Branch {
  const next = new Map<number, Branch[]>()
  for (const b of bs) for (const [p, n] of b.next) (next.get(p) ?? next.set(p, []).get(p)!).push(n)
  const out: Branch = { data: true, next: new Map(), order: [] }
  for (const [p, ns] of [...next].sort(([x], [y]) => x - y)) {
    out.next.set(p, union(ns))
    out.order.push(p)
  }
  return out
}

/**
 * The calculated items of the axis's fields `real` put under every item of the field before, shown ones only: each
 * last among them, with the items under the ones it names below it.
 */
export function withCalculatedItems(root: Branch, real: readonly number[], items: ReadonlyMap<number, Map<number, Template<number>>>, shown: readonly (boolean[] | undefined)[]): void {
  const walk = (b: Branch, level: number) => {
    if (level === real.length) return
    const f = real[level]
    const own = items.get(f)
    for (const n of b.next.values()) walk(n, level + 1)
    if (!own) return
    for (const [p, t] of own) {
      if (!(shown[f]?.[p] ?? true) || b.next.has(p)) continue
      b.next.set(p, union(t.args.flatMap((q) => (b.next.has(q) ? [b.next.get(q)!] : []))))
      b.order.push(p)
    }
  }
  walk(root, 0)
}

/** Every path of an axis's tree, the empty one (its total) first. */
function paths(root: Branch): number[][] {
  const out: number[][] = [[]]
  const walk = (b: Branch, path: number[]) => {
    for (const p of b.order) {
      out.push([...path, p])
      walk(b.next.get(p)!, [...path, p])
    }
  }
  walk(root, [])
  return out
}

/**
 * The formulas of every cell of the trees calculated ahead, in as many batches as they wait on each other, so that the
 * layout reads them done: `value` of each value field at each pair of paths.
 */
export function prefetch(batch: Batch, rows: Branch, columns: Branch, values: number, value: (rows: number[], columns: number[], v: number) => CellValue): void {
  const [rs, cs] = [paths(rows), paths(columns)]
  do {
    for (const r of rs)
      for (const c of cs)
        for (let v = 0; v < values; v++) {
          try {
            value(r, c, v)
          } catch (e) {
            if (e !== PENDING) throw e
          }
        }
  } while (batch.run())
}
