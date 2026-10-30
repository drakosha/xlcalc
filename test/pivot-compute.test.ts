import { describe, expect, it } from 'vitest'
import type { CellValue, PivotCache, PivotTable } from '../src/contract.js'
import { computePivot } from '../src/pivot/compute.js'
import { refreshCache } from '../src/pivot/refresh.js'
import type { PivotNode, PivotResult } from '../src/pivot/result.js'

const SOURCE: CellValue[][] = [
  ['Region', 'Product', 'Amount', 'Qty'],
  ['North', 'Apple', 10, 1],
  ['South', 'Banana', 20, 2],
  ['north', 'Cherry', 'n/a', 3],
  ['East', 'Apple', null, 4],
  ['South', 'Apple', 5, 5],
  ['East', 'Banana', { error: 'NA' }, 6],
  [null, 'Cherry', 7, 7],
]

const EMPTY: PivotCache = { id: 1, source: { type: 'range', sheet: 'D', range: 'A1:D8' }, fields: [] }

/** A cache refreshed from `rows` and a PivotTable over it built by `make`. */
function pivot(make: (t: PivotTable) => PivotTable, rows = SOURCE) {
  const first = refreshCache(EMPTY, [], { rows })
  if (typeof first === 'string') throw new Error(first)
  const table = make({ name: 'PT1', cache: 1, location: { range: 'A3:B9', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 }, fields: first.cache.fields.map(() => ({})) })
  return computePivot(table, first.cache)
}

/** An axis's items as their labels, nested. */
function labels(r: PivotResult, nodes: PivotNode[]): unknown[] {
  return nodes.map((n) => {
    const label = n.field === 'values' ? `Σ${n.item}` : r.cache.fields[n.field].items![r.table.fields[n.field].items![n.item].item!]
    return n.children.length ? [label, labels(r, n.children)] : label
  })
}

/** The item index of a label on a field of the PivotTable. */
const at = (r: PivotResult, field: number, label: unknown) => r.table.fields[field].items!.findIndex((x) => r.cache.fields[field].items![x.item!] === label)

describe('pivot cache refresh', () => {
  it('lists text in any case as one item, its first spelling kept, and keeps numbers without items', () => {
    const r = refreshCache(EMPTY, [], { rows: SOURCE })
    if (typeof r === 'string') throw new Error(r)
    expect(r.cache.fields.map((f) => f.name)).toEqual(['Region', 'Product', 'Amount', 'Qty'])
    expect(r.cache.fields[0].items).toEqual(['North', 'South', 'East', null])
    expect(r.cache.fields[2].items).toEqual([10, 20, 'n/a', null, 5, { error: 'NA' }, 7])
    expect(r.cache.fields[3].items).toBeUndefined()
    expect(r.cache.fields[3].facts).toEqual({ types: ['number'], integer: true, min: 1, max: 7 })
    expect(r.cache.records![2]).toEqual([0, 2, 2, 3])
  })

  it('numbers repeated headers and refuses an empty one', () => {
    const r = refreshCache(EMPTY, [], { rows: [['A', 'a', 'A'], [1, 2, 3]] })
    expect(typeof r !== 'string' && r.cache.fields.map((f) => f.name)).toEqual(['A', 'a2', 'A3'])
    expect(refreshCache(EMPTY, [], { rows: [['A', null], [1, 2]] })).toBe('field_name')
  })
})

describe('pivot calculation', () => {
  it('sums by region, the blank region last, the grand total over every record', () => {
    const r = pivot((t) => ({ ...t, rows: [0], values: [{ field: 3 }] }))
    expect(labels(r, r.rows.items)).toEqual(['East', 'North', 'South', null])
    expect(r.value([at(r, 0, 'North')], [], 0)).toBe(4)
    expect(r.value([], [], 0)).toBe(28)
  })

  it('gives an error among the values as the value, text as no number', () => {
    const r = pivot((t) => ({ ...t, rows: [0], values: [{ field: 2 }, { field: 2, function: 'count' }, { field: 2, function: 'average' }] }))
    const [north, east] = [at(r, 0, 'North'), at(r, 0, 'East')]
    expect(r.value([north], [], 0)).toBe(10)
    expect(r.value([north], [], 1)).toBe(2)
    expect(r.value([east], [], 0)).toEqual({ error: 'NA' })
    expect(r.value([east], [], 1)).toEqual({ error: 'NA' })
    expect(r.value([], [], 2)).toEqual({ error: 'NA' })
    expect(labels(r, r.columns.items)).toEqual(['Σ0', 'Σ1', 'Σ2'])
  })

  it('leaves a cell of blank values empty and makes a sum of text 0', () => {
    const rows: CellValue[][] = [['R', 'V'], ['a', null], ['b', 'x'], ['c', 2]]
    const r = pivot((t) => ({ ...t, rows: [0], values: [{ field: 1 }, { field: 1, function: 'average' }] }), rows)
    expect(r.value([at(r, 0, 'a')], [], 0)).toBeNull()
    expect(r.value([at(r, 0, 'a')], [], 1)).toBeNull()
    expect(r.value([at(r, 0, 'b')], [], 0)).toBe(0)
    expect(r.value([at(r, 0, 'b')], [], 1)).toEqual({ error: 'DIV0' })
    expect(r.hasData([at(r, 0, 'a')], [])).toBe(true)
  })

  it('crosses rows and columns, a missing pair empty', () => {
    const r = pivot((t) => ({ ...t, rows: [0], columns: [1], values: [{ field: 3 }] }))
    const [south, banana, cherry] = [at(r, 0, 'South'), at(r, 1, 'Banana'), at(r, 1, 'Cherry')]
    expect(r.value([south], [banana], 0)).toBe(2)
    expect(r.value([south], [cherry], 0)).toBeNull()
    expect(r.hasData([south], [cherry])).toBe(false)
    expect(r.value([], [cherry], 0)).toBe(10)
  })

  it('hides items and filters by a page item, the totals over what is left', () => {
    const r = pivot((t) => {
      const fields = t.fields.slice()
      fields[0] = { items: [{ item: 0 }, { item: 1, hidden: true }, { item: 2 }, { item: 3 }] }
      return { ...t, fields, rows: [0], pages: [{ field: 1, item: 0 }], values: [{ field: 3 }] }
    })
    expect(labels(r, r.rows.items)).toEqual(['North', 'East'])
    expect(r.value([], [], 0)).toBe(5)
    expect(r.pages).toEqual([{ field: 1, items: [0], all: false }])
  })

  it('sorts by label descending and by value', () => {
    const byLabel = pivot((t) => ({ ...t, rows: [1], values: [{ field: 3 }], fields: t.fields.map((f, i) => (i === 1 ? { sort: { order: 'descending' } } : f)) }))
    expect(labels(byLabel, byLabel.rows.items)).toEqual(['Cherry', 'Banana', 'Apple'])
    const byValue = pivot((t) => ({ ...t, rows: [0], values: [{ field: 3 }], fields: t.fields.map((f, i) => (i === 0 ? { sort: { order: 'descending', by: { value: 0 } } } : f)) }))
    expect(labels(byValue, byValue.rows.items)).toEqual(['East', 'South', null, 'North'])
  })

  it('keeps the top items by value within each item before, ties kept', () => {
    const r = pivot((t) => ({ ...t, rows: [0], values: [{ field: 3 }], filters: [{ field: 0, type: 'count', args: [2], value: 0 }] }))
    expect(labels(r, r.rows.items)).toEqual(['East', 'South', null])
    expect(r.value([], [], 0)).toBe(24)
  })

  it('filters labels with wildcards in any case', () => {
    const r = pivot((t) => ({ ...t, rows: [1], values: [{ field: 3 }], filters: [{ field: 1, type: 'captionBeginsWith', args: ['?a'] }] }))
    expect(labels(r, r.rows.items)).toEqual(['Banana'])
  })

  it('sorts months by the custom list', () => {
    const rows: CellValue[][] = [['M', 'V'], ['Mar', 1], ['jan', 2], ['Apple', 3], ['Feb', 4]]
    const r = pivot((t) => ({ ...t, rows: [0], values: [{ field: 1 }] }), rows)
    expect(labels(r, r.rows.items)).toEqual(['jan', 'Feb', 'Mar', 'Apple'])
  })

  it('places the value fields on the rows at a position', () => {
    const r = pivot((t) => ({ ...t, rows: [0, 'values'], values: [{ field: 3 }, { field: 2, function: 'count' }] }))
    expect(labels(r, r.rows.items)[0]).toEqual(['East', ['Σ0', 'Σ1']])
  })
})
