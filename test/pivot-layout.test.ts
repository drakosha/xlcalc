/**
 * The layout of a PivotTable from its calculation (`layoutPivot`), by hand over a small cache: the lines of each axis,
 * the place and the cells in the three forms, subtotals at the top and below, blank lines, value fields across either
 * axis, the Filters area, the text for errors and empty cells, the words of ru-RU; and the step that writes laid out
 * PivotTables to their sheets (`placePivots`), refusing an area over data, a table or another PivotTable.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput, PivotCache, PivotTable } from '../src/contract.js'
import { formatAddress } from '../src/model.js'
import { layoutPivot, type PivotLaidOut } from '../src/pivot/layout.js'
import { placePivots } from '../src/pivot/place.js'
import type { PivotNode, PivotResult } from '../src/pivot/result.js'

/** Region, Product and a number: records summed by hand into the calculation's tree and values. */
const cache: PivotCache = {
  id: 1,
  source: { type: 'range', sheet: 'Data', range: 'A1:C5' },
  fields: [{ name: 'Region', items: ['North', 'South'] }, { name: 'Product', items: ['A', 'B'] }, { name: 'Amount' }, { name: 'Note', items: ['x', null] }],
  records: [
    [0, 0, 1, 0],
    [0, 1, 2, 0],
    [1, 0, 4, 1],
  ],
}

const fields = (): PivotTable['fields'] => [{ items: [{ item: 0 }, { item: 1 }] }, { items: [{ item: 0 }, { item: 1 }] }, {}, { items: [{ item: 0 }, { item: 1 }] }]

const table = (more: Partial<PivotTable> = {}): PivotTable => ({
  name: 'PT',
  cache: 1,
  location: { range: 'A3:B4', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 },
  fields: fields(),
  values: [{ field: 2, name: 'Sum of Amount' }],
  ...more,
})

/** The calculation of `t` over the records: items with records, in the field's order; values summed, null where none (the count for a second value field). */
function calc(t: PivotTable): PivotResult {
  const recs = cache.records!
  const axis = (fs: PivotTable['rows']) => {
    const list = fs ?? []
    const real = list.filter((f): f is number => f !== 'values')
    const build = (k: number, path: number[]): PivotNode[] => {
      if (k === list.length) return []
      const f = list[k]
      if (f === 'values') return (t.values ?? []).map((_, v) => ({ field: 'values', item: v, children: build(k + 1, path) }))
      const at = real.indexOf(f)
      return t.fields[f].items!.flatMap((it, i) => {
        const p = [...path, i]
        const some = recs.some((r) => p.every((x, j) => r[real[j]] === t.fields[real[j]].items![x].item))
        return some ? [{ field: f, item: i, children: build(k + 1, p) }] : []
      }).filter(() => at >= 0)
    }
    return { fields: list, items: build(0, []), real }
  }
  const rows = axis(t.rows)
  const cols = axis(t.columns)
  return {
    table: t,
    cache,
    rows: { fields: rows.fields, items: rows.items },
    columns: { fields: cols.fields, items: cols.items },
    pages: (t.pages ?? []).map((p) => ({ field: p.field, items: p.item === undefined ? [] : [p.item], all: p.item === undefined })),
    value(r, c, v) {
      const hit = recs.filter((rec) => r.every((x, j) => rec[rows.real[j]] === t.fields[rows.real[j]].items![x].item) && c.every((x, j) => rec[cols.real[j]] === t.fields[cols.real[j]].items![x].item))
      if (!hit.length) return null
      return v === 0 ? hit.reduce((s, rec) => s + (rec[2] as number), 0) : hit.length
    },
    hasData: () => true,
  }
}

/** The area's cells as rows of values, `''` for an empty cell, an error by its code. */
function grid(l: PivotLaidOut): (string | number | boolean)[][] {
  const by = new Map(l.cells.map((c) => [formatAddress(c.row, c.col), c]))
  const out: (string | number | boolean)[][] = []
  for (let r = l.area.r1; r <= l.area.r2; r++) {
    const row: (string | number | boolean)[] = []
    for (let c = l.area.c1; c <= l.area.c2; c++) {
      const x = by.get(formatAddress(r, c))
      row.push(x?.error ? `#${x.error}` : typeof x?.value === 'object' ? x.value.date : (x?.value ?? ''))
    }
    out.push(row)
  }
  return out
}

const laid = (more: Partial<PivotTable>, locale?: string) => layoutPivot(calc(table(more)), { locale })

describe('layoutPivot', () => {
  it('lays two row fields out in compact form, subtotals at the top', () => {
    const l = laid({ rows: [0, 1] })
    expect(l.rowLines).toEqual([{ items: [0] }, { items: [0], repeat: 1 }, { items: [1], repeat: 1 }, { items: [1] }, { items: [0], repeat: 1 }, { type: 'grand', items: [0] }])
    expect(l.columnLines).toEqual([{ items: [] }])
    expect(l.location).toEqual({ range: 'A3:B9', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1, areas: { table: 'A3:B9', columnHeaders: 'A3:B3', rowLabels: 'A4:A9', data: 'B4:B9' } })
    expect(grid(l)).toEqual([['Row Labels', 'Sum of Amount'], ['North', 3], ['A', 1], ['B', 2], ['South', 4], ['A', 4], ['Grand Total', 7]])
    expect(l.cells.find((c) => c.row === 4 && c.col === 0)?.indent).toBe(1)
  })

  it('gives each field a column in outline form, and in tabular form puts the item on its first line and the subtotal below', () => {
    expect(grid(laid({ rows: [0, 1], layout: 'outline' }))).toEqual([
      ['Region', 'Product', 'Sum of Amount'], ['North', '', 3], ['', 'A', 1], ['', 'B', 2], ['South', '', 4], ['', 'A', 4], ['Grand Total', '', 7],
    ])
    const t = laid({ rows: [0, 1], layout: 'tabular' })
    expect(t.rowLines).toEqual([{ items: [0, 0] }, { items: [1], repeat: 1 }, { type: 'subtotal', items: [0] }, { items: [1, 0] }, { type: 'subtotal', items: [1] }, { type: 'grand', items: [0] }])
    expect(grid(t)).toEqual([['Region', 'Product', 'Sum of Amount'], ['North', 'A', 1], ['', 'B', 2], ['North Total', '', 3], ['South', 'A', 4], ['South Total', '', 4], ['Grand Total', '', 7]])
  })

  it('repeats labels in outline and tabular forms, and puts subtotals below, of several functions, and blank lines', () => {
    const repeated = laid({ rows: [0, 1], layout: 'tabular', repeatLabels: true })
    expect(grid(repeated).slice(1, 3)).toEqual([['North', 'A', 1], ['North', 'B', 2]])
    const f = fields()
    f[0] = { ...f[0], subtotalTop: false, blankLine: true }
    expect(laid({ rows: [0, 1], fields: f }).rowLines).toEqual([
      { items: [0] }, { items: [0], repeat: 1 }, { items: [1], repeat: 1 }, { type: 'subtotal', items: [0] }, { type: 'blank', items: [0] },
      { items: [1] }, { items: [0], repeat: 1 }, { type: 'subtotal', items: [1] }, { type: 'blank', items: [1] }, { type: 'grand', items: [0] },
    ])
    const g = fields()
    g[0] = { ...g[0], subtotals: ['sum', 'count'] }
    const two = laid({ rows: [0, 1], fields: g })
    expect(two.rowLines.slice(0, 5)).toEqual([{ items: [0] }, { items: [0], repeat: 1 }, { items: [1], repeat: 1 }, { type: 'sum', items: [0] }, { type: 'count', items: [0] }])
    expect(grid(two).slice(1, 6)).toEqual([['North', ''], ['A', 1], ['B', 2], ['North Sum', 3], ['North Count', 3]])
    const none = fields()
    none[0] = { ...none[0], subtotals: [] }
    expect(grid(laid({ rows: [0, 1], fields: none })).slice(1, 3)).toEqual([['North', ''], ['A', 1]])
  })

  it('lays the columns out under their caption, subtotals after each group, the grand total at the right', () => {
    const l = laid({ rows: [3], columns: [0, 1] })
    expect(l.columnLines).toEqual([{ items: [0, 0] }, { items: [1], repeat: 1 }, { type: 'subtotal', items: [0] }, { items: [1, 0] }, { type: 'subtotal', items: [1] }, { type: 'grand', items: [0] }])
    expect(l.location).toEqual({ range: 'A3:G8', firstHeaderRow: 1, firstDataRow: 3, firstDataColumn: 1, areas: { table: 'A3:G8', columnHeaders: 'A3:G5', rowLabels: 'A6:A8', data: 'B6:G8' } })
    expect(grid(l)).toEqual([
      ['Sum of Amount', 'Column Labels', '', '', '', '', ''],
      ['', 'North', '', 'North Total', 'South', 'South Total', 'Grand Total'],
      ['Row Labels', 'A', 'B', '', 'A', '', ''],
      ['x', 1, 2, 3, '', '', 3],
      ['(blank)', '', '', '', 4, 4, 4],
      ['Grand Total', 1, 2, 3, 4, 4, 7],
    ])
  })

  it('lays two value fields across the rows and across the columns', () => {
    const values = [{ field: 2, name: 'Sum of Amount' }, { field: 2, name: 'Count of Amount', function: 'count' as const }]
    const rows = laid({ rows: [0, 'values'], values, valuesOnRows: true })
    expect(rows.rowLines).toEqual([{ items: [0] }, { items: [0], repeat: 1 }, { items: [1], repeat: 1, value: 1 }, { items: [1] }, { items: [0], repeat: 1 }, { items: [1], repeat: 1, value: 1 }, { type: 'grand', items: [0] }, { type: 'grand', items: [0], value: 1 }])
    expect(grid(rows)).toEqual([['Row Labels', ''], ['North', ''], ['Sum of Amount', 3], ['Count of Amount', 2], ['South', ''], ['Sum of Amount', 4], ['Count of Amount', 1], ['Total Sum of Amount', 7], ['Total Count of Amount', 3]])
    const cols = laid({ rows: [0], columns: ['values'], values })
    expect(cols.location).toEqual({ range: 'A3:C6', firstHeaderRow: 0, firstDataRow: 1, firstDataColumn: 1, areas: { table: 'A3:C6', columnHeaders: 'A3:C3', rowLabels: 'A4:A6', data: 'B4:C6' } })
    expect(grid(cols)).toEqual([['Row Labels', 'Sum of Amount', 'Count of Amount'], ['North', 3, 2], ['South', 4, 1], ['Grand Total', 7, 3]])
    // the Values row shown: a caption row over the value fields' names
    const shown = laid({ rows: [0], columns: ['values'], values, valuesRow: true })
    expect(shown.location).toEqual({ range: 'A3:C7', firstHeaderRow: 1, firstDataRow: 2, firstDataColumn: 1, areas: { table: 'A3:C7', columnHeaders: 'A3:C4', rowLabels: 'A5:A7', data: 'B5:C7' } })
    expect(grid(shown).slice(0, 2)).toEqual([['', 'Column Labels', ''], ['Row Labels', 'Sum of Amount', 'Count of Amount']])
  })

  it('lays the Filters area out above the table, moving it down when the rows above are too few', () => {
    const l = laid({ rows: [0], pages: [{ field: 1 }, { field: 3, item: 1 }], location: { range: 'A1:B4', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 } })
    expect(l.location).toEqual({ range: 'A4:B7', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1, pageRows: 2, pageColumns: 1, areas: { page: 'A1:B2', table: 'A4:B7', columnHeaders: 'A4:B4', rowLabels: 'A5:A7', data: 'B5:B7' } })
    expect(l.area).toEqual({ r1: 0, c1: 0, r2: 6, c2: 1 })
    expect(grid(l).slice(0, 3)).toEqual([['Product', '(All)'], ['Note', '(blank)'], ['', '']])
    const across = laid({ rows: [0], pages: [{ field: 1 }, { field: 3 }], pagesAcrossFirst: true, location: { range: 'A5:B8', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 } })
    expect(across.location).toEqual({ range: 'A5:B8', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1, pageRows: 1, pageColumns: 2, areas: { page: 'A3:E3', table: 'A5:B8', columnHeaders: 'A5:B5', rowLabels: 'A6:A8', data: 'B6:B8' } })
    expect(grid(across)[0]).toEqual(['Product', '(All)', '', 'Note', '(All)'])
  })

  it('shows its text for errors and for empty cells, its captions, and the words of ru-RU', () => {
    const l = laid({ rows: [3], columns: [0], empty: { show: true, text: '-' }, captions: { grandTotal: 'All', rowHeader: 'Notes' } })
    // a dash alone is the accounting zero, which Excel shows as 0
    expect(grid(l).slice(1)).toEqual([['Notes', 'North', 'South', 'All'], ['x', 3, 0, 3], ['(blank)', 0, 4, 4], ['All', 3, 4, 7]])
    expect(grid(laid({ rows: [3], columns: [0], empty: { show: true, text: '0' } })).slice(2, 4)).toEqual([['x', 3, 0, 3], ['(blank)', 0, 4, 4]])
    expect(grid(laid({ rows: [3] }, 'ru-RU'))).toEqual([['Названия строк', 'Sum of Amount'], ['x', 3], ['(пусто)', 4], ['Общий итог', 7]])
  })

  it('keeps the place of the drop zones for a PivotTable with no field', () => {
    const l = laid({ values: [] })
    expect(l.location).toEqual({ range: 'A3:C20', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 0, areas: { table: 'A3:C20', data: 'A4:C20' } })
    expect(l.cells).toEqual([])
  })
})

describe('placePivots', () => {
  const book = (cells: BookInput['sheets'][0]['cells'] = {}, more: Partial<BookInput['sheets'][0]> = {}): BookInput => ({
    sheets: [{ name: 'P', cells: { A3: { value: 'old' }, B5: { value: 5, style: 0 }, ...cells }, pivotTables: [table({ location: { range: 'A3:B5', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 } })], ...more }],
    pivotCaches: [cache],
  })

  it('clears the old area, keeping its cells styles, writes the new cells and the layout', () => {
    const t = table({ rows: [3] })
    const l = layoutPivot(calc(t), {})
    const r = placePivots(book(), [{ sheet: 'P', table: t, laid: l }])
    if (!r.ok) throw new Error(r.message)
    const sh = r.book.sheets[0]
    expect(sh.cells.A3).toEqual({ value: 'Row Labels' })
    expect(sh.cells.B5).toEqual({ value: 4, style: 0 })
    expect(placePivots(book({ B6: { value: 1, style: 0 } }), [{ sheet: 'P', table: t, laid: l }], 'refuse').ok).toBe(false)
    // a compact form's row labels left aligned, as Excel keeps them (\`test/pivot-place-looks.test.ts\`)
    expect([sh.cells.A6, r.book.styles]).toEqual([{ value: 'Grand Total', style: 0 }, [{ alignment: { horizontal: 'left' } }]])
    expect(sh.pivotTables![0]).toMatchObject({ rows: [3], location: { range: 'A3:B6' }, rowLines: l.rowLines })
  })

  it('refuses, when asked to, an area over cells that hold something, a table or another PivotTable, changing nothing', () => {
    const t = table({ rows: [0, 1] })
    const l = layoutPivot(calc(t), {})
    const over = placePivots(book({ A9: { value: 1 } }), [{ sheet: 'P', table: t, laid: l }], 'refuse')
    expect(over).toMatchObject({ ok: false, code: 'pivot', message: "there's already data in P!A3:B9, where the PivotTable PT would be laid out" })
    expect(placePivots(book({ A9: { style: 0 } }), [{ sheet: 'P', table: t, laid: l }], 'refuse').ok).toBe(true)
    expect(placePivots(book({}, { tables: [{ name: 'T', range: 'A9:B11', columns: [{ name: 'a' }, { name: 'b' }] }] }), [{ sheet: 'P', table: t, laid: l }], 'refuse')).toMatchObject({ ok: false, code: 'pivot', message: 'the PivotTable PT would overlap the table T' })
    const other = { ...table({ name: 'PT2', location: { range: 'A9:B11', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 } }) }
    const b = book()
    b.sheets[0].pivotTables!.push(other)
    expect(placePivots(b, [{ sheet: 'P', table: t, laid: l }], 'refuse')).toMatchObject({ ok: false, code: 'pivot' })
    // both laid again at once: the second's old area is free, the new areas must not meet
    const t2 = { ...other, rows: [3] as PivotTable['rows'] }
    const l2 = layoutPivot(calc(t2), {})
    expect(placePivots(b, [{ sheet: 'P', table: t, laid: l }, { sheet: 'P', table: t2, laid: l2 }], 'refuse')).toMatchObject({ ok: false, code: 'pivot', message: expect.stringContaining('PT2') })
    const far = { ...t2, location: { ...t2.location, range: 'D10:E13' } }
    const moved = placePivots({ ...b, sheets: [{ ...b.sheets[0], pivotTables: [b.sheets[0].pivotTables![0], far] }] }, [{ sheet: 'P', table: t, laid: l }, { sheet: 'P', table: far, laid: layoutPivot(calc(far), {}) }])
    expect(moved.ok).toBe(true)
  })

  it('shrinks a PivotTable laid out over cells that hold something to #SPILL! at its corner, keeping its new lines', () => {
    const t = table({ rows: [0, 1] })
    const l = layoutPivot(calc(t), {})
    const b = book({ A9: { value: 'kept' } })
    const other = table({ name: 'PT2', location: { range: 'A11:B12', firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 1 } })
    b.sheets[0].pivotTables!.push(other)
    const r = placePivots(b, [{ sheet: 'P', table: t, laid: l }, { sheet: 'P', table: other, laid: layoutPivot(calc(other), {}) }])
    if (!r.ok) throw new Error(r.message)
    const sh = r.book.sheets[0]
    expect(sh.cells).toMatchObject({ A3: { formula: '#SPILL!' }, A9: { value: 'kept' }, B5: { style: 0 }, A11: { value: 'Sum of Amount' }, A12: { value: 7 } })
    expect(sh.cells.B3).toBeUndefined()
    expect(sh.pivotTables![0]).toMatchObject({ location: { range: 'A3', firstHeaderRow: 0, firstDataRow: 0, firstDataColumn: 0 }, rowLines: l.rowLines })
    expect(sh.pivotTables![1].location.range).toBe('A11:A12')
  })
})
