/**
 * PivotTables read from the books Excel for Windows saved (`test/excel/win/pivot/basic`, built through COM from the
 * descriptions of `bench/coverage/pivot/basic.ts`), against what Excel showed of each (`basic.json`): the PivotTable's
 * range and the top left of its values, the label of each saved line where Excel showed it, its fields and value fields
 * as the case built them; and GETPIVOTDATA of every line by its items gives the value Excel showed there.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { compute, type BookInput, type CellValue, type PivotCache, type PivotTable, type PivotValue } from '../src/index.js'
import { readXlsx, type XlsxRead } from '../src/xlsx.js'
import { openKept } from '../src/xlsx-keep.js'
import { formatAddress } from '../src/model.js'
import { parseRange } from '../src/structure/area.js'
import { ERROR_TEXT } from '../src/syntax.js'
import { buildPivotCases } from '../bench/coverage/pivot/basic.js'
import type { PivotAnswer, PivotArea } from '../bench/coverage/pivot/cases.js'

const dir = new URL('./excel/win/pivot/', import.meta.url)
const answers = (JSON.parse(readFileSync(new URL('basic.json', dir), 'utf8')) as { answers: Record<string, PivotAnswer> }).answers
const cases = buildPivotCases()
const read = async (id: string): Promise<XlsxRead> => readXlsx(readFileSync(new URL(`basic/${id}.xlsx`, dir)))

/** A cache item as GETPIVOTDATA's argument: text quoted, a date as its serial number, a blank as `(blank)`. */
function literal(v: PivotValue): string {
  if (v === null) return '"(blank)"'
  if (typeof v === 'string') return `"${v.replaceAll('"', '""')}"`
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (typeof v === 'number') return String(v)
  if ('date' in v) return String(Math.round((Date.parse(`${v.date}Z`) - Date.UTC(1899, 11, 30)) / 864e5))
  return `"${ERROR_TEXT[v.error]}"`
}

/** Each saved line of an axis with the fields and items it names (cache items), and its value field. */
function lines(t: PivotTable, c: PivotCache, axis: 'rows' | 'columns') {
  const fields = t[axis] ?? []
  let prev: number[] = []
  return (axis === 'rows' ? t.rowLines : t.columnLines)!.map((l) => {
    const items = [...prev.slice(0, l.repeat ?? 0), ...l.items]
    prev = items
    const depth = l.type === 'grand' ? 0 : items.length
    const named: [string, PivotValue][] = []
    let value = l.value ?? 0
    for (let k = 0; k < depth; k++) {
      const f = fields[k]
      if (f === 'values') value = items[k]
      else named.push([c.fields[f].name, c.fields[f].items![t.fields[f].items![items[k]].item!]])
    }
    return { type: l.type, named, value, depth }
  })
}

const near = (a: unknown, b: unknown) => (typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b)) : a === b || JSON.stringify(a) === JSON.stringify(b))

describe('PivotTables of the books Excel for Windows saved', () => {
  it('the place, the top left of the values and each line of every PivotTable of the set basic, and GETPIVOTDATA of each cell of its values', async () => {
    const bad: string[] = []
    let cells = 0
    for (const c of cases) {
      const r = await read(c.id)
      const shown = answers[c.id].phases[0].pivots.PT1 as PivotArea
      const sheet = r.book.sheets.find((s) => s.name === 'P')!
      const t = sheet.pivotTables![0]
      const cache = r.book.pivotCaches!.find((x) => x.id === t.cache)!
      if (t.location.range !== shown.table) bad.push(`${c.id}: range ${t.location.range} vs ${shown.table}`)
      const range = parseRange(t.location.range)!
      const area = parseRange(shown.area)!
      const body = shown.body ? parseRange(shown.body) : null
      const top = range.r1 + t.location.firstDataRow
      const left = range.c1 + t.location.firstDataColumn
      // without a value field Excel's values are the row labels
      if (body && t.values && (body.r1 !== top || body.c1 !== left)) bad.push(`${c.id}: values at ${formatAddress(top, left)} vs ${shown.body}`)
      const at = (row: number, col: number): CellValue => shown.values[row - area.r1]?.[col - area.c1] ?? null
      const rows = lines(t, cache, 'rows')
      const columns = lines(t, cache, 'columns')
      // a row's label: the innermost item it names, a value field's name, Grand Total
      rows.forEach((l, k) => {
        const label = at(top + k, range.c1)
        const want = l.type === 'grand' ? (t.rows?.includes('values') ? `Total ${t.values![l.value].name}` : 'Grand Total') : l.depth > l.named.length && (t.rows ?? []).indexOf('values') === l.depth - 1 ? t.values![l.value].name : l.named.at(-1)?.[1]
        const text = want === null ? '(blank)' : typeof want === 'object' && want ? ('error' in want ? ERROR_TEXT[want.error] : undefined) : typeof want === 'boolean' ? String(want).toUpperCase() : want
        if (text !== undefined && t.rows?.length && !near(typeof label === 'object' && label ? ERROR_TEXT[label.error] : String(label).toUpperCase(), String(text).toUpperCase())) bad.push(`${c.id}: row ${k} ${label} vs ${text}`)
      })
      if (!t.values?.length) continue
      const formulas: Record<string, { formula: string }> = {}
      const wanted: [string, CellValue][] = []
      // a line of an outer item where the value fields lie across the rows shows none
      const above = (l: (typeof rows)[number]) => (t.rows ?? []).includes('values') && l.type === undefined && l.depth < t.rows!.length
      rows.forEach((rl, i) =>
        columns.forEach((cl, j) => {
          if (above(rl)) return
          const v = at(top + i, left + j)
          const pairs = [...rl.named, ...cl.named].flatMap(([f, item]) => [`"${f}"`, literal(item)])
          const value = (t.rows ?? []).includes('values') ? rl.value : cl.value
          const a = formatAddress(wanted.length, 0)
          formulas[a] = { formula: `GETPIVOTDATA("${t.values![value].name}",P!${formatAddress(range.r1, range.c1)}${pairs.map((p) => `,${p}`).join('')})` }
          wanted.push([a, v])
        }),
      )
      const book: BookInput = { ...r.book, sheets: [...r.book.sheets, { name: 'Check', cells: formulas }] }
      const out = compute(book).sheets.at(-1)!.cells
      for (const [a, v] of wanted) {
        cells++
        // a value Excel showed as nothing is 0; an error it shows by its text (`err` of shown-errors) is the error
        const got = out[a].value
        if (!near(got, v ?? 0) && !(typeof got === 'object' && got && typeof v === 'string')) bad.push(`${c.id}: ${formulas[a].formula} = ${JSON.stringify(got)} vs ${JSON.stringify(v)}`)
      }
    }
    expect(bad).toEqual([])
    expect(cells).toBeGreaterThan(400)
  }, 60_000)

  it("each PivotTable's fields and value fields as its case built them, with no issue of PivotTables, the same through openKept", async () => {
    const fn = { sum: 'Sum', count: 'Count', average: 'Average', max: 'Max', min: 'Min', product: 'Product', countNums: 'Count', stdDev: 'StdDev', stdDevp: 'StdDevp', var: 'Var', varp: 'Varp' }
    for (const c of cases) {
      const data = readFileSync(new URL(`basic/${c.id}.xlsx`, dir))
      const r = await readXlsx(data)
      expect(r.issues.filter((i) => i.code === 'pivotTable')).toEqual([])
      expect((await openKept(data)).book.sheets.map((s) => s.pivotTables)).toEqual(r.book.sheets.map((s) => s.pivotTables))
      const t = r.book.sheets.find((s) => s.name === 'P')!.pivotTables![0]
      const cache = r.book.pivotCaches!.find((x) => x.id === t.cache)!
      const names = (fs: (number | 'values')[] | undefined) => {
        const out = fs?.flatMap((f) => (f === 'values' ? [] : [cache.fields[f].name]))
        return out?.length ? out : undefined
      }
      const ops = Object.assign({}, ...c.ops) as { rows?: string[]; columns?: string[]; pages?: string[] }
      // a row field made a value field leaves the rows (Count of Region by Region): Excel saved Region on Values alone
      if (c.id.startsWith('region-count')) delete ops.rows
      expect([names(t.rows), names(t.columns), names(t.pages?.map((p) => p.field))]).toEqual([ops.rows, ops.columns, ops.pages])
      const values = c.ops.flatMap((o) => ('value' in o ? [o.value] : []))
      expect(t.values?.map((v) => [cache.fields[v.field].name, v.function ?? 'sum', v.name])).toEqual(values.length ? values.map((v) => [v.field, v.fn ?? 'sum', v.caption ?? `${fn[v.fn ?? 'sum']} of ${v.field}`]) : undefined)
      expect(t.cache).toBe(cache.id)
      // Excel 365 hides the Values row (Excel 2010's hideValuesRow)
      expect(t.valuesRow).toBeUndefined()
    }
    // a PivotTable Excel for Mac laid out from a part without the extension shows it
    const mac = await readXlsx(readFileSync(new URL('./excel/keep/made/pivot.xlsx', import.meta.url)))
    expect(mac.book.sheets[1].pivotTables![0].valuesRow).toBe(true)
  }, 60_000)
})
