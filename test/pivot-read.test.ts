/**
 * PivotTables read from the books Excel for Windows saved (`test/excel/win/pivot/<set>`, built through COM from the
 * descriptions of `bench/coverage/pivot/`), against what Excel showed of each (`<set>.json`): the PivotTable's range and
 * the top left of its values, the label of each saved line (the set basic), its fields and value fields as the case
 * built them; and GETPIVOTDATA of every line by its items gives the value Excel showed there.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { compute, type BookInput, type CellValue, type PivotCache, type PivotFilterType, type PivotTable, type PivotValue } from '../src/index.js'
import { readXlsx, type XlsxRead } from '../src/xlsx.js'
import { openKept } from '../src/xlsx-keep.js'
import { formatAddress } from '../src/model.js'
import { parseRange } from '../src/structure/area.js'
import { ERROR_TEXT } from '../src/syntax.js'
import { buildPivotCases } from '../bench/coverage/pivot/basic.js'
import type { PivotAnswer, PivotArea } from '../bench/coverage/pivot/cases.js'

const dir = new URL('./excel/win/pivot/', import.meta.url)
const cases = buildPivotCases()

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
  return (axis === 'rows' ? t.rowLines : t.columnLines ?? [{ items: [] }])!.map((l) => {
    const items = [...prev.slice(0, l.repeat ?? 0), ...l.items]
    prev = items
    const depth = l.type === 'grand' ? 0 : items.length
    const named: [string, PivotValue][] = []
    let value = l.value ?? 0
    for (let k = 0; k < depth; k++) {
      const f = fields[k]
      if (f === 'values') value = items[k]
      else {
        // an item by its caption where it has one, else its value (of its group where the field is grouped)
        const it = t.fields[f].items![items[k]]
        named.push([t.fields[f].name ?? c.fields[f].name, it.caption ?? (c.fields[f].group?.items ?? c.fields[f].items)![it.item!]])
      }
    }
    return { type: l.type, named, value, depth }
  })
}

const near = (a: unknown, b: unknown) => (typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b)) : a === b || JSON.stringify(a) === JSON.stringify(b))

/** The sets of PivotTables Excel for Windows saved. */
const SETS = ['basic', 'layout', 'calc', 'edit', 'filter', 'refresh', 'sort', 'layout-held', 'filter-held', 'sort-held', 'edit-held']

/**
 * One PivotTable against what Excel showed of it: its range, the top left of its values, with `labels` each row's label
 * (the innermost item it names, a value field's name, Grand Total); and GETPIVOTDATA of each value Excel showed (the
 * first line of the same items where several show them: a field's subtotals of several functions); the cells checked.
 */
function checked(r: XlsxRead, sheet: string, t: PivotTable, shown: PivotArea, labels: boolean, id: string, bad: string[]): number {
  const cache = r.book.pivotCaches!.find((x) => x.id === t.cache)!
  if (t.location.range !== shown.table) bad.push(`${id}: range ${t.location.range} vs ${shown.table}`)
  const range = parseRange(t.location.range)!
  const area = parseRange(shown.area)!
  const body = shown.body ? parseRange(shown.body) : null
  const top = range.r1 + t.location.firstDataRow
  const left = range.c1 + t.location.firstDataColumn
  // without a value field Excel's values are the row labels
  if (body && t.values && (body.r1 !== top || body.c1 !== left)) bad.push(`${id}: values at ${formatAddress(top, left)} vs ${shown.body}`)
  const at = (row: number, col: number): CellValue => shown.values[row - area.r1]?.[col - area.c1] ?? null
  const rows = lines(t, cache, 'rows')
  const columns = lines(t, cache, 'columns')
  if (labels)
    rows.forEach((l, k) => {
      const label = at(top + k, range.c1)
      const want = l.type === 'grand' ? (t.rows?.includes('values') ? `Total ${t.values![l.value].name}` : 'Grand Total') : l.depth > l.named.length && (t.rows ?? []).indexOf('values') === l.depth - 1 ? t.values![l.value].name : l.named.at(-1)?.[1]
      const text = want === null ? '(blank)' : typeof want === 'object' && want ? ('error' in want ? ERROR_TEXT[want.error] : undefined) : typeof want === 'boolean' ? String(want).toUpperCase() : want
      if (text !== undefined && t.rows?.length && !near(typeof label === 'object' && label ? ERROR_TEXT[label.error] : String(label).toUpperCase(), String(text).toUpperCase())) bad.push(`${id}: row ${k} ${label} vs ${text}`)
    })
  if (!t.values?.length) return 0
  const formulas: Record<string, { formula: string }> = {}
  const wanted: [string, CellValue][] = []
  const seen = new Set<string>()
  const ref = `'${sheet}'!${formatAddress(range.r1, range.c1)}`
  rows.forEach((rl, i) =>
    columns.forEach((cl, j) => {
      const v = at(top + i, left + j)
      if (v === null || rl.type === 'blank' || cl.type === 'blank') return
      const pairs = [...rl.named, ...cl.named].flatMap(([f, item]) => [`"${f}"`, literal(item)])
      const value = (t.rows ?? []).includes('values') ? rl.value : cl.value
      const formula = `GETPIVOTDATA("${t.values![value].name}",${ref}${pairs.map((p) => `,${p}`).join('')})`
      if (seen.has(formula)) return
      seen.add(formula)
      const a = formatAddress(wanted.length, 0)
      formulas[a] = { formula }
      wanted.push([a, v])
    }),
  )
  const book: BookInput = { ...r.book, sheets: [...r.book.sheets, { name: 'Check', cells: formulas }] }
  const out = compute(book).sheets.at(-1)!.cells
  for (const [a, v] of wanted) {
    // an error Excel shows by a text of the PivotTable's (`err`) is the error
    const got = out[a].value
    if (!near(got, v) && !(typeof got === 'object' && got && typeof v === 'string')) bad.push(`${id}: ${formulas[a].formula} = ${JSON.stringify(got)} vs ${JSON.stringify(v)}`)
  }
  return wanted.length
}

describe('PivotTables of the books Excel for Windows saved', () => {
  it('the place, the top left of the values and each line of every PivotTable of every set, after each phase, and GETPIVOTDATA of each value Excel showed', async () => {
    const bad: string[] = []
    let cells = 0
    for (const set of SETS) {
      const all = (JSON.parse(readFileSync(new URL(`${set}.json`, dir), 'utf8')) as { answers: Record<string, PivotAnswer> }).answers
      for (const [id, answer] of Object.entries(all))
        for (const [k, phase] of answer.phases.entries()) {
          const file = new URL(`${set}/${id}${k < answer.phases.length - 1 ? `-${k + 1}` : ''}.xlsx`, dir)
          if (!existsSync(file)) continue
          const r = await readXlsx(readFileSync(file))
          for (const [name, shown] of Object.entries(phase.pivots)) {
            if ('error' in shown) continue
            // a PivotTable by its name, or by its sheet and name where the set has several sheets of them (`P!PT1`)
            const [on, alone] = name.includes('!') ? name.split('!') : [undefined, name]
            const sheet = r.book.sheets.find((x) => (on === undefined || x.name === on) && x.pivotTables?.some((t) => t.name === alone))
            if (!sheet) {
              bad.push(`${set}/${id} ${k}: no ${name}`)
              continue
            }
            cells += checked(r, sheet.name, sheet.pivotTables!.find((t) => t.name === alone)!, shown, set === 'basic', `${set}/${id} ${k} ${name}`, bad)
          }
        }
    }
    expect(bad).toEqual([])
    expect(cells).toBeGreaterThan(2000)
  }, 120_000)

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

  it('filters and sorting as the cases of the sets filter and sort made them', async () => {
    const FILE: Record<string, [type: PivotFilterType, bottom?: true]> = {
      TopCount: ['count'], BottomCount: ['count', true], TopPercent: ['percent'], BottomPercent: ['percent', true], TopSum: ['sum'], BottomSum: ['sum', true],
      SpecificDate: ['dateEqual'], NotSpecificDate: ['dateNotEqual'], Before: ['dateOlderThan'], BeforeOrEqualTo: ['dateOlderThanOrEqual'], After: ['dateNewerThan'], AfterOrEqualTo: ['dateNewerThanOrEqual'],
      DateBetween: ['dateBetween'], DateNotBetween: ['dateNotBetween'], AllDatesInPeriodJanuary: ['M1'], AllDatesInPeriodQuarter2: ['Q2'],
    }
    // the object model's names of the others: Caption and Value, Is and DoesNot (CaptionIsGreaterThanOrEqualTo, captionGreaterThanOrEqual)
    const fileType = (t: string): [PivotFilterType, true?] => FILE[t] ?? [t.replace(/^(Caption|Value)/, (m) => m.toLowerCase()).replace('DoesNot', 'Not').replace('Is', '').replace(/Equals$/, 'Equal').replace('OrEqualTo', 'OrEqual').replace(/Not(Begin|End)With$/, 'Not$1sWith').replace(/NotContain$/, 'NotContains') as PivotFilterType]
    const iso = (v: unknown) => (typeof v === 'object' && v && 'd' in v ? new Date(Date.UTC(1899, 11, 30) + (v as { d: number }).d * 864e5).toISOString().slice(0, 10) : v)
    for (const set of ['filter', 'sort']) {
      const { buildPivotCases: build } = (await import(`../bench/coverage/pivot/${set}.ts`)) as { buildPivotCases: () => typeof cases }
      for (const c of build()) {
        const file = new URL(`${set}/${c.id}.xlsx`, dir)
        if (c.ops.some((o) => 'read' in o || 'refresh' in o || 'item' in o) || !existsSync(file)) continue
        const r = await readXlsx(readFileSync(file))
        const t = r.book.sheets.flatMap((x) => x.pivotTables ?? [])[0]
        const cache = r.book.pivotCaches!.find((x) => x.id === t.cache)!
        const field = (name: string) => cache.fields.findIndex((f) => f.name === name)
        const value = (name?: string) => (name === undefined ? {} : { value: t.values!.findIndex((v) => v.name === name) })
        const filters = c.ops.flatMap((o) => ('filter' in o ? [o] : []))
        const want = filters.map((o) => {
          const [type, bottom] = fileType(o.type)
          const args = [o.value1, o.value2].filter((v) => v !== undefined).map(iso)
          return { field: field(o.filter), type, ...(args.length ? { args } : {}), ...value(o.by), ...(bottom ? { bottom } : {}) }
        })
        expect([c.id, [...(t.filters ?? [])].sort((a, b) => a.field - b.field || a.type.localeCompare(b.type))]).toEqual([c.id, want.sort((a, b) => a.field - b.field || a.type.localeCompare(b.type))])
        const sorts = c.ops.flatMap((o) => ('sort' in o ? [o] : []))
        expect([c.id, t.fields.flatMap((f, k) => (f.sort ? [[k, f.sort]] : []))]).toEqual([c.id, sorts.map((o) => [field(o.sort), { order: o.order, ...(o.by ? { by: value(o.by) } : {}) }])])
      }
    }
  }, 60_000)
})
