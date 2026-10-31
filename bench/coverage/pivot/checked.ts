/**
 * Fields placed from Excel's PivotTable Fields list, for Excel for Windows on a screen (`bench/windows/run.ts
 * pivot-fields <set>`): a case's fields, each checked in the list (or sent to the columns by its menu) on an empty
 * `PT1` at `P!A3` over the rows of `data.ts`, in turn; the answer is the PivotTable as the book Excel saved holds it:
 * the fields of its rows and columns, its value fields with their functions, and the fields grouping added. The
 * engine's prediction places each field as `suggestPivotField` says, through `editPivotTable` with `groupDates`
 * (`npx tsx bench/coverage/pivot/checked.ts <set>` writes `<set>.predicted.json` beside the set).
 */
import { writeFileSync } from 'node:fs'
import type { BookInput, PivotTable, SheetInput } from '../../../src/contract.js'
import { formatAddress } from '../../../src/model.js'
import { parseRange } from '../../../src/structure/area.js'
import { createState } from '../../../src/engine.js'
import { registerPivot, suggestPivotField } from '../../../src/index.js'
import { pivotEngine } from '../../../src/pivot/index.js'
import { clockOptions } from '../../../src/result.js'
import { pivotOp } from '../../../src/structure/pivot-ops.js'
import { stateValues } from '../../../src/structure/values.js'
import { startBook } from './drive.js'

/** A field placed: checked in the list, or sent to the columns by its menu. */
export interface Placed {
  name: string
  to?: 'columns'
}

/** A cell put on `P` before the PivotTable is made: a value, a formula, a number format. */
export interface Under {
  cell: string
  value?: string | number
  formula?: string
  format?: string
}

export interface CheckedCase {
  id: string
  /** The rows (`DATA`, `kinds` by default). */
  data?: string
  fields: Placed[]
  /** Cells, cells merged, a table (with headers) on `P` before the PivotTable is made. */
  under?: { cells?: Under[]; merge?: string; table?: string }
  /** A window Excel opens on placing a field answered: `yes` (Return) or `no` (Escape); without it the case fails. */
  answer?: 'yes' | 'no'
  /** Cells of `P` read in the book saved (`A21:C35`): each one's value or formula and number format; and `P`'s merges and tables. */
  watch?: string
}

export interface CheckedSet {
  buildCheckedCases(): CheckedCase[]
}

/** A PivotTable as placing its fields left it: the fields on its axes by name (`Σ` for Σ Values), its value fields, the fields grouping added with what they group by. */
export interface CheckedAnswer {
  rows: string[]
  columns: string[]
  values: { field: string; function: string }[]
  groups: { field: string; base: string; by: string }[]
  /** The PivotTable's range, and the cells watched: by address, `=formula` or the value, with `| format`. */
  range?: string
  cells?: Record<string, string>
  merges?: string[]
  tables?: string[]
}

/** `PT1` of `P` as the answer reads it. */
export function checkedOf(book: BookInput, watch?: string): CheckedAnswer {
  const sheet = book.sheets.find((s) => s.name === 'P')!
  const t = sheet.pivotTables!.find((p) => p.name === 'PT1') as PivotTable
  const fields = book.pivotCaches!.find((c) => c.id === t.cache)!.fields
  const names = (axis?: PivotTable['rows']) => (axis ?? []).map((f) => (f === 'values' ? 'Σ' : fields[f].name))
  return {
    rows: names(t.rows),
    columns: names(t.columns),
    values: (t.values ?? []).map((d) => ({ field: fields[d.field].name, function: d.function ?? 'sum' })),
    groups: fields.flatMap((f, i) => (f.group?.range ? [{ field: f.name, base: fields[f.group.base ?? i].name, by: f.group.range.by }] : [])),
    ...(watch ? { range: t.location.range, cells: watched(sheet, watch), merges: sheet.merges ?? [], tables: (sheet.tables ?? []).map((x) => x.range) } : {}),
  }
}

/** The cells of `range` that hold something or have a number format, each as `=formula` or its value, `| format` after. */
function watched(sheet: SheetInput, range: string): Record<string, string> {
  const x = parseRange(range)!
  const out: Record<string, string> = {}
  for (let r = x.r1; r <= x.r2; r++)
    for (let c = x.c1; c <= x.c2; c++) {
      const a = formatAddress(r, c)
      const cell = sheet.cells[a]
      if (!cell || (cell.value === undefined && cell.formula === undefined && cell.numberFormat === undefined)) continue
      const v = cell.formula !== undefined ? `=${cell.formula}` : JSON.stringify(cell.value ?? null)
      out[a] = cell.numberFormat === undefined ? v : `${v} | ${typeof cell.numberFormat === 'string' ? cell.numberFormat : JSON.stringify(cell.numberFormat)}`
    }
  return out
}

const CLOCK = clockOptions({ now: '2026-10-31', seed: 0 })

/** The engine's placing of a case's fields, or its refusal. */
export function predictChecked(c: CheckedCase): CheckedAnswer | { refused: string } {
  const book = predictBook(c)
  return typeof book === 'string' ? { refused: book } : checkedOf(book, c.watch)
}

/** The book with a case's fields placed by the engine, or its refusal's message. */
export function predictBook(c: CheckedCase): BookInput | string {
  registerPivot(pivotEngine)
  let book = withUnder(startBook({ id: c.id, data: c.data ?? 'kinds', ops: [] }), c.under)
  for (const p of c.fields) {
    const cache = book.pivotCaches![0]
    const t = book.sheets.find((s) => s.name === 'P')!.pivotTables![0]
    const f = cache.fields.findIndex((x) => x.name === p.name)
    const s = suggestPivotField(cache, f)
    const set = p.to === 'columns' ? { columns: [...(t.columns ?? []), f] } : s.axis === 'values' ? { values: [...(t.values ?? []), { field: f, function: s.function }] } : { rows: [...(t.rows ?? []), f] }
    const r = pivotOp(book, { op: 'editPivotTable', sheet: 'P', pivot: 'PT1', set, groupDates: true }, stateValues(createState(book, CLOCK)))
    if (!r.ok) return r.message
    book = r.book
  }
  return book
}

/** The book with a case's cells, merge and table on `P`, as the step puts them there. */
function withUnder(book: BookInput, under: CheckedCase['under']): BookInput {
  if (!under) return book
  return {
    ...book,
    sheets: book.sheets.map((s) => {
      if (s.name !== 'P') return s
      const cells = { ...s.cells }
      for (const u of under.cells ?? []) cells[u.cell] = { ...(u.formula !== undefined ? { formula: u.formula } : u.value !== undefined ? { value: u.value } : {}), ...(u.format !== undefined ? { numberFormat: u.format } : {}) }
      const x = under.table ? parseRange(under.table)! : null
      const tables = x ? [{ name: 'Table1', range: under.table!, columns: Array.from({ length: x.c2 - x.c1 + 1 }, (_, k) => ({ name: String(cells[formatAddress(x.r1, x.c1 + k)]?.value ?? `Column${k + 1}`) })) }] : undefined
      return { ...s, cells, ...(under.merge ? { merges: [under.merge] } : {}), ...(tables ? { tables } : {}) }
    }),
  }
}

if (process.argv[1]?.endsWith('checked.ts')) {
  const set = process.argv[2]
  const { buildCheckedCases } = (await import(`./${set}.ts`)) as CheckedSet
  const out = Object.fromEntries(buildCheckedCases().map((c) => [c.id, predictChecked(c)]))
  writeFileSync(new URL(`${set}.predicted.json`, import.meta.url), JSON.stringify(out, null, 1) + '\n')
  console.log(`${Object.keys(out).length} cases`)
}
