/**
 * Fields placed from Excel's PivotTable Fields list, for Excel for Windows on a screen (`bench/windows/run.ts
 * pivot-fields <set>`): a case's fields, each checked in the list (or sent to the columns by its menu) on an empty
 * `PT1` at `P!A3` over the rows of `data.ts`, in turn; the answer is the PivotTable as the book Excel saved holds it:
 * the fields of its rows and columns, its value fields with their functions, and the fields grouping added. The
 * engine's prediction places each field as `suggestPivotField` says, through `editPivotTable` with `groupDates`
 * (`npx tsx bench/coverage/pivot/checked.ts <set>` writes `<set>.predicted.json` beside the set).
 */
import { writeFileSync } from 'node:fs'
import type { BookInput, PivotTable } from '../../../src/contract.js'
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

export interface CheckedCase {
  id: string
  /** The rows (`DATA`, `kinds` by default). */
  data?: string
  fields: Placed[]
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
}

/** `PT1` of `P` as the answer reads it. */
export function checkedOf(book: BookInput): CheckedAnswer {
  const t = book.sheets.find((s) => s.name === 'P')!.pivotTables!.find((p) => p.name === 'PT1') as PivotTable
  const fields = book.pivotCaches!.find((c) => c.id === t.cache)!.fields
  const names = (axis?: PivotTable['rows']) => (axis ?? []).map((f) => (f === 'values' ? 'Σ' : fields[f].name))
  return {
    rows: names(t.rows),
    columns: names(t.columns),
    values: (t.values ?? []).map((d) => ({ field: fields[d.field].name, function: d.function ?? 'sum' })),
    groups: fields.flatMap((f, i) => (f.group?.range ? [{ field: f.name, base: fields[f.group.base ?? i].name, by: f.group.range.by }] : [])),
  }
}

const CLOCK = clockOptions({ now: '2026-10-31', seed: 0 })

/** The engine's placing of a case's fields, or its refusal. */
export function predictChecked(c: CheckedCase): CheckedAnswer | { refused: string } {
  registerPivot(pivotEngine)
  let book = startBook({ id: c.id, data: c.data ?? 'kinds', ops: [] })
  for (const p of c.fields) {
    const cache = book.pivotCaches![0]
    const t = book.sheets.find((s) => s.name === 'P')!.pivotTables![0]
    const f = cache.fields.findIndex((x) => x.name === p.name)
    const s = suggestPivotField(cache, f)
    const set = p.to === 'columns' ? { columns: [...(t.columns ?? []), f] } : s.axis === 'values' ? { values: [...(t.values ?? []), { field: f, function: s.function }] } : { rows: [...(t.rows ?? []), f] }
    const r = pivotOp(book, { op: 'editPivotTable', sheet: 'P', pivot: 'PT1', set, groupDates: true }, stateValues(createState(book, CLOCK)))
    if (!r.ok) return { refused: r.message }
    book = r.book
  }
  return checkedOf(book)
}

if (process.argv[1]?.endsWith('checked.ts')) {
  const set = process.argv[2]
  const { buildCheckedCases } = (await import(`./${set}.ts`)) as CheckedSet
  const out = Object.fromEntries(buildCheckedCases().map((c) => [c.id, predictChecked(c)]))
  writeFileSync(new URL(`${set}.predicted.json`, import.meta.url), JSON.stringify(out, null, 1) + '\n')
  console.log(`${Object.keys(out).length} cases`)
}
