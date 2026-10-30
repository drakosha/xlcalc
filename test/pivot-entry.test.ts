/**
 * An entry into a PivotTable's cell (`pivotEntry`) against Excel's set `edit`: the book before the entry (`<case>-1`)
 * and the entry give Excel's refusal, or no change, or an edit of the definition that, applied, gives the definition
 * Excel saved after it (`<case>`), field by field of what an entry changes.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { CellInput, PivotTable, PivotTablePatch } from '../src/contract.js'
import { pivotEntry } from '../src/pivot/entry.js'
import { readXlsx } from '../src/xlsx.js'

const DIR = new URL('./excel/win/pivot/', import.meta.url).pathname
const { cases, answers } = JSON.parse(readFileSync(`${DIR}edit.json`, 'utf8')) as {
  cases: { id: string; ops: { edit?: string; to?: unknown }[] }[]
  answers: Record<string, { phases: { errors?: string[] }[] }>
}

/** The edit applied as `editPivotTable` sets it: a field replaced, null removed; `fields` by the field's index. */
function applied(t: PivotTable, set: PivotTablePatch): PivotTable {
  const out: Record<string, unknown> = { ...t }
  for (const [k, v] of Object.entries(set)) {
    if (k === 'fields') continue
    if (v === null) delete out[k]
    else out[k] = v
  }
  const fields = t.fields.map((f) => ({ ...f }))
  for (const { field, ...p } of set.fields ?? []) Object.assign(fields[field], p)
  return { ...(out as unknown as PivotTable), fields }
}

/** What an entry changes of a definition. */
const changed = (t: PivotTable) => ({ fields: t.fields.map((f) => ({ name: f.name, items: f.items?.map((i) => [i.item, i.caption]), subtotalCaption: f.subtotalCaption })), rows: t.rows, columns: t.columns, pages: t.pages, values: t.values?.map((v) => v.name), captions: t.captions })

/** The entry as a session's cell input: text typed or set, a number, a formula, a cleared cell. */
function inputOf(to: unknown): CellInput | null {
  if (to === null) return null
  if (typeof to === 'object' && to !== null && 'formula' in to) return { formula: String((to as { formula: string }).formula) }
  if (typeof to === 'object' && to !== null && 'typed' in to) return { value: (to as { typed: string }).typed }
  return { value: to as string | number }
}

describe('an entry into a PivotTable, against Excel', () => {
  it.each(cases.map((c) => c.id))('%s', async (id) => {
    const c = cases.find((x) => x.id === id)!
    const { edit, to } = c.ops[c.ops.length - 1]
    const [sheet, cell] = edit!.split('!')
    const before = (await readXlsx(readFileSync(`${DIR}edit/${id}-1.xlsx`))).book
    const r = pivotEntry(before, sheet, cell, inputOf(to))
    const errors = answers[id].phases[answers[id].phases.length - 1].errors
    if (r === null) {
      expect(errors).toBeUndefined()
      return
    }
    if (errors) {
      expect(r).toMatchObject({ ok: false, code: 'pivot' })
      return
    }
    if (!r.ok) throw new Error(r.message)
    const t = before.sheets.find((s) => s.name === sheet)!.pivotTables![0]
    const after = (await readXlsx(readFileSync(`${DIR}edit/${id}.xlsx`))).book.sheets.find((s) => s.name === sheet)!.pivotTables![0]
    const set = r.op && r.op.op === 'editPivotTable' ? r.op.set : {}
    expect(changed(applied(t, set))).toEqual(changed(after))
  })
})
