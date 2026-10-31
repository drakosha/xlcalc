/**
 * GETPIVOTDATA against Excel for Windows (`test/excel/win/pivot/gpd-1.json`, the formulas of
 * `bench/coverage/pivot/gpd-1.ts` over the books of the set basic): each case's watched cells as Excel computed them,
 * and over a PivotTable of another workbook the reads of the harness the engine can make, a computation without that
 * workbook (#REF!) and the file opened again (the value it stores). The engine's prediction was committed before the
 * answers were read (`gpd-1.predicted.json`); the differences it had are fitted (`docs/pivot-tables.md`).
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { compute, type BookInput, type CellValue } from '../src/index.js'
// the set of Excel ru-RU computes in ru-RU
import './locales.js'
import { buildPivotCases } from '../bench/coverage/pivot/gpd-1.js'
import { buildPivotCases as heldCases } from '../bench/coverage/pivot/gpd-held.js'
import { buildPivotCases as held2Cases } from '../bench/coverage/pivot/gpd-held-2.js'
import { buildPivotCases as held3Cases } from '../bench/coverage/pivot/gpd-held-3.js'
import { predictCase } from '../bench/coverage/pivot/gpd-predict.js'
import type { PivotAnswer, PivotCase } from '../bench/coverage/pivot/cases.js'

const answers = (JSON.parse(readFileSync(new URL('./excel/win/pivot/gpd-1.json', import.meta.url), 'utf8')) as { answers: Record<string, PivotAnswer> }).answers
const near = (a: CellValue, b: CellValue) => (typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b)) : JSON.stringify(a) === JSON.stringify(b))

/** Formulas Excel did not take as they were typed (an odd number of pairs, `0x800A03EC`), whose cells are left empty. */
const REFUSED = new Set(['subtotals-functions G!A2'])

/** Each formula of a set over a PivotTable of its book against Excel's value: the cells that differ, and the count. */
async function compared(set: string, cases: PivotCase[], locale?: string): Promise<[string[], number]> {
  const all = (JSON.parse(readFileSync(new URL(`./excel/win/pivot/${locale ? `${locale}/` : ''}${set}.json`, import.meta.url), 'utf8')) as { answers: Record<string, PivotAnswer> }).answers
  const bad: string[] = []
  let n = 0
  for (const c of cases) {
    if (c.ops.some((o) => 'external' in o)) continue
    if (!all[c.id]) continue
    const got = await predictCase(c, set, locale)
    for (const [cell, v] of Object.entries(all[c.id].phases[0].cells ?? {})) {
      if (REFUSED.has(`${c.id} ${cell}`)) continue
      n++
      if (!near(got[cell], v)) bad.push(`${c.id} ${cell}: ${JSON.stringify(got[cell])} vs ${JSON.stringify(v)}`)
    }
  }
  return [bad, n]
}

describe('GETPIVOTDATA against Excel for Windows: the set gpd-1', () => {
  it('every formula over a PivotTable of the book', async () => {
    expect(await compared('gpd-1', buildPivotCases())).toEqual([[], 101])
  }, 60_000)

  it('the set in Excel ru-RU: a value field by its ru-RU name, an English one #REF!', async () => {
    expect(await compared('gpd-1', buildPivotCases(), 'ru-RU')).toEqual([[], 101])
  }, 60_000)

  it('the held-out set gpd-held, fitted after its run (51 of 55 predicted: an inner field without the one between, subtotals of several functions, a field by its own name alone)', async () => {
    expect(await compared('gpd-held', heldCases())).toEqual([[], 54])
  }, 60_000)

  it('the held-out set gpd-held-3 of renamed items, fitted after its run (30 of 31 predicted: a renamed boolean is not named by its text)', async () => {
    expect(await compared('gpd-held-3', held3Cases())).toEqual([[], 31])
  }, 60_000)

  it('the held-out set gpd-held-2 over the PivotTables its cases built, fitted after its run (57 of 61 predicted: the cell above the top left of a PivotTable without a Filters area)', async () => {
    expect(await compared('gpd-held-2', held2Cases())).toEqual([[], 61])
  }, 60_000)

  it('over a PivotTable of another workbook: #REF! computed without it, the value the file stores when opened again', () => {
    const c = buildPivotCases().find((x) => x.id === 'other-book')!
    const o = c.ops.find((x) => 'external' in x) as { cells: Record<string, string> }
    const reads = answers['other-book'].phases[0].external!.reads
    const cells = Object.keys(o.cells).map((k) => k.split('!')[1])
    const book: BookInput = { sheets: [{ name: 'G', cells: Object.fromEntries(cells.map((a) => [a, { formula: o.cells[`G!${a}`].slice(1).replace('[Src.xlsx]', '[1]') }])) }], externalBooks: [{ target: 'Src.xlsx', sheets: [{ name: 'P', cells: {} }] }] }
    const value = (r: ReturnType<typeof compute>) => Object.fromEntries(cells.map((a) => [`G!${a}`, r.sheets[0].cells[a].value]))
    expect(value(compute(book))).toEqual(reads['closed, computed'])
    expect(reads['reopened, computed']).toEqual(reads['closed, computed'])
    const stored = { sheets: [{ name: 'G', cells: Object.fromEntries(cells.map((a) => [a, { value: reads.reopened[`G!${a}`] }])) }] }
    const held = value(compute(book, { cached: stored }))
    expect(Object.entries(reads.reopened).filter(([k, v]) => !near(held[k], v))).toEqual([])
  })
})
