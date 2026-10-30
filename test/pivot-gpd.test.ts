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
import { buildPivotCases } from '../bench/coverage/pivot/gpd-1.js'
import { predictCase } from '../bench/coverage/pivot/gpd-predict.js'
import type { PivotAnswer } from '../bench/coverage/pivot/cases.js'

const answers = (JSON.parse(readFileSync(new URL('./excel/win/pivot/gpd-1.json', import.meta.url), 'utf8')) as { answers: Record<string, PivotAnswer> }).answers
const near = (a: CellValue, b: CellValue) => (typeof a === 'number' && typeof b === 'number' ? Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b)) : JSON.stringify(a) === JSON.stringify(b))

describe('GETPIVOTDATA against Excel for Windows: the set gpd-1', () => {
  it('every formula over a PivotTable of the book', async () => {
    const bad: string[] = []
    let n = 0
    for (const c of buildPivotCases()) {
      if (c.ops.some((o) => 'external' in o)) continue
      const got = await predictCase(c)
      for (const [cell, v] of Object.entries(answers[c.id].phases[0].cells ?? {})) {
        n++
        if (!near(got[cell], v)) bad.push(`${c.id} ${cell}: ${JSON.stringify(got[cell])} vs ${JSON.stringify(v)}`)
      }
    }
    expect(bad).toEqual([])
    expect(n).toBe(101)
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
