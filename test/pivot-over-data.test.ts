/**
 * A new PivotTable over cells that are not empty, against Excel for Windows (the set `over-data`, made through COM, and
 * `checked-over`, its fields checked in the field list so that it grows over cells below): Excel 365 asks nothing and
 * shows `#SPILL!` in the PivotTable's top-left cell, the cells in the way kept, over a value, a formula, merged cells
 * and a table; a number format alone is no obstacle, a value in the top-left cell is taken over. `addPivotTable` with
 * `xlcalc/pivot` does the same; without it, it refuses.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { buildPivotCases } from '../bench/coverage/pivot/over-data.js'
import { buildCheckedCases } from '../bench/coverage/pivot/checked-over.js'
import { checkedOf, predictBook } from '../bench/coverage/pivot/checked.js'
import { drive } from '../bench/coverage/pivot/drive.js'
import { rounded, shown } from '../bench/coverage/pivot/phases.js'
import { editStructure, registerPivot, type BookInput } from '../src/index.js'
import { pivotEngine } from '../src/pivot/index.js'
import { setPivotRefresh, pivotRefresh } from '../src/pivot/hook.js'

registerPivot(pivotEngine)

describe('over-data against Excel for Windows', () => {
  const answers = JSON.parse(readFileSync(new URL('./excel/win/pivot/over-data.json', import.meta.url), 'utf8')).answers
  for (const c of buildPivotCases())
    it(c.id, () => {
      for (const [k, p] of drive(c).entries()) {
        expect(p.errors).toEqual([])
        for (const [name, want] of Object.entries(answers[c.id].phases[k].pivots as Record<string, { area: string; values: unknown[][] }>)) {
          expect(p.book.sheets[1].pivotTables!.find((t) => t.name === name)!.location.range, `${k} ${name}`).toBe(want.area)
          expect(rounded(shown(p.book, 'P', want.area)), `${k} ${name}`).toEqual(rounded(want.values as never))
        }
      }
    })
})

describe('checked-over against Excel for Windows', () => {
  const answers = JSON.parse(readFileSync(new URL('./excel/win/pivot-fields/checked-over.json', import.meta.url), 'utf8')).answers
  for (const c of buildCheckedCases())
    it(c.id, () => {
      const book = predictBook(c)
      if (typeof book === 'string') throw new Error(book)
      expect(checkedOf(book, c.watch)).toEqual(answers[c.id])
    })
})

describe('addPivotTable over cells that are not empty', () => {
  const book = (): BookInput => ({
    sheets: [
      { name: 'D', cells: { A1: { value: 'R' }, B1: { value: 'V' }, A2: { value: 'a' }, B2: { value: 1 } } },
      { name: 'P', cells: { A3: { value: 'top' }, B8: { value: 'x' } } },
    ],
  })
  const add = { op: 'addPivotTable', sheet: 'P', cell: 'A3', source: { type: 'range', sheet: 'D', range: 'A1:B2' } } as const

  it('spills with xlcalc/pivot, the top-left cell taken over; lays out once the place is free', () => {
    const r = editStructure(book(), add)
    if (!r.ok) throw new Error(r.message)
    expect(r.book.sheets[1].pivotTables![0].location.range).toBe('A3')
    expect(r.book.sheets[1].cells.A3).toEqual({ formula: '#SPILL!' })
    expect(r.book.sheets[1].cells.B8).toEqual({ value: 'x' })
    // the value in the way taken away
    const cleared = { ...r.book, sheets: [r.book.sheets[0], { ...r.book.sheets[1], cells: { A3: r.book.sheets[1].cells.A3 } }] }
    const laid = editStructure(cleared, { op: 'refreshPivots' })
    if (!laid.ok) throw new Error(laid.message)
    expect(laid.book.sheets[1].pivotTables![0].location.range).toBe('A3:C20')
    expect(laid.book.sheets[1].cells.A3?.formula).toBeUndefined()
  })

  it('without xlcalc/pivot refuses, as it cannot lay the PivotTable out', () => {
    const refresh = pivotRefresh()
    setPivotRefresh(null)
    try {
      expect(editStructure(book(), add)).toMatchObject({ ok: false, code: 'pivot' })
    } finally {
      setPivotRefresh(refresh)
    }
  })
})
