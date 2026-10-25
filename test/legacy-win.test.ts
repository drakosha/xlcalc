/**
 * Formulas of the old kind in Excel for Windows (`test/excel/win/legacy/`, `bench/windows`): each case stood at the row
 * it had in the Mac's run, and Excel for Windows shows the formula as the Mac does (`Formula2`, `@` where it intersects)
 * and computes the same value (numbers as the measurement compares them, a volatile call's by type). `legacy.test.ts`
 * checks the engine against the Mac's answers, so where Windows gives the same, the engine matches it too; the cases
 * where Windows parts from the Mac are listed in `WIN_LEGACY` with the reason.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { compare } from '../bench/coverage/compare.js'
import type { Val } from '../bench/coverage/formulajs.js'
import type { Probe } from '../bench/coverage/probes.js'
import { show } from './excel-cases.js'
import { storedSets } from './stored-sets.js'
import { pinnedCases } from './stored-cases.js'

interface Stored {
  cases: Probe[]
  values: Record<string, Val>
  formulas: Record<string, string>
  rows: Record<string, number>
}

const mac = (name: string) => JSON.parse(readFileSync(new URL(`./excel/legacy/${name}.json`, import.meta.url), 'utf8')) as Stored
const SETS = storedSets<Stored>('win/legacy')

/** Volatile calls: their values are compared by type only. */
const VOLATILE = /\b(RAND|RANDBETWEEN|RANDARRAY|NOW|TODAY)\(/

/** `<set>:<case id>` → why Excel for Windows parts from Excel for Mac. */
const WIN_LEGACY: Readonly<Record<string, string>> = {}

describe.skipIf(!existsSync(new URL('./excel/win/legacy/', import.meta.url)) || !readdirSync(new URL('./excel/win/legacy/', import.meta.url)).length)('legacy formulas in Excel for Windows', () => {
  for (const set of SETS)
    it(`${set.name}: as on the Mac, at the Mac's rows, but the cases listed`, () => {
      const m = mac(set.name)
      expect(set.cases).toEqual(m.cases)
      expect(set.rows).toEqual(m.rows)
      const off = set.cases.flatMap((c) => {
        const value = VOLATILE.test(c.formula) ? m.values[c.id].t === set.values[c.id].t : compare(m.values[c.id], set.values[c.id]).ok
        const same = value && set.formulas[c.id] === m.formulas[c.id]
        return same === !WIN_LEGACY[`${set.name}:${c.id}`] ? [] : [`${c.id}: Mac ${show(m.values[c.id])} ${m.formulas[c.id]}, Windows ${show(set.values[c.id])} ${set.formulas[c.id]}`]
      })
      expect(off).toEqual([])
    })

  it('every WIN_LEGACY entry names a stored case', () => expect(pinnedCases(SETS, WIN_LEGACY)).toEqual([]))
})
