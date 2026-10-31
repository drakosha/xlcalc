import { readFileSync } from 'node:fs'
import { expect, it } from 'vitest'
import { drive } from '../bench/coverage/pivot/drive.js'
import { buildPivotCases } from '../bench/coverage/pivot/order-probe.js'
import { buildPivotCases as heldCases } from '../bench/coverage/pivot/order-check.js'
import { buildPivotCases as sortCases } from '../bench/coverage/pivot/sort.js'
import { rounded, shown } from '../bench/coverage/pivot/phases.js'
import type { PivotAnswer } from '../bench/coverage/pivot/cases.js'

const root = new URL('./excel/win/pivot/', import.meta.url)
for (const [set, cases] of [['order-check', heldCases()], ['order-probe', buildPivotCases().filter(c => c.id.startsWith('lists-'))], ['sort', sortCases().filter(c => c.id === 'months' || c.id === 'months-down')]] as const) {
  const answers = JSON.parse(readFileSync(new URL(`${set}.json`, root), 'utf8')).answers as Record<string, PivotAnswer>
  for (const c of cases)
    it(`creates and sorts ${set}/${c.id} from source rows`, () => {
      const got = drive(c)
      for (const [k, phase] of got.entries()) {
        expect(phase.errors).toEqual([])
        const want = answers[c.id].phases[k].pivots.PT1
        if ('error' in want) throw new Error(want.error)
        expect(rounded(shown(phase.book, 'P', want.area)), `phase ${k}`).toEqual(rounded(want.values))
      }
    })
}
