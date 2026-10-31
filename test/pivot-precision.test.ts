import { readFileSync } from 'node:fs'
import { expect, it } from 'vitest'
import { replays, shown } from '../bench/coverage/pivot/phases.js'
import { buildPivotCases as arithmetic } from '../bench/coverage/pivot/order-probe.js'
import { buildPivotCases as numeric } from '../bench/coverage/pivot/numeric-order.js'
import { buildPivotCases as showCases } from '../bench/coverage/pivot/show.js'
import { buildPivotCases as calcCases } from '../bench/coverage/pivot/calc.js'
import { buildPivotCases as calcHeld } from '../bench/coverage/pivot/calc-held.js'
import { buildPivotCases as filtered } from '../bench/coverage/pivot/filter.js'
import { buildPivotCases as partitionHeld } from '../bench/coverage/pivot/filter-order-held.js'
import { buildPivotCases as filteredHeld } from '../bench/coverage/pivot/filter-held.js'
import { buildPivotCases as showItems } from '../bench/coverage/pivot/show-items.js'
import { buildPivotCases as numericHeld } from '../bench/coverage/pivot/numeric-held.js'
import { buildPivotCases as numericHeldExact } from '../bench/coverage/pivot/numeric-held-exact.js'
import { buildPivotCases as basicCases } from '../bench/coverage/pivot/basic.js'
import type { PivotAnswer } from '../bench/coverage/pivot/cases.js'

const root = new URL('./excel/win/pivot/', import.meta.url)
for (const [set, cases] of [
  ['order-probe', arithmetic().filter(c => c.id.startsWith('arithmetic-'))],
  ['numeric-order', numeric()],
  ['show', showCases().filter(c => ['normal', 'index', 'run-percent-region'].includes(c.id))],
  ['calc', calcCases()],
  ['calc-held', calcHeld()],
  ['basic', basicCases()],
  ['filter', filtered()],
  ['filter-held', filteredHeld()],
  ['filter-order-held', partitionHeld()],
  ['show-items', showItems()],
  ['numeric-held', numericHeld()],
  ['numeric-held-exact', numericHeldExact()],
] as const) {
  const answers = JSON.parse(readFileSync(new URL(`${set}.json`, root), 'utf8')).answers as Record<string, PivotAnswer>
  for (const c of cases)
    it(`keeps every binary64 digit of ${set}/${c.id}`, async () => {
      const phases = answers[c.id].phases
      for (const replay of await replays(set, c, phases.length)) {
        expect(replay.result.ok).toBe(true)
        if (!replay.result.ok) continue
        const want = phases[Number(replay.at.split('/')[1])].pivots.PT1
        if ('error' in want) throw new Error(want.error)
        expect(shown(replay.result.book, 'P', want.area), replay.at).toEqual(want.values)
      }
    })
}
