/**
 * Conditional formats against Excel 16.0.20430 for Windows, en-US and ru-RU (`test/excel/win/cond/`, see
 * `test/cond-excel.ts`): what the read cells show after each case, the rules Excel saved, and what the engine shows on
 * the saved book, the rules Excel made itself through COM (the set made) among them. Excel for Windows lays text out in
 * pixels, so `#` across a column may count otherwise (`PIXEL_HASHES`); other known differences are pinned in `WIN_KNOWN`.
 */
import { describe, expect, it } from 'vitest'
import type { CondSet } from '../bench/coverage/cond/cases.js'
import { checkCondSets, condSets, PIXEL_HASHES, pinnedCases, type Known } from './cond-excel.js'
import { COND_KNOWN, macOnly } from './cond-excel.test.js'
import './locales.js'

export const WIN_KNOWN: Known = {
  // the same as on the Mac: the bar lengths Excel drops, the order after a row deletion, a colour at an exact fraction
  ...Object.fromEntries(Object.entries(COND_KNOWN).filter(([k]) => !macOnly(k))),
  // a colour between a percentile point and the highest at 0.4 of the way: Excel's channel is one more
  'made:made:saved:S!L10': 'a colour one channel apart at an exact fraction past a percentile point',
}

const SETS = [...condSets('win/cond/'), ...condSets('win/cond/ru-RU/')]

describe('conditional formats against Excel for Windows', () => {
  it('every set is answered', () => expect(SETS.map((s) => s.name).sort()).toEqual(['both', 'edge-held', 'files', 'fills-held', 'formats-ru-held', 'held', 'held-2', 'looks', 'looks-ru', 'made', 'sqref-held', 'struct-held']))
  it('every pinned difference names a case', () => expect(pinnedCases(SETS, WIN_KNOWN)).toEqual([]))
})

checkCondSets(SETS, WIN_KNOWN, async (name) => (await import(`../bench/coverage/cond/${name}.ts`)) as CondSet, PIXEL_HASHES)
