/**
 * Conditional formats against Excel 16.113 for Mac, en-US and ru-RU (`test/excel/cond/`, see `test/cond-excel.ts`):
 * what the read cells show after each case, the rules Excel saved, and what the engine shows on the saved book. Known
 * differences are pinned in `COND_KNOWN` with their reasons.
 */
import { describe, expect, it } from 'vitest'
import type { CondSet } from '../bench/coverage/cond/cases.js'
import { checkCondSets, condSets, pinnedCases, type Known } from './cond-excel.js'

/** Excel drops the lengths of an Excel 2010 bar of 5 and 95 percent, which then read as the format's defaults, 10 and 90. */
const LENGTHS = 'Excel saves the Excel 2010 bar of lengths 5 and 95 without them: they read as 10 and 90'

export const COND_KNOWN: Known = {
  'files:kinds': LENGTHS,
  'files:kinds-insert-row': LENGTHS,
  'files:kinds-delete-col': LENGTHS,
  'files:kinds-move': LENGTHS,
  'files:kinds-rename': LENGTHS,
  // a colour between a percentile point and another at an exact third: Excel's channel is one more (its percentile's last bit)
  'files:kinds:S!G8': 'a colour one channel apart at an exact fraction past a percentile point',
  'files:kinds:saved:S!G8': 'the same on the saved book',
  // the order of rules: Excel put a cell value rule whose reference a row deletion lost after the others
  'struct-held:delete-row-2': 'Excel puts a rule whose reference a deletion of rows lost after the others; the engine keeps its place (and Excel kept it deleting a column)',
  'struct-held:delete-rows-2-4': 'Excel puts a rule whose reference a deletion of rows lost after the others; the engine keeps its place',
  // Excel for Mac orders Cyrillic before Latin in ru-RU; the engine follows Windows (Latin first)
  'looks-ru:ru:Лист!D10': 'Excel for Mac ru-RU: "b" > "в"; Windows and the engine: Latin before Cyrillic',
}

const SETS = [...condSets('cond/'), ...condSets('cond/ru-RU/').map((s) => ({ ...s, name: s.name }))]

describe('conditional formats against Excel for Mac', () => {
  it('every set is answered', () => expect(SETS.map((s) => s.name).sort()).toEqual(['files', 'held', 'held-2', 'looks', 'looks-ru', 'struct-held']))
  it('every pinned difference names a case', () => expect(pinnedCases(SETS, COND_KNOWN)).toEqual([]))
})

checkCondSets(SETS, COND_KNOWN, async (name) => (await import(`../bench/coverage/cond/${name}.ts`)) as CondSet)
