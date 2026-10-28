/**
 * Conditional formats against Excel 16.113 for Mac, en-US and ru-RU (`test/excel/cond/`, see `test/cond-excel.ts`):
 * what the read cells show after each case, the rules Excel saved, and what the engine shows on the saved book. Known
 * differences are pinned in `COND_KNOWN` with their reasons.
 */
import { describe, expect, it } from 'vitest'
import type { CondSet } from '../bench/coverage/cond/cases.js'
import { checkCondSets, condSets, type Known } from './cond-excel.js'
import { pinnedCases } from './stored-cases.js'
import './locales.js'

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
  // Excel for Mac orders Cyrillic before Latin in ru-RU; the engine follows Windows (Latin first)
  'looks-ru:ru:Лист!D10': 'Excel for Mac ru-RU: "b" > "в"; Windows and the engine: Latin before Cyrillic',
  // Excel for Mac's short month names in ru-RU; the engine follows Windows (three letters)
  'formats-ru-held:formats:Лист!A5': 'Excel for Mac ru-RU writes the short month name март; Windows and the engine мар',
  'formats-ru-held:formats:Лист!A6': 'Excel for Mac ru-RU writes the short month name март; Windows and the engine мар',
  // both Excels save two rules alike but for a format given by number and by its code as one rule over both cells
  'formats-ru-held:formats': "Excel saves a rule of format 14 and one of m/d/yyyy over the cell below as one rule over both, as it reads a rule's format by its code",
  // what the model has no fill for, each reported by the reader: a gradient (Excel's display format gives it black), a
  // dxf's none, which takes the cell's own fill away, and a dxf's solid fill without its colour, white on a cell without one
  ...Object.fromEntries(['A7', 'B7', 'C7'].map((a) => [`fills-held:fills:S!${a}`, 'a gradient fill is not in the model'])),
  ...Object.fromEntries(['fills-held:fills:S!C1', 'fills-held:fills:S!C2', 'fills-held:fills:S!C3', 'fills-held-2:fills:S!C4'].map((k) => [k, "a conditional format's fill taken away (patternType none) is not in the model"])),
  ...Object.fromEntries(['B4', 'B6'].map((a) => [`fills-held:fills:S!${a}`, "a conditional format's solid fill without its colour, Excel's white on a cell without a fill, is not in the model"])),
}

/** The pins of what only Excel for Mac does: its order of text and month names in ru-RU. */
export const macOnly = (key: string) => key.startsWith('looks-ru:') || key.startsWith('formats-ru-held:formats:')

const SETS = [...condSets('cond/'), ...condSets('cond/ru-RU/').map((s) => ({ ...s, name: s.name }))]

describe('conditional formats against Excel for Mac', () => {
  it('every set is answered', () => expect(SETS.map((s) => s.name).sort()).toEqual(['both', 'edge-held', 'files', 'fills-held', 'fills-held-2', 'formats-ru-held', 'held', 'held-2', 'looks', 'looks-ru', 'order', 'sqref-held', 'struct-held']))
  it('every pinned difference names a case', () => expect(pinnedCases(SETS, COND_KNOWN, true)).toEqual([]))
})

checkCondSets(SETS, COND_KNOWN, async (name) => (await import(`../bench/coverage/cond/${name}.ts`)) as CondSet)
