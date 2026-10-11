/**
 * Acceptance against Excel for Windows on the formula cases (`test/excel/win/`, `bench/windows/`): the measurement,
 * the probes and the case sets, each in the locale it was computed in. Known differences are pinned in `WIN_KNOWN`
 * (`excel-win-cases.ts`); formulas Excel for Windows did not take are listed in `WIN_REJECTED` with the reason.
 */
import { describe, expect, it } from 'vitest'
import { describeWinOff, runWinSet, WIN_KNOWN, WIN_REASONS, WIN_SETS } from './excel-win-cases.js'

/** Cases whose formula Excel for Windows did not take when it opened the book, by `<locale>:<source>:<case id>`, and why. */
const WIN_REJECTED: Readonly<Record<string, string>> = {
  'en-US:name-case-held:bool#5': 'ISLOGICAL({falſe}): Excel for Windows reads falſe as a name, which an array constant does not hold, and does not open the book; the engine refuses the formula (#ERROR!)',
}

const keys = new Set<string>()

for (const set of WIN_SETS) {
  const { current, missing, outcomes } = await runWinSet(set)
  for (const o of outcomes) keys.add(o.key)

  describe(`acceptance against Excel for Windows ${set.tag}: ${set.name}`, () => {
    it('the stored answers are for the current cases, every case answered or not taken', () => {
      expect(current).toBe(true)
      expect(missing).toEqual([])
    })

    it('the formulas Excel for Windows did not take are the ones listed', () => {
      const listed = Object.keys(WIN_REJECTED).filter((k) => k.startsWith(`${set.tag}:${set.name}:`))
      expect((set.rejected ?? []).map((id) => `${set.tag}:${set.name}:${id}`)).toEqual(listed)
    })

    for (const group of [...new Set(outcomes.map((o) => o.group))])
      it(`${group}: all cases match Excel for Windows`, () => {
        expect(outcomes.filter((o) => o.group === group && !o.ok).map(describeWinOff)).toEqual([])
      })
  })
}

describe('known differences from Excel for Windows', () => {
  it('every WIN_KNOWN entry names a stored case', () => {
    expect(Object.keys(WIN_KNOWN).filter((k) => !keys.has(k))).toEqual([])
  })

  it('every reason of WIN_REASONS has a case', () => {
    const used = new Set(Object.values(WIN_KNOWN).map((k) => k.reason))
    expect(Object.keys(WIN_REASONS).filter((r) => !used.has(r as keyof typeof WIN_REASONS))).toEqual([])
  })
})
