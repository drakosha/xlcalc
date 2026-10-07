/**
 * What cells show against Excel for Windows (`test/excel/win/display/`, `bench/windows`): every case's shown text and
 * colour, computed by `compute` in the book the harness laid out, as `display-excel.test.ts` does for the Mac. Excel
 * for Windows lays text out in pixels (a digit of Calibri 11 is 7 pixels at 96 dpi) where the engine follows the Mac's
 * points, so cases near a column's edge part; they are listed in `WIN_SHOWN` with the reason.
 */
import { describe, expect, it } from 'vitest'
import type { StoredDisplay } from '../bench/coverage/display.js'
import type { DisplayCase } from '../bench/coverage/display/cases.js'
import { compute } from '../src/index.js'
import { CLOCK, displayBook, sameShown } from './display-cases.js'
import { winFiles } from './excel-win-cases.js'

const SETS = [...winFiles<StoredDisplay>('display/').map((s) => ({ ...s, locale: 'en-US' })), ...winFiles<StoredDisplay>('display/ru-RU/').map((s) => ({ ...s, locale: 'ru-RU' }))]

/** Cases that show otherwise than in Excel for Windows, by `<locale>:<set>:<id>`, and why. */
const WIN_SHOWN: Readonly<Record<string, string>> = {}

for (const set of SETS) {
  describe(`cells shown as in Excel for Windows ${set.locale}: set ${set.name}`, async () => {
    const built = ((await import(`../bench/coverage/display/${set.name}.ts`)) as { buildDisplayCases: () => DisplayCase[] }).buildDisplayCases()

    it('the stored cases are the current ones, all answered', () => {
      expect(set.cases).toEqual(built)
      expect(built.filter((c) => !set.answers[c.id])).toEqual([])
    })

    it('every case shows what Excel for Windows showed, but the cases listed', () => {
      const { book, at } = displayBook(set.cases, set.locale)
      const cells = compute(book, { clock: CLOCK }).sheets[0].cells
      const off = set.cases.flatMap((c) => {
        const out = cells[at.get(c.id)!]
        const engine = { text: out.text!, ...(out.color ? { color: out.color } : {}) }
        const key = `${set.locale}:${set.name}:${c.id}`
        return sameShown(set.answers[c.id], engine, out.value) === !WIN_SHOWN[key] ? [] : [`${key} ${JSON.stringify(c.cell)} width ${c.width ?? 'default'}: Windows ${JSON.stringify(set.answers[c.id])}, engine ${JSON.stringify(engine)}${WIN_SHOWN[key] ? ' (listed, but matches)' : ''}`]
      })
      expect(off).toEqual([])
    })
  })
}

describe('known differences from Excel for Windows: shown text', () => {
  it('every WIN_SHOWN entry names a stored case', () => {
    const keys = new Set(SETS.flatMap((s) => s.cases.map((c) => `${s.locale}:${s.name}:${c.id}`)))
    expect(Object.keys(WIN_SHOWN).filter((k) => !keys.has(k))).toEqual([])
  })
})
