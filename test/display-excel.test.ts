/**
 * What cells show against Excel (`test/excel/display/`, Excel 16.113 for Mac in en-US and in ru-RU): every case's
 * shown text (Range.Text) and its section's colour, computed by `compute` in a book laid out as the harness laid
 * it out; held-out cases also keep the answer predicted before Excel ran.
 */
import { describe, expect, it } from 'vitest'
import { compute, displayValue } from '../src/index.js'
import { CLOCK, DISPLAY_SETS, displayBook, sameShown } from './display-cases.js'

/** Held-out cases whose prediction Excel did not confirm, by `<locale>:<set>:<id>` (the engine was fixed after the run). */
const MISSED: readonly string[] = []

for (const set of DISPLAY_SETS) {
  describe(`cells shown as in Excel ${set.locale}: set ${set.name}`, async () => {
    const built = ((await import(`../bench/coverage/display/${set.name}.ts`)) as { buildDisplayCases: () => typeof set.cases }).buildDisplayCases()

    it('the stored cases are the current ones, all answered', () => {
      expect(set.cases).toEqual(built)
      expect(built.filter((c) => !set.answers[c.id])).toEqual([])
    })

    it('every case shows what Excel showed', () => {
      const { book, at } = displayBook(set.cases, set.locale)
      const cells = compute(book, { clock: CLOCK }).sheets[0].cells
      const off = set.cases.flatMap((c) => {
        const out = cells[at.get(c.id)!]
        const engine = { text: out.text!, ...(out.color ? { color: out.color } : {}) }
        return sameShown(set.answers[c.id], engine, out.value) ? [] : [`${c.id} ${JSON.stringify(c.cell)} width ${c.width ?? 'default'}: Excel ${JSON.stringify(set.answers[c.id])}, engine ${JSON.stringify(engine)}`]
      })
      expect(off).toEqual([])
    })

    it('displayValue shows each value as its cell', () => {
      const off = set.cases.flatMap((c) => {
        if (c.cell.formula !== undefined || (c.span ?? 1) > 1) return []
        const value = c.cell.value === undefined || (typeof c.cell.value === 'object' && c.cell.value !== null) ? null : c.cell.value
        const shown = displayValue(value, { format: c.cell.format, locale: set.locale, width: c.width })
        return sameShown(set.answers[c.id], shown, value) ? [] : [`${c.id}: Excel ${JSON.stringify(set.answers[c.id])}, displayValue ${JSON.stringify(shown)}`]
      })
      expect(off).toEqual([])
    })

    const held = set.cases.filter((c) => c.predicted)
    if (held.length)
      it('held-out cases: the prediction committed before Excel ran is what Excel showed, but for the misses listed', () => {
        const missed = held.filter((c) => JSON.stringify(c.predicted) !== JSON.stringify(set.answers[c.id])).map((c) => `${set.locale}:${set.name}:${c.id}`)
        expect(missed).toEqual(MISSED.filter((k) => k.startsWith(`${set.locale}:${set.name}:`)))
      })
  })
}
