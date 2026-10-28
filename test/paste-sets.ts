/**
 * Paste Special against stored answers of Excel (`test/excel/paste/` of the Mac, `test/excel/win/paste/` of Windows):
 * each case's book after `editStructure`'s `copyRange` with `paste` and `transpose`, written and read back, shows what
 * the file Excel saved shows (`bench/coverage/paste/shown.ts`). A paste Excel refused (it leaves the book as it was,
 * with or without an error) is one `editStructure` refuses. Kinds the engine does not have are left out; a known
 * difference must still differ.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { writeXlsx } from '../src/xlsx.js'
import type { PasteAnswer, PasteCase, PasteCases } from '../bench/coverage/paste/cases.js'
import { engineHas, engineOutcome, shownFile, type Shown } from '../bench/coverage/paste/shown.js'

/** The paste sets with answers of Excel. */
export const PASTE_SETS = ['probe', 'held', 'merges', 'fit', 'kinds', 'ops']

/** Where the engine pastes otherwise than Excel for Mac, by set and case. */
export const PASTE_KNOWN: Readonly<Record<string, string>> = {
  'held/table-formats': "Excel pastes the look a table's style gives its cells as their formats; the engine does not draw table styles",
  'kinds/allExceptBorders-column': 'a whole column that cuts a merged area (A5:B5): Excel leaves out the merged cell, the engine pastes it',
  ...Object.fromEntries(
    ['number-add-all', 'number-add-values', 'formula-add-formulas', 'formula-add-all', 'formula-subtract-values'].map((id) => [
      `ops/${id}`,
      'Excel refuses an operation onto a formula it holds as a dynamic array formula, as a file xlcalc writes holds every formula; the engine combines it as Excel combines a formula of the ordinary kind (the set ops-held)',
    ]),
  ),
}

/**
 * Checks the sets `sets` stored in `test/excel/<dir>/` that are there; `known` names by `<set>/<case>` where the engine
 * differs, and why; `as` puts both books in the form a platform saves (its units), the same for the engine and Excel.
 */
export function pasteSets(excel: string, dir: string, sets: readonly string[], known: Readonly<Record<string, string>>, as: (x: Shown) => Shown = (x) => x): void {
  const stored = sets.filter((set) => existsSync(new URL(`./excel/${dir}/${set}.json`, import.meta.url)))
  describe.each(stored)(`paste set %s against ${excel}`, (set) => {
    const s = JSON.parse(readFileSync(new URL(`./excel/${dir}/${set}.json`, import.meta.url), 'utf8')) as PasteCases & { answers: Record<string, PasteAnswer> }
    const cases = s.cases.filter((c) => engineHas(c))
    it.each(cases.map((c) => [c.id, c] as [string, PasteCase]))('%s', async (_, c) => {
      const book = s.books[c.book]
      const saved = as(await shownFile(new Uint8Array(readFileSync(new URL(`./excel/${dir}/${set}/${c.id}.xlsx`, import.meta.url)))))
      const out = await engineOutcome(book, c)
      const r = 'refused' in out ? out : { shown: as(out.shown) }
      const why = known[`${set}/${c.id}`]
      if (why) expect('refused' in r ? null : r.shown, why).not.toEqual(saved)
      else if ('refused' in r) expect(saved, `${c.id} refused: ${r.refused}`).toEqual(as(await shownFile((await writeXlsx({ book })).data)))
      else expect(r.shown).toEqual(saved)
    })
  })
  it(`every known difference from ${excel} names a stored case`, () => {
    const ids = new Set(stored.flatMap((set) => (JSON.parse(readFileSync(new URL(`./excel/${dir}/${set}.json`, import.meta.url), 'utf8')) as PasteCases).cases.map((c) => `${set}/${c.id}`)))
    expect(Object.keys(known).filter((k) => !ids.has(k))).toEqual([])
  })
}
