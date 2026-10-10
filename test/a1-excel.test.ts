/**
 * Sheet names in references against Excel for Mac (`test/excel/tables/[ru-RU/]sheet-quotes*.json`, the sets
 * `bench/coverage/tables/sheet-quotes*.ts`): `quoteSheetIfNeeded` writes each name as Excel wrote it in `Range.Formula`,
 * in en-US and ru-RU alike, and with the locale's booleans as `Range.FormulaLocal` in ru-RU. The held-out predictions
 * committed before Excel ran keep their misses listed.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredTables } from '../bench/coverage/tables.js'
import { cellOf } from '../bench/coverage/tables/sheet-quotes-held.js'
import { LOCALES, quoteSheetIfNeeded } from '../src/index.js'

const SETS = ['sheet-quotes-held', 'sheet-quotes-r1c1-held']
const DIRS: [string, string][] = [['en-US', ''], ['ru-RU', 'ru-RU/']]

/** Names whose prediction Excel did not confirm (the rule was fixed after each run), by `<set>:<name>`. */
const MISSED: readonly string[] = [
  'sheet-quotes-held:R0', 'sheet-quotes-held:R1048577', 'sheet-quotes-held:R2D2', 'sheet-quotes-held:C3PO', 'sheet-quotes-held:RC1C',
  'sheet-quotes-r1c1-held:C1.x', 'sheet-quotes-r1c1-held:R1C0', 'sheet-quotes-r1c1-held:R1C16385', 'sheet-quotes-r1c1-held:RCC1', 'sheet-quotes-r1c1-held:R1.', 'sheet-quotes-r1c1-held:xR1',
]

for (const set of SETS) {
  describe(`sheet names in references as Excel writes them: ${set}`, async () => {
    const mod = (await import(`../bench/coverage/tables/${set}.ts`)) as { NAMES: string[]; PREDICTED: Record<string, string> }
    const ref = (name: string) => `F!${cellOf(name, mod.NAMES)}`

    for (const [tag, dir] of DIRS) {
      const file = new URL(`./excel/tables/${dir}${set}.json`, import.meta.url)
      if (!existsSync(file)) continue
      const stored = JSON.parse(readFileSync(file, 'utf8')) as StoredTables
      const answer = stored.answers.names

      it(`${tag}: Range.Formula`, () => {
        expect(mod.NAMES.filter((n) => answer.formulas[ref(n)] !== `=${quoteSheetIfNeeded(n)}!A1`)).toEqual([])
      })

      it(`${tag}: Range.FormulaLocal, the locale's booleans quoted`, () => {
        const locale = LOCALES.get(tag)!
        expect(mod.NAMES.filter((n) => answer.local![ref(n)] !== `=${quoteSheetIfNeeded(n, locale)}!A1`)).toEqual([])
      })
    }

    it('the prediction committed before Excel ran is what Excel wrote, but for the misses listed', () => {
      const answer = (JSON.parse(readFileSync(new URL(`./excel/tables/${set}.json`, import.meta.url), 'utf8')) as StoredTables).answers.names
      const missed = mod.NAMES.filter((n) => answer.formulas[ref(n)] !== mod.PREDICTED[n]).map((n) => `${set}:${n}`)
      expect(missed).toEqual(MISSED.filter((k) => k.startsWith(`${set}:`)))
    })
  })
}
