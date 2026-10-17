/**
 * `valueComparator` against Excel's Sort command, as Excel for Mac and Windows answered it (`bench/coverage/sort.ts`,
 * `bench/windows/sort.ts`): each case's column, sorted stably by its cells' values, comes in the order Excel left it.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { createSession, valueComparator } from '../src/index.js'
import { rangeOf, sortBook, type SortAnswers, type SortSet } from '../bench/coverage/sort/cases.js'

const DIRS = { 'Excel for Mac': 'excel/sort', 'Excel for Windows': 'excel/win/sort' }

for (const [platform, dir] of Object.entries(DIRS)) {
  const url = new URL(`./${dir}/`, import.meta.url)
  describe(`the Sort command, ${platform}`, () => {
    for (const file of readdirSync(url).filter((f) => f.endsWith('.json'))) {
      const name = file.slice(0, -5)
      it(name, async () => {
        const { set } = (await import(`../bench/coverage/sort/${name}.ts`)) as { set: SortSet }
        const { answers } = JSON.parse(readFileSync(new URL(file, url), 'utf8')) as SortAnswers
        const s = createSession(sortBook(set))
        set.cases.forEach((c, i) => {
          const col = rangeOf(set, i).replace(/\d.*/, '')
          const cells = c.cells.map((text, r) => ({ text, value: s.getCell('S', `${col}${r + 1}`)?.value ?? null }))
          const cmp = valueComparator(c.descending ? { descending: true } : {})
          expect([c.id, cells.sort((a, b) => cmp(a.value, b.value)).map((x) => x.text)]).toEqual([c.id, answers[c.id]])
        })
      })
    }
  })
}
