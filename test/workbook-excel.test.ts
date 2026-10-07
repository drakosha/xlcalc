/**
 * Book-level features against Excel en-US (`test/excel/workbook/<set>.json`, from `run.ts workbook <set>`): each
 * stored case is its generator's current case, and at each read the cells of its rectangles hold what Excel showed.
 * A held-out set also keeps the answers predicted before Excel ran, and the engine gives those too.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredWorkbook } from '../bench/coverage/workbook.js'
import type { WorkbookCase } from '../bench/coverage/workbook/cases.js'
import { differences, engineReads } from './workbook-cases.js'

const dir = new URL('./excel/workbook/', import.meta.url)
const SETS = existsSync(dir)
  ? readdirSync(dir)
      .filter((f) => f.endsWith('.json'))
      .map((f) => ({ name: f.slice(0, -5), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredWorkbook) }))
  : []

/** `<set>:<case>` → why the engine's answer differs from Excel's. */
const KNOWN: Readonly<Record<string, string>> = {}

for (const set of SETS) {
  const built = ((await import(`../bench/coverage/workbook/${set.name}.ts`)) as { buildWorkbookCases: () => WorkbookCase[] }).buildWorkbookCases()

  describe(`workbook features against Excel en-US: set ${set.name}`, () => {
    it('the stored cases are the current ones, all answered', () => {
      expect(set.cases).toEqual(built)
      expect(built.filter((c) => !set.answers[c.id])).toEqual([])
    })

    for (const c of set.cases) {
      it(`${c.id}: ${c.about}`, () => {
        const reads = engineReads(c)
        const excel = set.answers[c.id].reads
        expect(reads.length).toBe(excel.length)
        const diffs = excel.flatMap((a, i) => differences(c, a.values, reads[i]).map((d) => `read ${i + 1}: ${d}`))
        const known = KNOWN[`${set.name}:${c.id}`]
        if (known) expect(diffs.length, known).toBeGreaterThan(0)
        else expect(diffs).toEqual([])
      })
      if (c.predicted)
        it(`${c.id}: the answer predicted before Excel ran`, () => {
          const reads = engineReads(c)
          expect(c.predicted!.flatMap((p, i) => differences(c, p, reads[i]))).toEqual([])
        })
    }
  })
}
