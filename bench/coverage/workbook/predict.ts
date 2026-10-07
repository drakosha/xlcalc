/**
 * The engine's answers to a held-out workbook set as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/workbook/predict.ts <set>` writes `workbook/<set>.predicted.json`, which the set reads.
 */
import { writeFileSync } from 'node:fs'
import { cellsOf } from '../case-book.js'
import type { WorkbookSet } from './cases.js'
import { engineReads } from '../../../test/workbook-cases.js'
import { shownValue } from '../../../test/spill-shown.js'

const set = process.argv[2]
const { buildWorkbookCases } = (await import(`./${set}.js`)) as WorkbookSet
const out: Record<string, Record<string, unknown>[]> = {}
for (const c of buildWorkbookCases())
  out[c.id] = engineReads(c).map((r) => {
    const vals: Record<string, unknown> = {}
    for (const g of c.read) for (const a of cellsOf(g.range).addrs) {
      const v = shownValue(r, g.sheet, a)
      if (v) vals[`${g.sheet}!${a}`] = v
    }
    return vals
  })
writeFileSync(new URL(`./${set}.predicted.json`, import.meta.url), JSON.stringify(out, null, 1) + '\n')
