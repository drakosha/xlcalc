/**
 * The engine's answers to a held-out table set as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/tables/predict.ts <set>` writes `tables/<set>.predicted.json` (values by `Sheet!A1`, formulas as
 * stored text, the tables, the steps refused).
 */
import { writeFileSync } from 'node:fs'
import type { TableSet } from './cases.js'
import { engineAnswer } from '../../../test/tables-cases.js'

const set = process.argv[2]
const { buildTableCases } = (await import(`./${set}.js`)) as TableSet
const out: Record<string, unknown> = {}
for (const c of buildTableCases()) {
  const { values, formulas, tables, refused } = engineAnswer(c)
  out[c.id] = { values, formulas, tables, refused: Object.keys(refused).map(Number) }
}
writeFileSync(new URL(`./${set}.predicted.json`, import.meta.url), JSON.stringify(out, null, 1) + '\n')
