/**
 * The engine's answers to a held-out table set as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/tables/predict.ts <set>` writes `tables/<set>.predicted.json` (values by `Sheet!A1`, formulas as
 * stored text, the tables, the steps refused, the names of a case that adds them).
 */
import { fileURLToPath } from 'node:url'
import { setModule, writeAnswers } from '../stored.js'
import type { TableSet } from './cases.js'
import { addsNames } from '../tables.js'
import { engineAnswer } from '../../../test/tables-cases.js'

const set = process.argv[2]
const { buildTableCases } = await setModule<TableSet>('tables', set)
const out: Record<string, unknown> = {}
for (const c of buildTableCases()) {
  const { values, formulas, tables, refused, names } = engineAnswer(c)
  out[c.id] = { values, formulas, tables, refused: Object.keys(refused).map(Number), ...(addsNames(c) ? { names } : {}) }
}
writeAnswers(fileURLToPath(new URL(`./${set}.predicted.json`, import.meta.url)), out)
