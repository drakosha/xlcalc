/**
 * The engine's prediction of a cells set (`cells`, `cells-held`), written before Excel's answers are read: each book
 * read with `readXlsx`, and what each cell of each of its pivot tables is (`pivotCells`), keyed `sheet!name` as the
 * harness keys Excel's reads. Run: `npx tsx bench/coverage/pivot/cells-predict.ts <set>`; it writes
 * `bench/coverage/pivot/<set>.predicted.json`. The comparison with Excel: `test/pivot-cells.test.ts`.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import type { PivotCells } from '../../../src/contract.js'
import { pivotCells } from '../../../src/pivot/cells.js'
import { readXlsx } from '../../../src/xlsx.js'
import type { PivotCase, PivotSet } from './cases.js'

const ROOT = new URL('../../../', import.meta.url)

/** What each cell of each pivot table of a case's book is. */
export async function predictedCells(c: PivotCase): Promise<Record<string, PivotCells>> {
  const book = (await readXlsx(readFileSync(new URL(c.file!, ROOT)))).book
  return Object.fromEntries(book.sheets.flatMap((s) => (s.pivotTables ?? []).map((t) => [`${s.name}!${t.name}`, pivotCells(book, s.name, t.name)!])))
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const set = process.argv[2]
  const { buildPivotCases } = (await import(`./${set}.ts`)) as PivotSet
  const out: Record<string, Record<string, PivotCells>> = {}
  for (const c of buildPivotCases()) out[c.id] = await predictedCells(c)
  writeFileSync(new URL(`${set}.predicted.json`, import.meta.url), JSON.stringify(out) + '\n')
  console.log(`${Object.keys(out).length} cases`)
}
