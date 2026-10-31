/**
 * The engine's prediction of an areas set (`areas`, `areas-held`), written before Excel's answers are read: for each case
 * every pivot table's `PivotLocation.areas` as the library gives it, keyed as the harness keys Excel's reads (`PT1` for
 * a case played from the rows, `<sheet>!<name>` for a book opened as it is). A case played from the rows is laid out
 * by the engine alone (`drive.ts`); a book is read with `readXlsx`. Run: `npx tsx bench/coverage/pivot/areas-predict.ts
 * <set>`; it writes `bench/coverage/pivot/<set>.predicted.json`. The comparison with Excel: `test/pivot-areas.test.ts`.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import type { BookInput, PivotAreas } from '../../../src/contract.js'
import { placedPivots } from '../../../src/pivot/area.js'
import { readXlsx } from '../../../src/xlsx.js'
import type { PivotCase, PivotSet } from './cases.js'
import { drive } from './drive.js'

const ROOT = new URL('../../../', import.meta.url)

/** Each pivot table of a case's book at its end, its areas, keyed as the harness reads them. */
export async function predictedAreas(c: PivotCase): Promise<Record<string, PivotAreas | undefined>> {
  if (c.file) {
    const book: BookInput = (await readXlsx(readFileSync(new URL(c.file, ROOT)))).book
    return Object.fromEntries(book.sheets.flatMap((s) => (s.pivotTables ?? []).map((t) => [`${s.name}!${t.name}`, t.location.areas])))
  }
  // the PivotTable no operation laid out has its areas as a session gives it
  const book = placedPivots(drive(c).at(-1)!.book)
  return Object.fromEntries((book.sheets.find((s) => s.name === 'P')?.pivotTables ?? []).map((t) => [t.name, t.location.areas]))
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const set = process.argv[2]
  const { buildPivotCases } = (await import(`./${set}.ts`)) as PivotSet
  const out: Record<string, Record<string, PivotAreas | undefined>> = {}
  for (const c of buildPivotCases()) out[c.id] = await predictedAreas(c)
  writeFileSync(new URL(`${set}.predicted.json`, import.meta.url), JSON.stringify(out, null, 1) + '\n')
  console.log(`${Object.keys(out).length} cases`)
}
