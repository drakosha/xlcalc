/**
 * The engine's prediction of a pivot set, written before Excel's answers are read: each case replayed over the books
 * Excel saved (`phases.ts`), and for each way a phase is made (`<case>/<phase>/<how>`) the cells of every PivotTable's
 * area on `P` (its Filters area too) as the engine lays it out, or the engine's refusal. Run:
 * `npx tsx bench/coverage/pivot/predict.ts group`; it writes `bench/coverage/pivot/<set>.predicted.json`.
 */
import { writeFileSync } from 'node:fs'
import type { BookInput, CellValue } from '../../../src/contract.js'
import { formatAddress } from '../../../src/model.js'
import { parseRange } from '../../../src/structure/area.js'
import type { PivotSet } from './cases.js'
import { phaseCount, replays, rounded, shown } from './phases.js'

/** Each PivotTable of `P` by name, its area's cells. */
function areas(book: BookInput): Record<string, CellValue[][]> {
  const sheet = book.sheets.find((s) => s.name === 'P')
  return Object.fromEntries(
    (sheet?.pivotTables ?? []).map((t) => {
      const x = parseRange(t.location.range)!
      const top = x.r1 - (t.location.pageRows ? t.location.pageRows + 1 : 0)
      return [t.name, rounded(shown(book, 'P', `${formatAddress(Math.max(0, top), x.c1)}:${formatAddress(x.r2, x.c2)}`))]
    }),
  )
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const set = process.argv[2]
  const { buildPivotCases } = (await import(`./${set}.ts`)) as PivotSet
  const out: Record<string, Record<string, CellValue[][]> | { refused: string }> = {}
  for (const c of buildPivotCases()) for (const { at, result } of await replays(set, c, phaseCount(c))) out[at] = result.ok ? areas(result.book) : { refused: result.message }
  writeFileSync(new URL(`${set}.predicted.json`, import.meta.url), JSON.stringify(out) + '\n')
  console.log(`${Object.keys(out).length} phases, ${Object.values(out).filter((v) => 'refused' in v).length} refused`)
}
