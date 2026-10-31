/**
 * The engine's prediction of a pivot set, written before Excel's answers are read: each case replayed over the books
 * Excel saved (`phases.ts`), and for each way a phase is made (`<case>/<phase>/<how>`) the cells of every PivotTable's
 * area on `P` (its Filters area too) as the engine lays it out, or the engine's refusal. A held-out set Excel has not run
 * yet is predicted with `--drive`: each case played on the engine alone (`drive.ts`), `<case>/<phase>/driven`, and the
 * operations refused, `<case>/<phase>/errors`. Run: `npx tsx bench/coverage/pivot/predict.ts group [--drive]`; add `--exact` to keep every bit of numbers instead of rounding to 15 significant digits; it
 * writes `bench/coverage/pivot/<set>.predicted.json`.
 */
import { writeFileSync } from 'node:fs'
import type { BookInput, CellValue } from '../../../src/contract.js'
import { formatAddress } from '../../../src/model.js'
import { parseRange } from '../../../src/structure/area.js'
import type { PivotSet } from './cases.js'
import { drive } from './drive.js'
import { phaseCount, replays, rounded, shown } from './phases.js'

/** Each PivotTable of `P` by name, its area's cells. */
function areas(book: BookInput, exact: boolean): Record<string, CellValue[][]> {
  const sheet = book.sheets.find((s) => s.name === 'P')
  return Object.fromEntries(
    (sheet?.pivotTables ?? []).map((t) => {
      const x = parseRange(t.location.range)!
      const top = x.r1 - (t.location.pageRows ? t.location.pageRows + 1 : 0)
      const cells = shown(book, 'P', `${formatAddress(Math.max(0, top), x.c1)}:${formatAddress(x.r2, x.c2)}`)
      return [t.name, exact ? cells : rounded(cells)]
    }),
  )
}

if (import.meta.url === `file://${process.argv[1]}`) {
  const set = process.argv[2]
  const { buildPivotCases } = (await import(`./${set}.ts`)) as PivotSet
  const out: Record<string, Record<string, CellValue[][]> | { refused: string } | { errors: string[] }> = {}
  for (const c of buildPivotCases())
    if (process.argv.includes('--drive'))
      drive(c).forEach((p, k) => {
        out[`${c.id}/${k}/driven`] = areas(p.book, process.argv.includes('--exact'))
        if (p.errors.length) out[`${c.id}/${k}/errors`] = { errors: p.errors }
      })
    else for (const { at, result } of await replays(set, c, phaseCount(c))) out[at] = result.ok ? areas(result.book, process.argv.includes('--exact')) : { refused: result.message }
  writeFileSync(new URL(`${set}.predicted.json`, import.meta.url), JSON.stringify(out) + '\n')
  console.log(`${Object.keys(out).length} entries, ${Object.values(out).filter((v) => 'refused' in v || 'errors' in v).length} refusals`)
}
