/**
 * The engine's answers to the xlsx step's new files, written before Excel saw them: `npx tsx bench/coverage/xlsx/predict.ts`
 * writes `xlsx/predicted.json`. For each file of `filters.ts`, the filter mode the reader gives sheet S and the values of
 * sheet F; for the round trip's books in `PREDICTED`, the values of their formulas, computed as the reader reads the file
 * written (a row of height 0 is hidden), and the places where the book read back from what Excel saved differs from the
 * one written (none predicted). `test/xlsx-excel.test.ts` checks Excel's answers against them, the misses named.
 */
import { writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import type { CellValue } from '../../../src/contract.js'
import { compute } from '../../../src/index.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import { buildRoundTrips, XLSX_CLOCK } from './books.js'
import { FILTER_FILES, filterFile } from './filters.js'

/** The round trip's books predicted. */
export const PREDICTED = ['hidden', 'zero-sizes', 'filtered', 'filtered-table', 'date1904', 'iteration', 'iteration-default', 'builtin', 'builtin-ru']

export interface Predicted {
  filters: Record<string, { filtered: boolean; values: Record<string, CellValue> }>
  roundtrip: Record<string, { differences: string[]; values: Record<string, CellValue> }>
}

export const PREDICTED_FILE = fileURLToPath(new URL('./predicted.json', import.meta.url))

/** The values of a result's formulas, by `Sheet!A1`. */
const formulaValues = (r: ReturnType<typeof compute>) =>
  Object.fromEntries(r.sheets.flatMap((s) => Object.entries(s.cells).flatMap(([a, c]) => (c.formula !== undefined ? [[`${s.name}!${a}`, c.value]] : []))))

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  const out: Predicted = { filters: {}, roundtrip: {} }
  for (const c of FILTER_FILES) {
    const read = await readXlsx(await filterFile(c))
    out.filters[c.id] = { filtered: !!read.book.sheets[0].filtered, values: formulaValues(compute(read.book, { clock: XLSX_CLOCK })) }
  }
  for (const c of (await buildRoundTrips()).filter((x) => PREDICTED.includes(x.id))) {
    const read = await readXlsx((await writeXlsx(c.book)).data)
    out.roundtrip[c.id] = { differences: [], values: formulaValues(compute(read.book, { clock: XLSX_CLOCK })) }
  }
  writeFileSync(PREDICTED_FILE, JSON.stringify(out, null, 1) + '\n')
  console.log(`predicted: ${Object.keys(out.filters).length} filter files, ${Object.keys(out.roundtrip).length} round-trip books`)
}
