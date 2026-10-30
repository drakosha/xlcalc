/**
 * The engine's prediction of the set `show`, written before Excel's answers are read: each case's book as Excel saved
 * it, its PivotTable laid out again by `xlcalc/pivot` from the saved cache and refreshed from its source, and the values
 * of its area (`P`), dates as serials, errors by code, numbers to 15 significant digits. Run:
 * `npx tsx bench/coverage/pivot/show-predict.ts show`, written to `bench/coverage/pivot/<set>.predicted.json`.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import type { BookInput, CellValue } from '../../../src/contract.js'
import { createState } from '../../../src/engine.js'
import { dateToSerial, formatAddress } from '../../../src/model.js'
import { pivotArea } from '../../../src/pivot/area.js'
import { refreshPivots } from '../../../src/pivot/engine.js'
import { clockOptions } from '../../../src/result.js'
import { stateValues } from '../../../src/structure/values.js'
import { ERROR_CODE } from '../../../src/syntax/errors.js'
import { readXlsx } from '../../../src/xlsx.js'
import type { PivotSet } from './cases.js'

const ROOT = new URL('../../../', import.meta.url)

/** The values of the area of each PivotTable on `P`, laid out again from the saved cache or refreshed from the source. */
function laidOut(book: BookInput, fromSource: boolean): Record<string, { area: string; values: CellValue[][] }> | string {
  const values = stateValues(createState(book, clockOptions({ now: '2026-10-30', seed: 0 })))
  const r = refreshPivots(book, undefined, values, values.today!(), fromSource)
  if (!r.ok) return r.message
  const sheet = r.book.sheets.find((s) => s.name === 'P')!
  const out: Record<string, { area: string; values: CellValue[][] }> = {}
  for (const t of sheet.pivotTables ?? []) {
    const x = pivotArea(t)!
    const grid = Array.from({ length: x.r2 - x.r1 + 1 }, (_, r) =>
      Array.from({ length: x.c2 - x.c1 + 1 }, (_, c): CellValue => {
        const cell = sheet.cells[formatAddress(x.r1 + r, x.c1 + c)]
        if (cell?.formula?.startsWith('#')) return { error: ERROR_CODE[cell.formula] }
        const v = cell?.value
        if (v === undefined || v === null) return null
        if (typeof v === 'object' && 'date' in v) return dateToSerial(v.date)
        return typeof v === 'number' ? Number(v.toPrecision(15)) : (v as CellValue)
      }),
    )
    out[t.name] = { area: `${formatAddress(x.r1, x.c1)}:${formatAddress(x.r2, x.c2)}`, values: grid }
  }
  return out
}

const set = process.argv[2] ?? 'show'
const { buildPivotCases } = (await import(`./${set}.js`)) as PivotSet
const prediction: Record<string, unknown> = {}
for (const c of buildPivotCases()) {
  const { book } = await readXlsx(readFileSync(new URL(`test/excel/win/pivot/${set}/${c.id}.xlsx`, ROOT)))
  // the cells Excel saved on P are its answer: left out, so that nothing of them reaches the prediction
  book.sheets = book.sheets.map((s) => (s.name === 'P' ? { ...s, cells: {} } : s))
  const saved = laidOut(book, false)
  const source = laidOut(book, true)
  prediction[c.id] = JSON.stringify(saved) === JSON.stringify(source) ? saved : { saved, source }
}
writeFileSync(new URL(`bench/coverage/pivot/${set}.predicted.json`, ROOT), `${JSON.stringify(prediction, null, 1)}\n`)
