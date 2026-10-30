/**
 * The step FORECAST.ETS.STAT 8 reports for dates, held out after the histogram of the steps was fitted: months of 2020
 * and 2021 on the 1st, 15th, 28th, 30th and 31st (a short month's last day where it has none), quarters from the 15th
 * of January to April of 2019 to 2021, short quarterly timelines of 2 to 4 dates, and half years and years from 2018,
 * each with a forecast ten days after its last date. Answers of Excel for Windows: `test/excel/win/cases/ets-step-held.json`.
 * `PREDICTED`, the engine's answers, was committed before Excel ran it.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'
import { dates, ETS, seriesSheet, STAT, steps } from './ets/sheet.js'

const { cells, series } = seriesSheet()

function build(): Probe[] {
  const { out, add } = collector()
  const timeline = (group: string, ts: number[]) => {
    const [t, v] = series(ts, steps(ts.length))
    add(group, STAT, `${STAT}(${v},${t},8,0)`)
    add(group, ETS, `${ETS}(${ts[ts.length - 1] + 10},${v},${t},0)`)
  }
  for (const year of [2020, 2021]) for (const d of [1, 15, 28, 30, 31]) timeline(`m-${year}-${d}`, dates(year, 1, d, 1, 12))
  for (const year of [2019, 2020, 2021]) for (const month of [1, 2, 3, 4]) timeline(`q-${year}-${month}`, dates(year, month, 15, 3, 12))
  for (const year of [2020, 2021]) for (const n of [2, 3, 4]) timeline(`q-short-${year}-${n}`, dates(year, 1, 1, 3, n))
  for (const every of [6, 12]) timeline(`months-${every}`, dates(2018, 1, 1, every, 12))
  return out
}

const PROBES = build()

export const SHEETS: SheetCells = { E: cells }

export const buildProbes = (): Probe[] => PROBES

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./ets-step-held.predicted.json', import.meta.url), 'utf8'))

/** The predictions Excel did not confirm: a forecast between half years, which Excel counts in the calendar's days (`ets-step-held-2`). */
export const MISSED: readonly string[] = ['months-6#2']
