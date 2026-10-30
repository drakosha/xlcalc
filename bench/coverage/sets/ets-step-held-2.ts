/**
 * What the one miss of `ets-step-held` led to, held out: half years and years are calendar timelines, so a target
 * between two of their dates lies on the calendar's days, while two and four months apart are read as numbers. Timelines
 * of each from other months and days (the 31st, 29 February), squares as values, targets 30 to 400 days after the last
 * date. Answers of Excel for Windows: `test/excel/win/cases/ets-step-held-2.json`. `PREDICTED`, the engine's answers, was
 * committed before Excel ran it.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'
import { dates, ETS, seriesSheet, STAT, steps } from './ets/sheet.js'

const { cells, series } = seriesSheet()

function build(): Probe[] {
  const { out, add } = collector()
  const cases: [every: number, month: number, d: number, n: number][] = [[6, 1, 1, 10], [6, 3, 15, 9], [6, 8, 31, 8], [12, 1, 1, 9], [12, 2, 29, 8], [12, 7, 10, 10], [2, 1, 1, 12], [4, 1, 1, 12], [4, 2, 15, 10]]
  for (const [every, month, d, n] of cases) {
    const ts = dates(2016, month, d, every, n)
    const [t, v] = series(ts, steps(n).map((x) => x * x))
    const group = `every-${every}-${month}-${d}`
    add(group, STAT, `${STAT}(${v},${t},8,0)`)
    add(group, ETS, ...[30, 90, 200, 400].map((k) => `${ETS}(${ts[n - 1] + k},${v},${t},0)`))
  }
  return out
}

const PROBES = build()

export const SHEETS: SheetCells = { E: cells }

export const buildProbes = (): Probe[] => PROBES

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./ets-step-held-2.predicted.json', import.meta.url), 'utf8'))
