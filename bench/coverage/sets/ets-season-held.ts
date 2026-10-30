/**
 * FORECAST.ETS's detection of the season and its longest season, held out: 102 series (`ets-season-held.series.json`)
 * that tell its branches apart: pure cosines of every period from 2 to 24, noise and a trend over five periods, lengths about 2.5 periods, two periods of near equal strength, short series of 2 to 10
 * points, lengths around 2,048, a constant, an exact line, missing slots and blank values completed and as zeros, seasons
 * given about 70 % of the points and about the largest seasonality taken. Each series is asked its season (if detected),
 * every statistic, three forecasts and an interval. Answers of Excel for Windows: `test/excel/win/cases/ets-season-held.json`.
 * `PREDICTED`, the engine's answers, was committed before Excel ran it.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'
import { CI, ETS, SEA, seriesSheet, STAT, steps } from './ets/sheet.js'

interface Held {
  id: string
  y: (number | null)[]
  t: number[]
  season: number
  completion: number
  targets: number[]
  level: number
}

const SERIES: Held[] = JSON.parse(readFileSync(new URL('./ets-season-held.series.json', import.meta.url), 'utf8'))
const { cells, series } = seriesSheet()

function build(): Probe[] {
  const { out, add } = collector()
  for (const s of SERIES) {
    const [t, v] = series(s.t, s.y)
    const opts = `${s.season},${s.completion},1`
    if (s.season === 1) add(s.id, SEA, `${SEA}(${v},${t},${s.completion},1)`)
    add(s.id, STAT, ...steps(8).map((k) => `${STAT}(${v},${t},${k},${opts})`))
    add(s.id, ETS, ...s.targets.map((x) => `${ETS}(${x},${v},${t},${opts})`))
    add(s.id, CI, `${CI}(${s.targets[1]},${v},${t},${s.level},${opts})`)
  }
  return out
}

const PROBES = build()

export const SHEETS: SheetCells = { E: cells }

export const buildProbes = (): Probe[] => PROBES

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./ets-season-held.predicted.json', import.meta.url), 'utf8'))
