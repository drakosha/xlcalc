/**
 * A number minus a date written as text, held out: a host's books hold `R-"01.03.2025"` with the value `#VALUE!` where
 * the engine in ru-RU gives a number. Day-first dates with dots, impossible ones, two-digit years, times, decimals with
 * a comma and a point, spaces around and inside, other separators and words after the date; minus, plus, a unary minus
 * and VALUE beside each. Computed in Excel for Mac ru-RU and en-US (`test/excel/cases/ru-RU/date-text-minus.json`,
 * `test/excel/cases/date-text-minus.json`); `date-text-minus.predicted.json` and `date-text-minus.ru.predicted.json`,
 * the engine's answers, were committed before Excel ran the set.
 */
import type { Probe, SheetCells } from '../probes.js'

/** The number the formulas start from. */
export const SHEETS: SheetCells = { N: { A1: 45800 } }

const TEXTS = [
  '01.03.2025', '1.3.2025', '01.03.25', '31.12.2024', '29.02.2024', '29.02.2025', '31.02.2025', '32.01.2025', '01.13.2025', '00.01.2025',
  '01.03.2025 18:30', '18:30', '18.30', '1,5', '1.5', '1 234', '1 234,5', ' 01.03.2025', '01.03.2025 ', '01. 03.2025', '01 .03.2025',
  '01/03/2025', '01-03-2025', '2025-03-01', '2025.03.01', '01.03.2025г.', '01.03.2025 г.', '1 мар 2025', '1 March 2025', '3/1/2025',
]

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add = (group: string, formula: string) => out.push({ id: `${group}#${out.filter((p) => p.group === group).length + 1}`, group, formula })
  for (const t of TEXTS) {
    const s = JSON.stringify(t)
    add('minus', `N!A1-${s}`)
    add('plus', `${s}+0`)
    add('unary', `-${s}`)
    add('value', `VALUE(${s})`)
  }
  return out
}
