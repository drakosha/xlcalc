/**
 * Excel cases for DATE with a day far out of its month: how far Excel lets the day roll over (a date text that
 * reads as a serial number makes such a day, `DATE(2025,1,"1.5")` in ru-RU). Computed as dynamic array formulas
 * (Excel 365), in Excel en-US (`docs/excel-cases/date-day.json`).
 */
import type { Probe } from '../probes.js'

export const DYNAMIC = true

const DAYS = ['32766', '32767', '32768', '32769', '40000', '46143', '65535', '65536', '65537', '100000', '1E+10', '-32767', '-32768', '-32769', '-40000', '-65536', '-100000', '32767.5', '32768.5', '-32768.5']

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add = (formula: string) => out.push({ id: `DATE#${out.length + 1}`, group: 'DATE', fn: 'DATE', formula })
  for (const d of DAYS) add(`DATE(2025,1,${d})`)
  for (const d of ['32767', '32768', '-32768', '-32769']) add(`DATE(1950,7,${d})`)
  for (const m of ['32767', '32768', '-32768', '-32769', '1000', '-1000']) add(`DATE(2000,${m},1)`)
  // where months stop: the month alone, the year with it, the date reached
  for (const m of ['2000', '4000', '8000', '12000', '16000', '20000', '24000', '28000', '30000', '32000', '-2000', '-12000', '-22000'])
    for (const y of ['1900', '2000', '5000']) add(`DATE(${y},${m},1)`)
  for (const [y, m] of [['1900', '96000'], ['1900', '32767'], ['9999', '12'], ['9999', '13'], ['1900', '-11'], ['1900', '-12']]) add(`DATE(${y},${m},1)`)
  for (const m of ['32500', '32700', '32750', '32755', '32756', '32757', '32760', '32765', '32766']) for (const y of ['1900', '2000']) add(`DATE(${y},${m},1)`)
  return out
}
