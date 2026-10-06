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
  return out
}
