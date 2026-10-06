/**
 * Excel cases for AND, OR and XOR given the words for TRUE and FALSE as text, directly and in a cell: which
 * text they take as a boolean. Computed in Excel en-US (`docs/excel-cases/logical-text.json`) and ru-RU
 * (`run.ts probes logical-text ru`, `docs/excel-cases/ru-RU/logical-text.json`). The sheet `G` holds text.
 * Computed as dynamic array formulas (Excel 365).
 */
import type { Probe, SheetCells } from '../probes.js'

export const DYNAMIC = true

export const SHEETS: SheetCells = { G: { A1: 'TRUE', A2: 'ИСТИНА', A3: 'FALSE', A4: 'ЛОЖЬ' } }

const ARGS = ['"TRUE"', '"true"', '"FALSE"', '"ИСТИНА"', '"истина"', '"ЛОЖЬ"', '" TRUE"', '"TRUE",FALSE', '"FALSE",TRUE', '"ИСТИНА",FALSE', '"ЛОЖЬ",TRUE', 'G!A1', 'G!A2', 'G!A1:A4', 'G!A1,TRUE', '{"TRUE"}', '{"ИСТИНА",TRUE}']

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  for (const fn of ['AND', 'OR', 'XOR'])
    for (const a of ARGS) out.push({ id: `${fn}#${out.filter((p) => p.fn === fn).length + 1}`, group: fn, fn, formula: `${fn === 'XOR' ? '_xlfn.' : ''}${fn}(${a})` })
  return out
}
