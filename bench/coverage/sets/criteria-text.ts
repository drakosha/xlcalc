/**
 * Excel cases for criteria over text cells that read as numbers, dates or booleans: which text a number
 * criterion's `=` and `<>` take as that number, and whether `"TRUE"` is a boolean criterion. Computed in Excel
 * en-US (`docs/excel-cases/criteria-text.json`) and ru-RU (`run.ts probes criteria-text ru`,
 * `docs/excel-cases/ru-RU/criteria-text.json`). The sheet `N` holds one value per cell, each read alone.
 * Computed as dynamic array formulas (Excel 365).
 */
import type { Probe, SheetCells } from '../probes.js'

export const DYNAMIC = true

export const SHEETS: SheetCells = {
  N: { A1: 'TRUE', B1: true, C1: 'ИСТИНА', D1: '2', E1: 2, F1: '1,000', G1: '1/31/2025', H1: '50%', I1: '$1,000', J1: '1 1/2', K1: '2025-01-31', L1: '2.5', M1: '2,5', O1: 45688, P1: 0.5, Q1: 1000 },
}

const CASES: [cell: string, criterion: string][] = [
  ['A1', '"TRUE"'], ['B1', '"TRUE"'], ['A1', 'TRUE'], ['C1', '"ИСТИНА"'], ['B1', '"ИСТИНА"'], ['D1', '"<>2"'], ['E1', '"<>2"'], ['D1', '"=2"'], ['D1', '2'], ['D1', '">1"'],
  ['D1', '"<3"'], ['F1', '1000'], ['F1', '1'], ['G1', '45688'], ['H1', '0.5'], ['I1', '1000'], ['J1', '1.5'], ['K1', '45688'], ['K1', '"2025-01-31"'], ['G1', '"1/31/2025"'],
  ['L1', '2.5'], ['M1', '2.5'], ['L1', '"<>2.5"'], ['M1', '"<>2,5"'], ['H1', '"<>0.5"'],
  // criterion text read as a number against numbers
  ['O1', '">=1/1/2025"'], ['O1', '"1/31/2025"'], ['O1', '"Jan 31, 2025"'], ['O1', '"<2025-02-01"'], ['O1', '"31.01.2025"'], ['O1', '"31 янв 2025"'], ['P1', '"50%"'],
  ['Q1', '"$1,000"'], ['Q1', '"1,000"'], ['Q1', '"1 000"'], ['P1', '"1/2"'], ['P1', '">1 1/4"'],
]

export function buildProbes(): Probe[] {
  return CASES.map(([cell, c], i) => ({ id: `criteria#${i + 1}`, group: 'criteria', formula: `COUNTIF(N!${cell},${c})` }))
}
