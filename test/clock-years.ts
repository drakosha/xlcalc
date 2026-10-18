/** Rows dated 5 January of 2025, 2026 or 2027, and the check that a criterion of that date without a year follows the clock's year. */
import { expect } from 'vitest'
import type { CellInput, Session } from '../src/index.js'

const year = (r: number) => 2025 + (r % 3)
const serial = (y: number) => (Date.UTC(y, 0, 5) - Date.UTC(1899, 11, 30)) / 86400000

/** Rows `from`…`to` of a sheet: in A the date of 5 January of the row's year, in B the row's number. */
export function datedRows(from: number, to: number): Record<string, CellInput> {
  const cells: Record<string, CellInput> = {}
  for (let r = from; r <= to; r++) {
    cells[`A${r}`] = { value: serial(year(r)) }
    cells[`B${r}`] = { value: r }
  }
  return cells
}

/** Moves the clock of `s` across years: each of `formulas` (sheet S) sums the numbers of the rows `from`…`to` of the clock's year. */
export function followsClockYear(s: Session, formulas: string[], from: number, to: number): void {
  for (const y of [2026, 2027, 2025, 2027, 2026]) {
    s.setClock({ now: `${y}-10-04T12:00:00Z`, seed: 7 })
    let sum = 0
    for (let r = from; r <= to; r++) if (year(r) === y) sum += r
    expect(formulas.map((a) => s.getCell('S', a)?.value)).toEqual(formulas.map(() => sum))
  }
}
