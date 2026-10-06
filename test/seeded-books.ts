/** Seeded random books of the structure tests: numbers in A1:H20 of three sheets, formulas in J1:N20 over them. */
import type { BookInput, CellInput } from '../src/contract.js'
import { indexToCol } from '../src/model.js'
import { mulberry32 } from '../src/synth.js'

export const SHEETS = ['S1', 'S2', 'S 3']

export function randomBook(seed: number): BookInput {
  const rnd = mulberry32(seed)
  const int = (n: number) => Math.floor(rnd() * n)
  const flag = () => (rnd() < 0.3 ? '$' : '')
  const cell = () => `${flag()}${indexToCol(int(8))}${flag()}${1 + int(20)}`
  const sheet = (own: string) => {
    const s = SHEETS[int(SHEETS.length)]
    return rnd() < 0.6 ? '' : s === own && rnd() < 0.5 ? '' : `'${s}'!`
  }
  const term = (own: string) => {
    if (rnd() < 0.6) return sheet(own) + cell()
    const [r1, c1] = [int(20), int(8)]
    return `SUM(${sheet(own)}${flag()}${indexToCol(c1)}${flag()}${r1 + 1}:${flag()}${indexToCol(c1 + int(8 - c1))}${flag()}${r1 + 1 + int(20 - r1)})`
  }
  return {
    sheets: SHEETS.map((name) => {
      const cells: Record<string, CellInput> = {}
      for (let r = 1; r <= 20; r++) for (let c = 0; c < 8; c++) if (rnd() < 0.8) cells[`${indexToCol(c)}${r}`] = { value: int(1000) }
      for (let r = 1; r <= 20; r++) for (let c = 9; c < 14; c++) if (rnd() < 0.7) cells[`${indexToCol(c)}${r}`] = { formula: Array.from({ length: 1 + int(3) }, () => term(name)).join(rnd() < 0.5 ? '+' : '-') }
      return { name, cells }
    }),
  }
}
