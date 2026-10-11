/**
 * A book of spilled arrays and seeded edits of it: data under FILTER, the sizes of SEQUENCE, values typed into
 * spill ranges and cleared, anchors replaced, two arrays over one cell, a whole column, a cycle through a spill
 * range and breaking it, a size read from the array's own range, several edits at once.
 */
import type { BookInput, CellReplace, CellInput } from '../src/contract.js'
import { mulberry32 } from '../src/synth.js'

const f = (formula: string): CellInput => ({ formula })

export function spillBook(): BookInput {
  const s: Record<string, CellInput> = {}
  for (let r = 1; r <= 12; r++) s[`A${r}`] = { value: (r * 7) % 13 }
  Object.assign(s, {
    B1: { value: 3 },
    B2: { value: 2 },
    D1: f('SEQUENCE(B1)'),
    E1: f('FILTER(A1:A12,A1:A12>6)'),
    F1: f('SUM(E1#)'),
    F2: f('ROWS(D1#)'),
    F3: f('D2*10'),
    F4: f('SUM(D:D)+SUM(sp)'),
    H1: f('SEQUENCE(1,4)'),
    H3: f('TRANSPOSE(D1#)'),
    // a column and a row that cross at M3 when B2 is 3 or more
    M1: f('SEQUENCE(B2)'),
    L3: f('SEQUENCE(1,3)'),
    P1: f('SEQUENCE(3)+P9'),
  })
  const w: Record<string, CellInput> = { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, B1: f('A:A*2'), C1: f('COUNT(B:B)'), C2: f('SUM(B:B)'), C3: f('B5+1') }
  return {
    sheets: [
      { name: 'S', cells: s },
      { name: 'W', cells: w },
      { name: 'T', cells: { A1: f('S!E1#'), B1: f('SUM(S!D1#)'), B2: f('S!D3') } },
    ],
    names: [{ name: 'sp', ref: "'S'!$D$1#" }],
  }
}

const pick = <T>(rnd: () => number, xs: readonly T[]): T => xs[Math.floor(rnd() * xs.length)]

/** Seeded edits of `spillBook`. */
export function spillEdits(seed: number, count: number): CellReplace[][] {
  const rnd = mulberry32(seed)
  const S = (cell: string, input: CellInput | null): CellReplace => ({ sheet: 'S', cell, input })
  const one = (): CellReplace => {
    const x = rnd()
    if (x < 0.25) return S(`A${1 + Math.floor(rnd() * 12)}`, pick(rnd, [{ value: Math.floor(rnd() * 13) }, { value: 'x' }, null]))
    if (x < 0.4) return S(pick(rnd, ['B1', 'B2']), { value: Math.floor(rnd() * 6) })
    // a value or an empty text typed where arrays spill, or cleared
    if (x < 0.6) return S(pick(rnd, ['D2', 'D4', 'E3', 'E6', 'I1', 'K1', 'M2', 'M4', 'J3', 'D7']), pick(rnd, [{ value: 5 }, { value: '' }, null, null]))
    if (x < 0.7) return S('D1', pick(rnd, [f('SEQUENCE(B1)'), f('SEQUENCE(B1,2)'), f('B1*2'), f('A1:A3'), null]))
    if (x < 0.78) return { sheet: 'W', cell: pick(rnd, ['B50', 'B2', 'A9']), input: pick(rnd, [{ value: 4 }, null]) }
    if (x < 0.8) return S('B1', f('D3')) // D1's size from its own range: no spill from a fresh start, so no cycle
    if (x < 0.85) return S('B1', { value: 2 + Math.floor(rnd() * 3) })
    if (x < 0.9) return S('P9', pick(rnd, [f('P2'), f('SUM(P1#)'), { value: 5 }])) // P1 spills into P2: a cycle through its range
    if (x < 0.95) return { sheet: 'T', cell: 'C1', input: pick(rnd, [f('S!E1#*2'), f('ROWS(S!H1#)'), null]) }
    return S('F5', pick(rnd, [f('E2#'), f('SUM(H1#)'), null]))
  }
  return Array.from({ length: count }, () => (rnd() < 0.15 ? [one(), one(), one()] : [one()]))
}
