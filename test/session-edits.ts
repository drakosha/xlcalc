/**
 * Seeded session edits over the synthetic book: values of the lookup table, criteria and data, formulas with
 * and without arrays, invalid formulas, clearing, a cycle and breaking it, several edits at once.
 */
import type { BookInput, CellEdit, CellInput } from '../src/contract.js'
import { classAEdits, classBEdit, mulberry32, S, type SynthBook } from '../src/synth.js'

/** The book with the edits applied, as a full computation reads it. */
export function apply(book: BookInput, edits: CellEdit[]): BookInput {
  return {
    names: book.names,
    sheets: book.sheets.map((s) => {
      const mine = edits.filter((e) => e.sheet === s.name)
      if (!mine.length) return s
      const cells = { ...s.cells }
      for (const e of mine) {
        if (e.input === null) delete cells[e.cell]
        else cells[e.cell] = e.input
      }
      return { name: s.name, cells }
    }),
  }
}

/** Seeded edits: a mix of values and formulas. */
export function edits(seed: number, count: number, synth: SynthBook): CellEdit[][] {
  const rnd = mulberry32(seed)
  const a = classAEdits(synth, seed, count)
  const formulas: CellInput[] = [
    { formula: `SUMIFS('${S.data}'!D:D,'${S.data}'!B:B,$A3,'${S.data}'!E:E,"расход")` },
    { formula: 'SUM(B3:M3)*2' },
    { formula: 'SUM(B3:M3),1' }, // invalid: syntax
    { formula: 'NOPE(1)' },
    { formula: 'ROUND(B3/3,2)' },
    // arrays: lifted operators and functions, functions with an array answer, an array result (#SPILL!)
    { formula: 'SUMPRODUCT((B4:M4>0)*B4:M4)' },
    { formula: 'MAX(IF(B5:M5>100,B5:M5,0))+SUM(B6:M6*{1,0,1,0,1,0,1,0,1,0,1,0})' },
    { formula: 'INDEX(SORT(B4:M4,1,-1,TRUE),1,1)+COUNTA(UNIQUE(B5:M5,TRUE))' },
    { formula: `SUM(FILTER('${S.data}'!D2:D40,'${S.data}'!D2:D40>0,0))` },
    { formula: 'B4:M4*2' },
  ]
  const out: CellEdit[][] = []
  for (let i = 0; i < count; i++) {
    const x = rnd()
    if (x < 0.4) out.push([a[i]])
    else if (x < 0.6) out.push([classBEdit(synth, seed * 1000 + i)])
    else if (x < 0.8) out.push([{ sheet: S.months, cell: `${'BCDEFGHIJKLMN'[Math.floor(rnd() * 12)]}${3 + Math.floor(rnd() * 5)}`, input: formulas[Math.floor(rnd() * formulas.length)] }])
    else if (x < 0.85) out.push([{ sheet: S.months, cell: 'P1', input: { formula: 'P2+1' } }, { sheet: S.months, cell: 'P2', input: { formula: 'P1+1' } }]) // a cycle
    else if (x < 0.9) out.push([{ sheet: S.months, cell: 'P2', input: { value: 1 } }]) // breaks the cycle
    else if (x < 0.95) out.push([{ sheet: S.params, cell: 'B3', input: null }]) // clearing
    else out.push([a[i], classBEdit(synth, seed * 1000 + i), { sheet: S.items, cell: 'B2', input: { value: 3 } }]) // several at once
  }
  return out
}
