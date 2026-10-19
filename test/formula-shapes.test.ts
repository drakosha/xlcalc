/**
 * Formulas prepared through their shape (`src/prepare.ts`, `src/syntax/shape.ts`) against each prepared on its own, on a
 * workbook of traps: one shape on two sheets whose names of one spelling differ, a shape a name of one sheet makes valid
 * there alone, `[@Qty]` in two tables and outside them, texts that differ in case, spaces and sheet quotes, references
 * at the sheet's edges, whole columns and rows, `$` mixed, spill references, references across sheets, other workbooks,
 * LET, array constants, numbers written otherwise than printed, strings with digits, unions, `@`, rejections. Each
 * formula's text and AST, or its whole rejection, are the same either way; the formulas meant to share a shape share it.
 */
import { describe, expect, it } from 'vitest'
import type { BookInput, CellInput } from '../src/contract.js'
import { createState } from '../src/engine.js'
import { parseAddress } from '../src/model.js'
import { prepare, prepareShaped } from '../src/prepare.js'
import type { Template } from '../src/syntax/shape.js'

/** The shapes of one build, each formula's key kept as it is looked up. */
class Watched extends Map<string, Template | null> {
  last: string | undefined
  override get(key: string) {
    this.last = key
    return super.get(key)
  }
}

/** How the formulas of a group share shapes: all one, each its own, or none (prepared on their own). */
type Sharing = 'one' | 'apart' | 'none'

const groups: { sheet: string; sharing?: Sharing; formulas: string[] }[] = [
  // a name of one spelling each sheet has its own of, spelled otherwise
  { sheet: 'Data', sharing: 'one', formulas: ['A2*rate', 'A3*rate', 'A4*rate'] },
  { sheet: 'Calc', sharing: 'one', formulas: ['A2*rate', 'A3*rate', 'A4*rate'] },
  { sheet: 'Data', sharing: 'one', formulas: ['A1*Calc!rate', 'A2*Calc!rate'] },
  // a name only Data sees: valid there, rejected on Other
  { sheet: 'Data', sharing: 'one', formulas: ['B2+only', 'B3+only'] },
  { sheet: 'Other', sharing: 'one', formulas: ['B2+only', 'B3+only'] },
  { sheet: 'Other', formulas: ['B4+nope', 'Tax*2'] },
  // case, spaces and sheet quotes
  { sheet: 'Calc', sharing: 'one', formulas: ['sum(a1:a3)', 'SUM( A2:A4 )', 'Sum(A3:A5)'] },
  { sheet: 'Calc', sharing: 'one', formulas: ["'Data'!B1+1", 'Data!B2+1', 'Data!B3 + 1'] },
  { sheet: 'Calc', sharing: 'apart', formulas: ["'data'!b4+1", 'DATA!B5+1', "'Data'!B6+1.5"] },
  { sheet: 'Calc', sharing: 'one', formulas: ['IF(A1,TRUE,false)', 'IF(A2,TRUE,false)'] },
  { sheet: 'Calc', sharing: 'apart', formulas: ['IF(A3,true,FALSE)', 'IF(A4,TRUE,false)+0'] },
  // references of every kind in one shape, at the sheet's edges, `$` mixed, spilled
  { sheet: 'Edge', sharing: 'one', formulas: ['SUM(A1)', 'SUM(A:A)', 'SUM(1:1)', 'SUM($A:$A)', 'SUM($1:$1)', 'SUM(A$1:$B2)', 'SUM(E2#)', 'SUM(XFD1048576)', 'SUM(B2:A1)'] },
  { sheet: 'Edge', sharing: 'one', formulas: ['$A1+A$1+$A$1+A1', '$A2+B$1+$A$1+B2', 'XFC1+XFD1+A1048576+XFD1048575'] },
  { sheet: 'Edge', sharing: 'one', formulas: ['E2#+1', 'E3#+1'] },
  { sheet: 'Edge', sharing: 'one', formulas: ['A1:A5 B2:C2', 'A2:A6 B3:C3'] },
  { sheet: 'Edge', sharing: 'one', formulas: ["SUM('Data:Calc'!A1)", 'SUM(Data:Calc!A2)', "SUM('Data:Calc'!A1:B2)"] },
  // other workbooks
  { sheet: 'Calc', sharing: 'one', formulas: ['[1]Sheet1!A1*2', "'[1]Sheet1'!A2*2"] },
  { sheet: 'Calc', sharing: 'apart', formulas: ["'[1]sheet1'!A3*2", '[Book2.xlsx]Sheet1!A4*2'] },
  { sheet: 'Calc', sharing: 'one', formulas: ['[1]!Rate+A1', '[1]!Rate+A2'] },
  { sheet: 'Calc', sharing: 'one', formulas: ["'[1]Sheet1'!#REF!+A1", "'[1]Sheet1'!#REF!+A2"] },
  { sheet: 'Calc', formulas: ['[3]Sheet1!A1', '[1]Nope!A1', '[1]!Missing+1'] },
  // LET, whose cells' words may be names
  { sheet: 'Calc', sharing: 'none', formulas: ['LET(x,A1,x*2)', 'LET(x,A2,x*2)', 'LET(a1,1,a1+A2)', 'LET(A1,1,A1+A3)', 'LET(b1,1,b1+B1)'] },
  // array constants keep their numbers; other numbers are holes, printed as Excel prints them
  { sheet: 'Calc', sharing: 'one', formulas: ['SUM({1,2;3,4})+A1', 'SUM({1,2;3,4})+A2'] },
  { sheet: 'Calc', sharing: 'apart', formulas: ['SUM({1,2;3,5})+A3', 'SUM({1,2;3,-4})+A4', 'INDEX({-1,"a",TRUE},1)&A4'] },
  { sheet: 'Calc', sharing: 'one', formulas: ['A1*1E3', 'A2*1000', 'A3*.5', 'A4*0.5', 'A5*1.0', 'A6*1.50', 'A7*1E-7', 'A8*123456789012345678'] },
  { sheet: 'Calc', sharing: 'one', formulas: ['-A1+2', '-A2+3', '-A3+1E2'] },
  // strings with digits are no holes
  { sheet: 'Calc', sharing: 'apart', formulas: ['"A1"&A1', '"A2"&A2', 'A1&"1"', 'A1&"""1"""'] },
  { sheet: 'Calc', sharing: 'one', formulas: ['"x"&A1', '"x"&A2'] },
  // unions, `@`, percent, error values
  { sheet: 'Calc', sharing: 'one', formulas: ['SUM((A1:A3,C1:C3))', 'SUM((A2:A4,C2:C4))'] },
  { sheet: 'Calc', sharing: 'one', formulas: ['@A1:A5', '@A2:A6', '@$A$1:$A$9'] },
  { sheet: 'Calc', sharing: 'one', formulas: ['A1%', 'A2%'] },
  { sheet: 'Calc', sharing: 'one', formulas: ['#REF!+A1', '#REF!+A2'] },
  { sheet: 'Calc', formulas: ["'Data'!#REF!", '#N/A', '#REF!A1+1', "'Data'!#REF!#"] },
  // rejections: one shape rejected at each of its formulas, and texts that do not read or are too long
  { sheet: 'Calc', sharing: 'one', formulas: ["'Nope'!A1+1", "'Nope'!A2+1"] },
  { sheet: 'Calc', formulas: ['FOO(A1)', 'A1+', 'SUM(A1', 'SUM(A1,)', 'IF()', '', '   ', 'A1+"open'] },
  { sheet: 'Calc', sharing: 'none', formulas: [`A1${' '.repeat(9000)}+1`] },
  { sheet: 'Calc', sharing: 'one', formulas: ['A1+1', 'A2 +  1'] },
]

/** Tables `T1` and `T2` on Data, each with a column Qty and its own column of `[@Qty]*2`, and the same formula outside them. */
function tableCells(): Record<string, CellInput> {
  const cells: Record<string, CellInput> = {}
  for (const [c, head] of [['D', 'Qty'], ['E', 'Price'], ['F', 'Total'], ['H', 'Qty'], ['I', 'Cost'], ['J', 'Total']]) cells[`${c}1`] = { value: head }
  for (let r = 2; r <= 4; r++) {
    for (const c of ['D', 'E', 'H', 'I']) cells[`${c}${r}`] = { value: r * 3 }
    cells[`F${r}`] = { formula: '[@Qty]*2' }
    cells[`J${r}`] = { formula: '[@Qty]*2' }
  }
  cells.L2 = { formula: '[@Qty]*2' }
  cells.L3 = { formula: 'T1[@Qty]*2' }
  return cells
}

/** The formulas of the workbook rejected: sheet, text (a long run of spaces shortened), check code. */
const REJECTED = [
  'Data "[@Qty]*2" unknown_name',
  'Data "T1[@Qty]*2" syntax',
  'Calc "[3]Sheet1!A1" external_ref',
  'Calc "[1]Nope!A1" external_ref',
  `Calc "'Nope'!A1+1" unknown_sheet`,
  `Calc "'Nope'!A2+1" unknown_sheet`,
  'Calc "FOO(A1)" unknown_function',
  'Calc "A1+" syntax',
  'Calc "SUM(A1" syntax',
  'Calc "IF()" syntax',
  'Calc "" empty',
  'Calc "   " empty',
  'Calc "A1+\\"open" syntax',
  'Calc "A1 … +1" too_long',
  'Other "B2+only" unknown_name',
  'Other "B3+only" unknown_name',
  'Other "B4+nope" unknown_name',
]

function trapBook(): { book: BookInput; places: { sheet: string; cell: string; group: number }[] } {
  const sheets: Record<string, Record<string, CellInput>> = { Data: tableCells(), Calc: {}, Other: {}, Edge: {} }
  const places: { sheet: string; cell: string; group: number }[] = []
  const next: Record<string, number> = {}
  groups.forEach((g, i) => {
    for (const formula of g.formulas) {
      // each sheet's formulas down columns M and on, past the cells the formulas read
      const k = (next[g.sheet] = (next[g.sheet] ?? 0) + 1)
      const cell = g.sheet === 'Edge' && k === 1 ? 'XFD1' : g.sheet === 'Edge' && k === 2 ? 'A1048576' : `M${k + 10}`
      sheets[g.sheet][cell] = { formula }
      places.push({ sheet: g.sheet, cell, group: i })
    }
  })
  const book: BookInput = {
    sheets: Object.entries(sheets).map(([name, cells]) => ({
      name,
      cells: { ...Object.fromEntries(Array.from({ length: 10 }, (_, r) => [`A${r + 1}`, { value: r + 1 }])), ...cells },
      ...(name === 'Data'
        ? {
            tables: [
              { name: 'T1', range: 'D1:F4', columns: [{ name: 'Qty' }, { name: 'Price' }, { name: 'Total' }] },
              { name: 'T2', range: 'H1:J4', columns: [{ name: 'Qty' }, { name: 'Cost' }, { name: 'Total' }] },
            ],
          }
        : {}),
    })),
    names: [
      { name: 'RATE', scope: 'Data', ref: "'Data'!$A$1" },
      { name: 'Rate', scope: 'Calc', ref: "'Calc'!$A$1" },
      { name: 'Only', scope: 'Data', ref: "'Data'!$B$1" },
      { name: 'Tax', ref: "'Data'!$A$2" },
    ],
    externalBooks: [{ target: 'Book2.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 1, A2: 2, A3: 3, A4: 4 } }], names: [{ name: 'Rate', ref: "'Sheet1'!$A$1" }] }],
  }
  return { book, places }
}

describe('formulas prepared through their shape', () => {
  const { book, places } = trapBook()
  const state = createState(book, { now: 46000, seed: 7 })
  const shapes = new Watched()
  // each formula in the order the state's build takes them, through its shape and on its own
  const seen = state.input.sheets.flatMap((sheet, s) =>
    Object.entries(sheet.cells)
      .filter(([, c]) => c.formula !== undefined)
      .map(([cell, c]) => {
        const { row, col } = parseAddress(cell)
        shapes.last = undefined
        const shaped = prepareShaped(state, shapes, s, row, col, c.formula!)
        return { sheet: sheet.name, cell, text: c.formula!, key: shapes.last, shaped, alone: prepare(state, s, row, col, c.formula!) }
      }),
  )

  it('gives each formula the text and AST, or the rejection, it has on its own', () => {
    expect(seen.length).toBe(places.length + 8)
    for (const x of seen) expect(x.shaped, `${x.sheet}!${x.cell}`).toStrictEqual(x.alone)
  })

  it('rejects where a formula on its own is rejected, and takes the rest', () => {
    const rejected = seen.flatMap((x) => (x.alone.ok ? [] : [`${x.sheet} ${JSON.stringify(x.text.replace(/ {4,}/, ' … '))} ${x.alone.invalid.code}`]))
    expect(rejected).toEqual(REJECTED)
    expect(seen.find((x) => x.sheet === 'Data' && x.cell === 'F2')!.shaped).toMatchObject({ ok: true, text: 'T1[[#This Row],[Qty]]*2' })
    expect(seen.find((x) => x.sheet === 'Data' && x.cell === 'J2')!.shaped).toMatchObject({ ok: true, text: 'T2[[#This Row],[Qty]]*2' })
  })

  it.each(groups.map((g, i) => [i, g] as const).filter(([, g]) => g.sharing))('group %i shares shapes as meant', (i, g) => {
    const keys = places.filter((p) => p.group === i).map((p) => seen.find((x) => x.sheet === p.sheet && x.cell === p.cell)!.key)
    if (g.sharing === 'none') expect(keys).toEqual(keys.map(() => undefined))
    else if (g.sharing === 'apart') expect(new Set(keys).size).toBe(keys.length)
    else {
      expect(new Set(keys).size).toBe(1)
      const t = shapes.get(keys[0]!)
      // a rejected shape keeps no template; a taken one is ready
      if (t !== null) expect(t?.at).toBeDefined()
    }
  })
})
