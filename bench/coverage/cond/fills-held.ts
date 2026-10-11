/**
 * Held out: fills as files other than Excel's may hold them, in a cell's format and in a conditional format's (a dxf):
 * `patternType="none"` with colours, a solid fill with one colour of the two, a fill without its pattern type, a
 * gradient, and a gradient beside a pattern in one fill (two kinds where xlsx has one). `writeXlsx` writes each book with
 * fills of placeholder colours, and each case's fills are put in by hand in their place (`raw`). Sheet S: A1:A7 each in
 * a cell format of its own (`CELL_FILLS`, the first one on A1); B1:C6 under one rule per row (`DXF_FILLS`), B without a
 * fill of its own and C in yellow, so a dxf that takes the fill away shows on C. The read looks show the fill Excel shows.
 * Predicted before Excel ran it (`fills-held.predicted.json`).
 */
import type { CellInput, CellStyle, ConditionalFormat } from '../../../src/contract.js'
import type { StoredBook } from '../stored.js'
import type { RawEdit } from '../xlsx/sent.js'
import type { CondCases } from './cases.js'

const RED = 'FFFF0000'
const GREEN = 'FF00B050'
const GRADIENT = `<gradientFill degree="90"><stop position="0"><color rgb="${RED}"/></stop><stop position="1"><color rgb="FF0070C0"/></stop></gradientFill>`

/** The fills of A1:A7 as the file holds them: the `<fill>` element's body. */
export const CELL_FILLS = [
  `<patternFill patternType="none"><bgColor rgb="${RED}"/></patternFill>`,
  `<patternFill patternType="none"><fgColor rgb="${RED}"/><bgColor rgb="${GREEN}"/></patternFill>`,
  `<patternFill patternType="solid"><bgColor rgb="${RED}"/></patternFill>`,
  `<patternFill patternType="solid"><fgColor rgb="${RED}"/><bgColor rgb="${GREEN}"/></patternFill>`,
  `<patternFill><fgColor rgb="${RED}"/></patternFill>`,
  `<patternFill><bgColor rgb="${RED}"/></patternFill>`,
  GRADIENT,
]

/** The fills of the rules over B1:C6, one per row, as the file holds them. */
export const DXF_FILLS = [
  `<patternFill patternType="none"><bgColor rgb="${RED}"/></patternFill>`,
  `<patternFill patternType="none"><fgColor rgb="${RED}"/><bgColor rgb="${GREEN}"/></patternFill>`,
  '<patternFill patternType="none"/>',
  `<patternFill patternType="solid"><fgColor rgb="${RED}"/></patternFill>`,
  `<patternFill patternType="solid"><fgColor rgb="${RED}"/><bgColor rgb="${GREEN}"/></patternFill>`,
  `<patternFill><fgColor rgb="${RED}"/></patternFill>`,
  GRADIENT,
]

/** The placeholder colour of the n-th fill. */
const stand = (n: number) => `#0A0A${(n + 1).toString(16).padStart(2, '0').toUpperCase()}`
export const STYLES = 'xl/styles.xml'

/** A cell format's fill as `writeXlsx` writes it, and a dxf's. */
export const cellFill = (n: number) => `<patternFill patternType="solid"><fgColor rgb="FF${stand(n).slice(1)}"/><bgColor indexed="64"/></patternFill>`
export const dxfFill = (n: number) => `<patternFill><bgColor rgb="FF${stand(n).slice(1)}"/></patternFill>`

/** S with `cells` cell formats on A1 down and `dxfs` rules over B1:C1 down; the yellow of C is the style after them. */
export function book(cells: number, dxfs: number): StoredBook {
  const styles: CellStyle[] = Array.from({ length: cells }, (_, i) => ({ fill: { color: stand(i) } }))
  const yellow = styles.push({ fill: { color: '#FFFF00' } }) - 1
  const at: Record<string, CellInput> = {}
  for (let r = 1; r <= Math.max(cells, dxfs); r++) {
    if (r <= cells) at[`A${r}`] = { value: r, style: r - 1 }
    if (r <= dxfs) Object.assign(at, { [`B${r}`]: { value: r }, [`C${r}`]: { value: r, style: yellow } })
  }
  const rules: ConditionalFormat[] = Array.from({ length: dxfs }, (_, i) => ({ type: 'expression', ranges: [`B${i + 1}:C${i + 1}`], formula: 'TRUE', style: { fill: { color: stand(cells + i) } } }))
  return { styles, sheets: [{ name: 'S', cells: at, ...(rules.length ? { conditionalFormats: rules } : {}) }] }
}

export function buildCondCases(): CondCases {
  const both = `${GRADIENT}<patternFill patternType="solid"><fgColor rgb="${GREEN}"/></patternFill>`
  const cells: RawEdit[] = CELL_FILLS.map((to, i) => ({ part: STYLES, from: cellFill(i), to }))
  const dxfs: RawEdit[] = DXF_FILLS.map((to, i) => ({ part: STYLES, from: dxfFill(CELL_FILLS.length + i), to }))
  return {
    books: { fills: book(CELL_FILLS.length, DXF_FILLS.length), one: book(1, 1) },
    cases: [
      { id: 'fills', book: 'fills', read: [['S', 'A1:C7']], raw: [...cells, ...dxfs] },
      { id: 'cell-gradient-pattern', book: 'one', read: [['S', 'A1:C1']], raw: [{ part: STYLES, from: cellFill(0), to: both }] },
      { id: 'dxf-gradient-pattern', book: 'one', read: [['S', 'A1:C1']], raw: [{ part: STYLES, from: dxfFill(1), to: `${GRADIENT}<patternFill><bgColor rgb="${GREEN}"/></patternFill>` }] },
    ],
  }
}
