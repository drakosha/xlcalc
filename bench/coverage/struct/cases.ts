/**
 * Structure cases for Excel (`run.ts struct <set>`): a book and one operation each. The base book has
 * every kind of reference the operations adjust: relative, absolute and mixed cells, ranges, whole
 * columns and rows, references from another sheet and to a sheet whose name needs quotes, global and
 * local names. No formula is inside a range it reads, before or after any operation of the sets, so
 * Excel never meets a circular reference.
 */
import type { FileParts } from '../case-book.js'
import type { BookInput, CellInput, CellRef, RangeRef, StructureOp } from '../../../src/contract.js'
import type { StoredBook } from '../stored.js'

export interface StructCase {
  id: string
  /** The key of the case's book in the set's `books`. */
  book: string
  op: StructureOp
  /** An operation taken before `op`, by Excel and the engine alike (cells merged, then split again); Excel for Windows only. */
  before?: StructureOp
  /** Cells outside the read rectangle to read after the operation. */
  extra?: CellRef[]
  /**
   * A copy that Excel applies otherwise than by copy and paste (`Range.Copy` with a destination): `fillDown` fills the
   * copy's target from its top row (Fill Down, `Range.FillDown`), `autofill` drags the source's fill handle over the
   * target (`Range.AutoFill`). The engine's `copyRange` is copy and paste: such a case records what Excel does.
   */
  via?: 'fillDown' | 'autofill'
}

export interface StructCases {
  books: Record<string, StoredBook>
  cases: StructCase[]
  /** Formulas written as dynamic array formulas (Excel 365), as the spill cases need. */
  dynamic?: true
  /** Rectangles whose look is read besides `LOOK` of `struct.ts` (`A1048570:C1048576` at the sheet's end). */
  look?: string[]
  /** Column widths are not compared: Excel takes the width of a span over most columns as the sheet's standard width when it loads the book, and reports the columns against that. */
  widths?: false
  /** What the files of the books hold beyond the model, by book: the autofilter of a `filtered` sheet. */
  files?: Record<string, FileParts>
}

export interface StructSet {
  buildStructCases(): StructCases
}

const f = (formula: string): CellInput => ({ formula })

/** S: data in A1:D12 and rows 30:31, formulas in F:H; T: formulas over S; `My Data`: a sheet name in quotes. */
export function baseBook(): BookInput {
  const s: Record<string, CellInput> = {}
  for (let r = 1; r <= 12; r++) {
    s[`A${r}`] = { value: r }
    s[`B${r}`] = { value: r * 10 }
    s[`C${r}`] = { value: `c${r}` }
  }
  for (let r = 2; r <= 9; r++) s[`D${r}`] = { value: 1 }
  s.A30 = { value: 100 }
  s.B31 = { value: true }
  Object.assign(s, {
    F1: f('A5'),
    F2: f('$A$5'),
    F3: f('$A5'),
    F4: f('A$5'),
    F5: f('SUM(A3:A8)'),
    F6: f('SUM($A$3:$A$8)'),
    F7: f('SUM(A5:A6)'),
    F8: f('SUM(A4:B4)'),
    F9: f('SUM(B:B)'),
    F10: f('SUM(A:B)'),
    F11: f('SUM(30:31)'),
    F12: f('SUM(A5:A1048576)'),
    G1: f('SUM(A1:A12)'),
    G2: f('SUM(B2:C9)'),
    G3: f('SUMIFS(B:B,A:A,">3")'),
    G4: f('VLOOKUP(5,A1:B12,2,FALSE)'),
    G5: f('A12'),
    G6: f('A1'),
    G7: f('A4+A5+A6'),
    G8: f('S!A5'),
    G9: f('SUM(S!A3:A8)'),
    G10: f('x'),
    G11: f('SUM(rng)'),
    G12: f('SUM(loc)'),
    H1: f('D5'),
    H2: f('SUM(C:D)'),
    H3: f('SUM(A30:B31)'),
    H4: f('A1048576'),
    H5: f('XFD1'),
    H6: f('SUM($B:$B)'),
    H7: f('SUM($30:$31)'),
    H8: f('SUM(B4:C4)'),
    H9: f('A6-A5'),
    H10: f('IF(A5>0,"pos",C5)'),
  })
  const t: Record<string, CellInput> = {
    A1: f('S!A5'),
    A2: f('SUM(S!A3:A8)'),
    A3: f('SUM(S!$A$3:$A$8)'),
    A4: f('SUM(S!B:B)'),
    A5: f('SUM(S!30:31)'),
    A6: f('S!$A5+S!A$5'),
    A7: f("SUM('My Data'!A1:A3)"),
    A8: f('A5'),
    A9: f('SUM(S!A5:B6)'),
    A10: f('x'),
    A11: f('SUM(tloc)'),
    A12: f('SUM(S!A1:D12)'),
    B1: { value: 'tee' },
  }
  const m: Record<string, CellInput> = { A1: { value: 1 }, A2: { value: 2 }, A3: { value: 3 }, B1: f('SUM(S!A1:A12)'), B2: f('S!C5') }
  return {
    sheets: [
      { name: 'S', cells: s },
      { name: 'T', cells: t },
      { name: 'My Data', cells: m },
    ],
    names: [
      { name: 'x', ref: "'S'!$A$5" },
      { name: 'rng', ref: "'S'!$A$3:$A$8" },
      { name: 'col', ref: "'S'!$B:$B" },
      { name: 'loc', ref: "'S'!$B$3:$B$8", scope: 'S' },
      { name: 'tloc', ref: "'S'!$A$5:$A$6", scope: 'T' },
    ],
  }
}

/** `Sheet!range` as a range of a sheet. */
export function rangeAt(ref: string): RangeRef {
  const i = ref.lastIndexOf('!')
  return { sheet: ref.slice(0, i), range: ref.slice(i + 1) }
}
