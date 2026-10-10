/**
 * Seeded random changes of every kind a session can undo, for the undo tests (`session-undo.test.ts`): the edits and
 * row operations of `session-fuzz.ts` over its book, with a table, names, styles and a SUBTOTAL added, and from a
 * stream of their own patches, new names and styles, layout edits (heights, widths, AutoFit, formats, hidden rows, merged
 * areas over spills, conditional formats, frozen panes, validation) and every structure operation, sheets renamed, inserted,
 * deleted, moved, hidden and shown among them; some are refused.
 */
import type { BookInput, CellPatch, CellStyle, LayoutEdit, NameInput, StructureOp } from '../src/contract.js'
import { indexToCol } from '../src/model.js'
import { mulberry32 } from '../src/synth.js'
import { fuzz, patches, type Step } from './session-fuzz.js'

export type Change =
  | Step
  | { patches: CellPatch[] }
  | { names: NameInput[] }
  | { styles: CellStyle[] }
  | { layout: LayoutEdit[] }

/** What a test does next: a change, or `undo` and `redo` some times. */
export type Move = { change: Change } | { undo: number } | { redo: number }

const STYLES: CellStyle[] = [{ font: { bold: true } }, { fill: { color: '#FFFF00' } }, { font: { italic: true } }]

/** The book of `fuzz` with a table and a SUBTOTAL on `T`, names used by formulas, and a style table. */
export function undoBook(seed: number): BookInput {
  const b = fuzz(seed).book
  const t = b.sheets[2]
  Object.assign(t.cells, {
    A4: { formula: 'SUBTOTAL(9,S!A1:A60)+SUM(n1)' },
    A5: { formula: 'SUM(Tab[Qty])+COUNTA(Tab[Price])+n2' },
    C1: { value: 'Qty' }, D1: { value: 'Price' },
    C2: { value: 3 }, D2: { value: 1.5 }, C3: { value: 4 }, D3: { value: 2.5 },
    E2: { formula: 'Tab[@Qty]*Tab[@Price]' }, E3: { formula: 'Tab[@Qty]*Tab[@Price]' },
  })
  t.tables = [{ name: 'Tab', range: 'C1:D3', columns: [{ name: 'Qty' }, { name: 'Price' }] }]
  // lines from the start: a change of them changes some of their entries
  Object.assign(b.sheets[0], { rows: { '2:4': { height: 20 }, '10': { hidden: true }, '30:40': { format: '0.0' } }, columns: { 'A:B': { width: 12 }, H: { format: '0.00' } } })
  return { ...b, styles: STYLES.slice(), names: [{ name: 'n1', ref: 'S!$A$1:$A$10' }, { name: 'n2', ref: '7' }] }
}

/** The moves of a seed: changes, and undos and redos between them. */
export function undoMoves(seed: number): () => Move {
  const steps = fuzz(seed).step
  const patch = patches(seed)
  const rnd = mulberry32(seed * 104729 + 3)
  const int = (n: number) => Math.floor(rnd() * n)
  const pick = <T>(xs: readonly T[]): T => xs[int(xs.length)]
  const sheet = () => pick(['S', 'D', 'T', 'S', 'S'])
  const span = (axis: 'r' | 'c') => {
    const a = axis === 'r' ? 1 + int(80) : int(12)
    const b = a + (rnd() < 0.5 ? 0 : 1 + int(4))
    const t = (i: number) => (axis === 'r' ? String(i) : indexToCol(i))
    return a === b ? t(a) : `${t(a)}:${t(b)}`
  }
  const layout = (): LayoutEdit => {
    const s = sheet()
    switch (int(9)) {
      case 0:
        return { sheet: s, rows: span('r'), set: pick([{ height: 10 + int(30) }, { hidden: true }, { hidden: null }, { style: int(3) }, { format: '0.0' }, { height: null, format: null }]) }
      case 1:
        return { sheet: s, columns: span('c'), set: pick([{ width: 2 + int(20) }, { format: '0.000' }, { style: int(3), hidden: true }, { width: null, hidden: null }]) }
      case 2:
        // merged areas over spills (S!E, F, G, I) stop them spilling
        return { sheet: 'S', merges: rnd() < 0.3 ? [] : [`${pick(['E', 'F', 'G', 'I', 'K'])}${2 + int(40)}:${pick(['F', 'G', 'K'])}${45 + int(5)}`] }
      case 3:
        return { sheet: s, conditionalFormats: rnd() < 0.3 ? [] : [{ type: 'cellIs', ranges: ['A1:B300'], operator: 'greaterThan', formulas: [String(int(40))], style: { font: { bold: true } } }] }
      case 4:
        return { sheet: s, frozen: rnd() < 0.5 ? null : { rows: 1 + int(3) } }
      case 5:
        return { sheet: s, validations: rnd() < 0.4 ? [] : [{ range: 'B1:B20', type: 'whole', operator: 'between', formula1: '0', formula2: '50' }] }
      case 6:
        return { sheet: s, filtered: rnd() < 0.5 }
      case 7:
        // AutoFit: the widths of the widest cells, bestFit, a hidden column shown
        return { sheet: s, columns: span('c'), fit: true }
      default:
        return { sheet: s, outline: rnd() < 0.5 ? null : { summaryBelow: false } }
    }
  }
  const op = (): StructureOp =>
    pick<() => StructureOp>([
      () => ({ op: 'insertColumns', sheet: pick(['D', 'T']), column: indexToCol(int(8)), count: 1 + int(2) }),
      () => ({ op: 'deleteColumns', sheet: pick(['D', 'S']), column: indexToCol(9 + int(6)), count: 1 }),
      () => ({ op: 'moveRange', from: { sheet: 'S', range: `B${300 + int(50)}:C${360 + int(40)}` }, to: { sheet: pick(['S', 'D']), cell: `K${500 + int(100)}` } }),
      () => ({ op: 'moveRange', from: { sheet: 'S', range: `A${1 + int(20)}:A${25 + int(20)}` }, to: { sheet: 'S', cell: `A${3 + int(30)}` } }),
      () => ({ op: 'copyRange', from: { sheet: 'S', range: `A${1 + int(50)}:B${60 + int(10)}` }, to: { sheet: pick(['S', 'D']), range: `M${1 + int(30)}` } }),
      () => ({ op: 'copyRange', from: { sheet: 'S', range: 'H1' }, to: { sheet: 'S', range: `L1:L${2 + int(10)}` } }),
      () => ({ op: 'renameSheet', sheet: pick(['T', 'U']), name: pick(['U', 'T', 'Тест']) }),
      () => ({ op: 'insertSheet', name: pick(['N', 'M']), before: pick(['S', 'D', undefined]) }) as StructureOp,
      () => ({ op: 'deleteSheet', sheet: pick(['N', 'M', 'N', 'D']) }),
      () => ({ op: 'moveSheet', sheet: pick(['S', 'D', 'T', 'N']), before: pick(['S', 'T', undefined]) }) as StructureOp,
      () => ({ op: 'hideSheet', sheet: pick(['D', 'T', 'N']), very: rnd() < 0.3 }),
      () => ({ op: 'unhideSheet', sheet: pick(['D', 'T', 'N']) }),
      () => ({ op: 'renameTable', table: pick(['Tab', 'Tab2']), name: pick(['Tab2', 'Tab']) }),
      () => ({ op: 'insertRows', sheet: 'T', row: 2 + int(3), count: 1 }),
      () => ({ op: 'deleteRows', sheet: 'T', row: 2 + int(4), count: 1 }),
      () => ({ op: 'convertTable', table: pick(['Tab', 'Tab2']) }),
      () => ({ op: 'addTable', sheet: 'T', table: { name: 'Tab', range: 'C1:D3', columns: [{ name: 'Qty' }, { name: 'Price' }] } }),
      // a table over empty cells: its header cells take the columns' names
      () => ({ op: 'addTable', sheet: pick(['T', 'S']), table: { name: 'New', range: `G${20 + int(5)}:H${30 + int(5)}`, columns: [{ name: 'A' }, { name: 'B' }] } }),
      () => ({ op: 'deleteTable', table: pick(['New', 'Tab']) }),
    ])()
  const change = (): Change => {
    const x = rnd()
    if (x < 0.4) return steps()
    if (x < 0.55) return { patches: patch() }
    if (x < 0.6) return { edits: [{ sheet: 'T', cell: pick(['C1', 'D1']), input: rnd() < 0.2 ? null : { value: pick(['Qty', 'Price', 'Cost', 'Qty', 5]) } }] }
    if (x < 0.65) return { names: [{ name: 'n1', ref: pick(['S!$A$1:$A$10', 'S!$A$5:$A$40', 'D!$A$1:$A$9']) }, ...(rnd() < 0.7 ? [{ name: 'n2', ref: String(int(9)) }] : [])] }
    if (x < 0.68) return { styles: [...STYLES, ...(rnd() < 0.5 ? [{ font: { size: 8 + int(10) } }] : [])] }
    if (x < 0.82) return { layout: Array.from({ length: 1 + (rnd() < 0.3 ? int(3) : 0) }, layout) }
    return { op: op() }
  }
  // after an undo, a redo half the time
  let undid = false
  return () => {
    const x = rnd()
    const last = undid
    undid = false
    if (last && x < 0.5) return { redo: 1 + (rnd() < 0.3 ? int(3) : 0) }
    if (x < 0.25) {
      undid = true
      return { undo: 1 + (rnd() < 0.3 ? int(4) : 0) }
    }
    return { change: change() }
  }
}
