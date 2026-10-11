/**
 * Held out, predicted by the engine before Excel took the steps (`paste-names-held.predicted.json`): a copy of a table's
 * cells pasted onto a target that is no multiple of them, where Excel pastes one copy at the target's top-left cell and
 * the engine also took the target's other cells as pasted (a formula there in another table got the copied table's
 * name); a row inserted into a table whose calculated column formula differs from its cells'; and defined names shaped
 * like a cell past the grid (`XFE1`, `A0`), like a cell of the grid (`R1`, `C5`, `XFD1`) and like R1C1 (`R1C1`), added
 * through the object model, with formulas typed after that read them.
 */
import { book, on, sales, table, type TableCase, type TableStep } from './cases.js'

const typed = (cell: string, formula: string): TableStep => ({ edit: { sheet: 'S', cell, value: { formula } } })
const added = (name: string, ref: string): TableStep => ({ name: { name, ref } })
const copy = (from: string, to: string): TableStep => ({ op: { op: 'copyRange', from: { sheet: 'S', range: from }, to: { sheet: 'S', range: to } } })

export function buildTableCases(): TableCase[] {
  const cases: TableCase[] = []
  const add = (id: string, about: string, b: TableCase['book'], steps: TableStep[], read = on('S', 'A1:H7')) => cases.push({ id, about, book: b, steps, read })
  // Side on F1:G6: a and b, b a calculated column of a times 2
  const side = () => table('Side', 'F1', ['a', 'b'], [[1, '=[@a]*2'], [3, '=[@a]*2'], [5, '=[@a]*2'], [7, '=[@a]*2'], [9, '=[@a]*2']], { columns: [{}, { formula: 'Side[[#This Row],[a]]*2' }] })
  add('paste-not-multiple', "two calculated cells of Sales pasted onto three of Side's: one copy, the third cell keeps its formula", book(['S', {}, sales(), side()]), [copy('D2:D3', 'G2:G4')])
  add('paste-multiple', "two calculated cells of Sales pasted onto four of Side's: two copies", book(['S', {}, sales(), side()]), [copy('D2:D3', 'G2:G5')])
  add('paste-not-multiple-cols', "a row of two cells of Sales pasted onto three cells of a row whose third is in Side's calculated column", book(['S', {}, sales(), side()]), [copy('C3:D3', 'E4:G4')])
  // T on A1:B4: the calculated column's formula times 2, its cells' times 3
  const differs = table('T', 'A1', ['a', 'b'], [[1, '=[@a]*3'], [2, '=[@a]*3'], [3, '=[@a]*3']], { columns: [{}, { formula: 'T[[#This Row],[a]]*2' }] })
  add('calc-differs-insert', "a row inserted into a table whose calculated column formula (times 2) differs from its cells' (times 3)", book(['S', { D1: '=SUM(T[b])' }, differs]), [{ op: { op: 'insertRows', sheet: 'S', row: 3 } }], on('S', 'A1:D6'))
  const names = book(['S', { A1: 100 }])
  add(
    'add-cell-shapes',
    'names shaped like a cell past the grid (XFE1, A0) or of it (R1, C5, XFD1, r1), like R1C1 (R1C1) and four letters (ZZZZ1), added and read',
    names,
    [added('XFE1', '1'), added('A0', '2'), added('R1', '3'), added('C5', '4'), added('XFD1', '5'), added('r1', '6'), added('R1C1', '7'), added('ZZZZ1', '8'), typed('C1', 'XFE1+0'), typed('C2', 'A0+0'), typed('C3', 'ZZZZ1+0')],
    on('S', 'A1:C3'),
  )
  return cases
}
