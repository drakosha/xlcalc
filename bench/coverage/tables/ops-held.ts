/**
 * Held out: structure operations over tables that the set `ops` does not have, predicted by the engine before Excel
 * took them (`ops-held.predicted.json`, from `predict.ts`): inserts beside another table's edges, the only data row
 * deleted, every data row with the totals row, pastes over header rows, moves between tables, renames Excel refuses,
 * headers typed as booleans and decimals, a conversion seen through names.
 */
import type { StructureOp } from '../../../src/contract.js'
import { book, on, sales, table, type CellSpec, type TableCase, type TableStep } from './cases.js'

const OVER: Record<string, CellSpec> = { J1: '=SUM(Sales[Qty])', J2: '=SUM(Sales[[Region]:[Price]])', J3: '=ROWS(Sales[#All])', J4: '=Sales[@Price]', J5: '=SUM(B2:C5)', J6: '=SUM(A1:D5)' }
const READ = on('S', 'A1:N12')
const op = (o: StructureOp): TableStep => ({ op: o })
const edit = (cell: string, value: number | string | null): TableStep => ({ edit: { sheet: 'S', cell, value } })

export function buildTableCases(): TableCase[] {
  const cases: TableCase[] = []
  const add = (id: string, about: string, steps: TableStep[], b = book(['S', OVER, sales()]), read = READ) => cases.push({ id, about, book: b, steps, read })
  const rows = (o: 'insertRows' | 'deleteRows', row: number, count?: number) => op({ op: o, sheet: 'S', row, ...(count ? { count } : {}) })
  const cols = (o: 'insertColumns' | 'deleteColumns', column: string, count?: number) => op({ op: o, sheet: 'S', column, ...(count ? { count } : {}) })

  add('ins-other-header', "a row inserted at another table's header row, below this one's end", [rows('insertRows', 6)], book(['S', OVER, sales(), table('Side', 'F6', ['a', 'b'], [[1, 2], [3, 4]])]))
  add('ins-past-other', 'a row inserted into a table past the end of a table beside it', [rows('insertRows', 4)], book(['S', OVER, sales(), table('Side', 'F1', ['a', 'b'], [[1, 2], [3, 4]])]))
  add('ins-other-last', "a row inserted at another table's last row, inside both", [rows('insertRows', 3)], book(['S', OVER, sales(), table('Side', 'F1', ['a', 'b'], [[1, 2], [3, 4]])]))
  add('del-only-row', 'the only data row deleted', [rows('deleteRows', 2)], book(['S', { J1: '=SUM(T[b])', J2: '=ROWS(T)', K2: 'outside' }, table('T', 'A1', ['a', 'b'], [[1, '=[@a]*2']], { columns: [{}, { formula: 'T[[#This Row],[a]]*2' }] })]), on('S', 'A1:K4'))
  const totals = sales()
  totals.table.range = 'A1:D6'
  totals.table.totals = true
  totals.table.columns[1].totals = { function: 'sum' }
  totals.cells.B6 = { formula: 'SUBTOTAL(109,Sales[Qty])' }
  add('del-data-totals-all', 'every data row and the totals row deleted', [rows('deleteRows', 2, 5)], book(['S', OVER, totals]))
  add('del-first-rows', 'the first two data rows deleted, a range over the data shrinks', [rows('deleteRows', 2, 2)])
  add('insc-before-last', 'a column inserted before the last one: A1 references beside it as usual', [cols('insertColumns', 'D')])
  add('delc-middle-span', 'a column inside a span of columns deleted: the span stays', [cols('deleteColumns', 'B')])
  add('copy-header', "a copy pasted over a table's header cell", [op({ op: 'copyRange', from: { sheet: 'S', range: 'J1' }, to: { sheet: 'S', range: 'B1' } })])
  add('cut-header', "part of a table's header row cut and pasted elsewhere", [op({ op: 'moveRange', from: { sheet: 'S', range: 'A1:B1' }, to: { sheet: 'S', cell: 'F1' } })])
  add('copy-in-table', 'a calculated cell copied down inside its table', [op({ op: 'copyRange', from: { sheet: 'S', range: 'D2' }, to: { sheet: 'S', range: 'D3:D5' } })])
  add('move-to-table', "a calculated cell cut into another table's data", [op({ op: 'moveRange', from: { sheet: 'S', range: 'D3' }, to: { sheet: 'S', cell: 'G3' } })], book(['S', OVER, sales(), table('Side', 'F1', ['a', 'b'], [[1, 2], [3, 4], [5, 6]])]))
  add('rename-cell', 'a table renamed to a cell address: Excel refuses', [op({ op: 'renameTable', table: 'Sales', name: 'AB12' })])
  const named = book(['S', { ...OVER, J7: '=SUM(rate)' }, sales()])
  named.names = [{ name: 'rate', ref: "'S'!$C$2" }]
  add('rename-name', 'a table renamed to a defined name: Excel refuses', [op({ op: 'renameTable', table: 'Sales', name: 'Rate' })], named)
  add('head-bool', 'TRUE typed into a header cell (Excel reads it as the boolean)', [edit('B1', 'TRUE')])
  add('head-decimal', 'a decimal typed into a header cell', [edit('C1', 1.5)])
  add('head-dup-three', 'a name typed that two columns to its right have, one with a number', [edit('A1', 'Qty')], book(['S', { J1: '=SUM(Sales[Qty2])', J2: '=SUM(Sales[Qty])' }, table('Sales', 'A1', ['Region', 'Qty', 'Qty2'], [['a', 1, 2], ['b', 3, 4]])]))
  const conv = book(['S', OVER, sales()], ['Other', { A1: '=SUM(nt)', A3: '=nr', A4: '=nr*2', B2: '=ROWS(nt)' }])
  conv.names = [{ name: 'nt', ref: 'Sales' }, { name: 'nr', ref: 'Sales[[#This Row],[Price]]' }]
  add('convert-names', 'a table converted: names over it keep their values', [op({ op: 'convertTable', table: 'Sales' })], conv, [...READ, ...on('Other', 'A1:B4')])
  return cases
}
