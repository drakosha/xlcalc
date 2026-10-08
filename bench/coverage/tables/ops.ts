/**
 * Structure operations over tables: rows and columns inserted and deleted inside, at the edges of and across tables,
 * tables renamed, converted and deleted, sheets renamed and deleted, header cells typed over, cells copied and moved.
 */
import type { StructureOp } from '../../../src/contract.js'
import { book, on, sales, table, type CellSpec, type TableCase, type TableStep } from './cases.js'

/** Formulas over Sales on S (column J) and on Other. */
const OVER: Record<string, CellSpec> = {
  J1: '=SUM(Sales[Qty])',
  J2: '=SUM(Sales[[Region]:[Price]])',
  J3: '=ROWS(Sales[#All])',
  J4: '=Sales[@Price]',
  J5: '=SUM(Sales[[#Totals],[Qty]])',
  J6: '=COLUMNS(Sales[#All])',
  J7: '=SUM(Sales[Price])',
  J8: '=SUM(B2:B5)',
  J9: '=Sales[[#Headers],[Qty]:[Total]]',
}
const OTHER: Record<string, CellSpec> = { A1: '=SUM(Sales[Total])', A2: '=SUM(Sales[[#All],[Qty]:[Price]])', A3: '=Sales[@Qty]', B1: '=ROWS(Sales)' }

const READ = [...on('S', 'A1:N12'), ...on('Other', 'A1:B5')]

const op = (o: StructureOp): TableStep => ({ op: o })
const edit = (cell: string, value: number | string | null | { formula: string }): TableStep => ({ edit: { sheet: 'S', cell, value } })

/** Sales with a totals row: a label, the sum of Qty, the average price. */
function salesTotals() {
  const t = sales()
  t.table.range = 'A1:D6'
  t.table.totals = true
  t.table.columns[0].totals = { label: 'Total' }
  t.table.columns[1].totals = { function: 'sum' }
  t.table.columns[2].totals = { function: 'average' }
  Object.assign(t.cells, { A6: { value: 'Total' }, B6: { formula: 'SUBTOTAL(109,Sales[Qty])' }, C6: { formula: 'SUBTOTAL(101,Sales[Price])' } })
  return t
}

const side = () => table('Side', 'F1', ['a', 'b'], [[1, 2], [3, 4], [5, 6]])
const below = () => table('Below', 'A8', ['x', 'y'], [[1, 2], [3, 4]])

export function buildTableCases(): TableCase[] {
  const cases: TableCase[] = []
  const add = (id: string, about: string, steps: TableStep[], b = book(['S', OVER, sales()], ['Other', OTHER]), read = READ) => cases.push({ id, about, book: b, steps, read })
  const rows = (o: 'insertRows' | 'deleteRows', row: number, count?: number) => op({ op: o, sheet: 'S', row, ...(count ? { count } : {}) })
  const cols = (o: 'insertColumns' | 'deleteColumns', column: string, count?: number) => op({ op: o, sheet: 'S', column, ...(count ? { count } : {}) })

  add('ins-inside', 'a row inserted inside the data: the table grows, the calculated column fills it', [rows('insertRows', 3)])
  add('ins-first-data', 'a row inserted at the first data row: the table grows', [rows('insertRows', 2)])
  add('ins-header', 'a row inserted at the header row: the table moves down', [rows('insertRows', 1)])
  add('ins-after', 'a row inserted right below the table: it does not grow', [rows('insertRows', 6)])
  add('ins-count', 'three rows inserted inside', [rows('insertRows', 4, 3)])
  add('ins-totals', 'a row inserted at the totals row: the table grows above it', [rows('insertRows', 6)], book(['S', OVER, salesTotals()], ['Other', OTHER]))
  add('ins-no-header', 'a row inserted at the first row of a table without a header row', [rows('insertRows', 2)], book(['S', { J1: '=SUM(T[b])', J2: '=ROWS(T)' }, table('T', 'A2', ['a', 'b'], [[1, 2], [3, 4]], { headers: false })]), on('S', 'A1:J6'))
  add('ins-two', 'a row inserted into two tables side by side: Excel refuses', [rows('insertRows', 3)], book(['S', OVER, sales(), side()], ['Other', OTHER]))
  add('ins-grow-move', 'a row inserted where one table grows and another, beside it, has its header row', [rows('insertRows', 3)], book(['S', OVER, sales(), table('Side', 'F3', ['a', 'b'], [[1, 2], [3, 4]])], ['Other', OTHER]))
  add('ins-below', 'a row inserted into a table below another, past the other', [rows('insertRows', 9)], book(['S', OVER, sales(), below()], ['Other', OTHER]))

  add('del-inside', 'a data row deleted: the table shrinks', [rows('deleteRows', 3)])
  add('del-all-data', 'every data row deleted: the table keeps one row, from below', [rows('deleteRows', 2, 4)])
  add('del-header', "the header row deleted alone: Excel refuses", [rows('deleteRows', 1)])
  add('del-header-data', 'the header row and a data row deleted: Excel refuses', [rows('deleteRows', 1, 2)])
  add('del-whole', 'the rows of the whole table deleted: references to it become #REF!', [rows('deleteRows', 1, 5)])
  add('del-whole-more', 'the whole table and a row below it deleted', [rows('deleteRows', 1, 6)])
  add('del-last', 'the last data row and the row below it deleted', [rows('deleteRows', 5, 2)])
  add('del-totals', 'the totals row deleted alone', [rows('deleteRows', 6)], book(['S', OVER, salesTotals()], ['Other', OTHER]))
  add('del-data-totals', 'the last data row and the totals row deleted', [rows('deleteRows', 5, 2)], book(['S', OVER, salesTotals()], ['Other', OTHER]))
  add('del-two', 'a row deleted from two tables side by side: Excel refuses', [rows('deleteRows', 3)], book(['S', OVER, sales(), side()], ['Other', OTHER]))
  add('del-one-whole-other', 'rows deleting a table below another and nothing of the other', [rows('deleteRows', 8, 3)], book(['S', OVER, sales(), below()], ['Other', OTHER]))

  add('insc-inside', 'a column inserted inside: a new column Column1, ranges over it grow', [cols('insertColumns', 'B')])
  add('insc-first', 'a column inserted at the first column: the table moves', [cols('insertColumns', 'A')])
  add('insc-after', 'a column inserted right of the table: it does not grow', [cols('insertColumns', 'E')])
  add('insc-two', 'two columns inserted before the last one', [cols('insertColumns', 'D', 2)])
  const named = table('Sales', 'A1', ['Region', 'Qty', 'Column1', 'Total'], [['North', 2, 10, 20], ['South', 3, 20, 60]])
  add('insc-taken', 'a column inserted where Column1 is taken: Column2', [cols('insertColumns', 'B')], book(['S', { J1: '=SUM(Sales[Column1])' }, named]), on('S', 'A1:J6'))
  add('insc-stacked', 'a column inserted into two tables one above the other: Excel refuses', [cols('insertColumns', 'B')], book(['S', OVER, sales(), below()], ['Other', OTHER]))
  add('delc-middle', 'a column deleted: references to it become #REF!, ranges over it shrink', [cols('deleteColumns', 'C')])
  add('delc-first', 'the first column deleted: a range starting there becomes #REF!', [cols('deleteColumns', 'A')])
  add('delc-last', 'the last column, the calculated one, deleted', [cols('deleteColumns', 'D')])
  add('delc-two', 'two columns deleted, the ends of a range among them', [cols('deleteColumns', 'B', 2)])
  add('delc-all', 'every column deleted: references to the table become #REF!', [cols('deleteColumns', 'A', 4)])
  add('delc-stacked', 'a column deleted from two tables one above the other: Excel refuses', [cols('deleteColumns', 'B')], book(['S', OVER, sales(), below()], ['Other', OTHER]))

  add('rename', 'the table renamed: every reference follows', [op({ op: 'renameTable', table: 'Sales', name: 'Ventes' })])
  add('rename-case', 'the table renamed to another case', [op({ op: 'renameTable', table: 'Sales', name: 'SALES' })])
  add('convert', 'the table converted to a range: references become references to its cells', [op({ op: 'convertTable', table: 'Sales' })])
  add('convert-totals', 'a table with a totals row converted', [op({ op: 'convertTable', table: 'Sales' })], book(['S', OVER, salesTotals()], ['Other', OTHER]))
  add('delete', 'the table deleted with its cells: references become #REF!', [op({ op: 'deleteTable', table: 'Sales' })])
  add('sheet-rename', 'the sheet of the table renamed: references do not change', [op({ op: 'renameSheet', sheet: 'S', name: 'Data' })], book(['S', OVER, sales()], ['Other', OTHER]), [...on('Data', 'A1:N12'), ...on('Other', 'A1:B5')])
  add('sheet-delete', 'the sheet of the table deleted: references from another sheet', [op({ op: 'deleteSheet', sheet: 'S' })], book(['S', OVER, sales()], ['Other', OTHER]), on('Other', 'A1:B5'))

  add('head-rename', 'a header cell typed over: the column is renamed and references follow', [edit('B1', 'Quantity')])
  add('head-number', 'a number typed into a header cell', [edit('C1', 2024)])
  add('head-clear', 'a header cell cleared: Column1', [edit('A1', null)])
  add('head-dup-left', 'a header given the name of a column to its right: that column becomes Price2', [edit('B1', 'Price')])
  add('head-dup-right', 'a header given the name of a column to its left, in another case', [edit('D1', 'price')])
  add('head-case', 'a header typed in another case', [edit('B1', 'QTY')])
  add('head-two', 'two headers cleared one after the other', [edit('A1', null), edit('C1', null)])

  add('copy-out', 'a formula of the calculated column copied out of the table, and one over the table to another sheet', [
    op({ op: 'copyRange', from: { sheet: 'S', range: 'D2' }, to: { sheet: 'S', range: 'F2' } }),
    op({ op: 'copyRange', from: { sheet: 'S', range: 'J1:J4' }, to: { sheet: 'Other', range: 'C1' } }),
  ], book(['S', OVER, sales()], ['Other', OTHER]), [...READ, ...on('Other', 'C1:C5')])
  add('move-table', 'the whole table cut and pasted elsewhere: the table moves', [op({ op: 'moveRange', from: { sheet: 'S', range: 'A1:D5' }, to: { sheet: 'S', cell: 'A20' } })], book(['S', OVER, sales()], ['Other', OTHER]), [...on('S', 'A1:N24'), ...on('Other', 'A1:B5')])
  add('move-out', 'a calculated cell cut out of the table', [op({ op: 'moveRange', from: { sheet: 'S', range: 'D3' }, to: { sheet: 'S', cell: 'F3' } })])
  add('move-in', 'a cell cut into the table', [op({ op: 'moveRange', from: { sheet: 'S', range: 'J8' }, to: { sheet: 'S', cell: 'C3' } })])
  return cases
}
