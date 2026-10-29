/**
 * Windows split into panes without freezing them: written by `writeXlsx` with lines on both axes and on one, scrolled,
 * zoomed, with each pane active, and under frozen panes, which take its place; and by `writeKept` set, changed and
 * removed over a file.
 */
import type { BookInput, SheetInput } from '../../../src/contract.js'
import type { FieldsCase } from './cases.js'

const cells = { A1: { value: 1 } }

const sheets: SheetInput[] = [
  { name: 'Both', cells, view: { split: { x: 120, y: 60, topLeft: 'F20', activePane: 'bottomRight' }, activeCell: 'G21', selection: ['G21'] } },
  { name: 'Columns', cells, view: { split: { x: 150.5, topLeft: 'H1', activePane: 'topRight' }, activeCell: 'H2', selection: ['H2'] } },
  { name: 'Rows', cells, view: { split: { y: 75, topLeft: 'A30', activePane: 'bottomLeft' } } },
  { name: 'Scrolled', cells, view: { topLeft: 'C5', split: { x: 200, y: 100, topLeft: 'K40' }, activeCell: 'D6', selection: ['D6'] } },
  { name: 'Zoomed', cells, view: { zoom: 200, split: { x: 120, y: 60, topLeft: 'F20', activePane: 'topRight' } } },
  { name: 'Frozen', cells, frozen: { rows: 2 }, view: { split: { x: 120, y: 60, topLeft: 'F20' } } },
  { name: 'Plain', cells },
]

const written: BookInput = { sheets }

export function buildFieldsCases(): FieldsCase[] {
  return [
    { id: 'written', book: written },
    {
      id: 'kept-set',
      book: { sheets: [{ name: 'S', cells }, { name: 'T', cells, view: { zoom: 150 } }] },
      kept: { views: { S: { split: { x: 100, y: 50, topLeft: 'E12', activePane: 'bottomRight' }, activeCell: 'E12', selection: ['E12'] }, T: { zoom: 150, split: { y: 90, topLeft: 'A25' } } } },
    },
    {
      id: 'kept-changed',
      book: written,
      kept: { views: { Both: { split: { x: 240, y: 30, topLeft: 'J5' } }, Columns: null, Rows: { split: { y: 75, topLeft: 'A30', activePane: 'bottomLeft' }, activeCell: 'A31', selection: ['A31'] }, Frozen: null } },
    },
  ]
}
