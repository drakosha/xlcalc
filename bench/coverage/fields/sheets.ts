/**
 * Tab colours, gridlines and headings: written by `writeXlsx` on sheets of each kind, and by `writeKept` set, changed
 * and removed over a file.
 */
import type { BookInput } from '../../../src/contract.js'
import type { FieldsCase } from './cases.js'

const cells = { A1: { value: 1 } }

const four: BookInput = {
  sheets: [
    { name: 'Red', cells, tabColor: '#FF0000', view: { gridlines: false } },
    { name: 'Blue', cells, tabColor: '#0070C0', view: { headings: false, zoom: 150 } },
    { name: 'Both', cells, view: { gridlines: false, headings: false }, outline: { summaryBelow: false } },
    { name: 'Plain', cells },
  ],
}

export function buildFieldsCases(): FieldsCase[] {
  return [
    { id: 'written', book: four },
    { id: 'kept-set', book: { sheets: [{ name: 'S', cells }, { name: 'T', cells, outline: { summaryRight: false } }] }, kept: [{ sheet: 'S', tabColor: '#00B050' }, { sheet: 'S', view: { gridlines: false } }, { sheet: 'T', tabColor: '#7030A0' }, { sheet: 'T', view: { headings: false } }] },
    { id: 'kept-changed', book: four, kept: [{ sheet: 'Red', tabColor: '#FFC000' }, { sheet: 'Red', view: { gridlines: true } }, { sheet: 'Blue', tabColor: null }, { sheet: 'Both', view: null }] },
  ]
}
