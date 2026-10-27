/**
 * The first look at Paste Special, before the engine has it: each kind of paste of one source onto a target that holds
 * cells, looks, a rule and a validation of its own, onto another sheet, tiled, transposed, over a merged area; and the
 * references of formulas pasted transposed. What was predicted before Excel ran it is in `probe.predicted.json`.
 */
import type { BookInput, CellInput, ConditionalFormat, DataValidation } from '../../../src/contract.js'
import { RED } from '../cond/cases.js'
import type { PasteCase, PasteCases, PasteWhat } from './cases.js'

/** S: the source A1:D3 and a target F1:I3 with its own; T: a target with its own; column A wide. */
function sourceBook(): BookInput {
  const cells: Record<string, CellInput> = {
    A1: { value: 1, style: 1, numberFormat: '0.00' },
    B1: { formula: 'A1*2', style: 2 },
    C1: { formula: '$A$1+A2' },
    D1: { value: 'link', link: { url: 'https://example.com/' } },
    A2: { value: 3, note: { text: 'n' } },
    B2: { formula: 'SEQUENCE(2)' },
    C2: { value: 'm' },
    A3: { value: 46037, numberFormat: 'm/d/yyyy' },
    C3: { formula: 'A9' },
    D3: { formula: 'NA()' },
    F1: { value: 'old', style: 3 },
    G2: { value: 9, numberFormat: '0%' },
    H3: { formula: 'G2+1' },
  }
  const rules: ConditionalFormat[] = [
    { type: 'cellIs', ranges: ['A1:B3'], operator: 'greaterThan', formulas: ['1'], style: RED },
    { type: 'expression', ranges: ['F1:I3'], formula: 'F1=9', style: { font: { italic: true } } },
  ]
  const validations: DataValidation[] = [
    { range: 'A1:A3', type: 'whole', operator: 'between', formula1: '0', formula2: '10' },
    { range: 'F1:F3', type: 'list', formula1: '"a,b"' },
  ]
  return {
    styles: [{}, { font: { bold: true }, fill: { color: '#FFFF00' } }, { font: { italic: true }, border: { bottom: { style: 'thin' } } }, { fill: { color: '#9BC2E6' } }],
    sheets: [
      { name: 'S', cells, merges: ['C2:D2'], conditionalFormats: rules, validations, columns: { A: { width: 20 } } },
      { name: 'T', cells: { B2: { value: 'own', style: 3 }, C3: { value: 5 } } },
    ],
  }
}

/** References pasted transposed: relative, mixed and absolute, inside the copy and outside it. */
function transposeBook(): BookInput {
  const cells: Record<string, CellInput> = {
    A1: { value: 1 },
    B1: { formula: 'A1*2' },
    C1: { formula: '$A1+B$1' },
    D1: { formula: 'A5+SUM(A1:B1)' },
    A2: { formula: 'B1+1' },
    B2: { formula: '$A$1+A$2' },
    C2: { value: 7 },
    D2: { formula: 'C2:C3*1' },
  }
  return { sheets: [{ name: 'S', cells }] }
}

/** A target merged over F1:G1, a source of two cells and one. */
function mergedBook(): BookInput {
  return {
    styles: [{}, { font: { bold: true } }],
    sheets: [{ name: 'S', cells: { A1: { value: 1, style: 1 }, B1: { value: 2 }, F1: { value: 'x' } }, merges: ['F1:G1'] }],
  }
}

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = []
  const add = (id: string, book: string, from: string, to: string, what: PasteWhat, transpose = false, toSheet = 'S') =>
    cases.push({ id, book, from: { sheet: 'S', range: from }, to: { sheet: toSheet, range: to }, what, ...(transpose ? { transpose } : {}) })
  for (const what of ['all', 'values', 'formulas', 'formats', 'validation', 'formulasAndNumberFormats', 'valuesAndNumberFormats', 'columnWidths', 'allExceptBorders'] as const) add(`${what}-F1`, 'source', 'A1:D3', 'F1', what)
  for (const what of ['values', 'formulas', 'formats'] as const) add(`${what}-T`, 'source', 'A1:D3', 'B2', what, false, 'T')
  add('formats-tiled', 'source', 'A1:B1', 'F10:I11', 'formats')
  add('values-tiled', 'source', 'A1', 'F10:G11', 'values')
  for (const what of ['all', 'values', 'formulas', 'formats'] as const) add(`${what}-transposed`, 'source', 'A1:D3', 'K1', what, true)
  add('refs-transposed', 'transpose', 'A1:D2', 'F1', 'formulas', true)
  add('refs-all-transposed', 'transpose', 'A1:D2', 'F1', 'all', true)
  add('merged-values', 'merged', 'A1:B1', 'F1', 'values')
  add('merged-formats', 'merged', 'A1:B1', 'F1', 'formats')
  add('merged-part-values', 'merged', 'A1', 'G1', 'values')
  add('merged-part-formats', 'merged', 'A1', 'G1', 'formats')
  return { books: { source: sourceBook(), transpose: transposeBook(), merged: mergedBook() }, cases }
}
