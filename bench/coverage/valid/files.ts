/**
 * Data validation and outlines as files hold them: books written by `writeXlsx` with every type of rule and every
 * setting, rules over whole columns, several ranges and other sheets, texts past Excel's limits, groups of rows and
 * columns nested and collapsed with their summaries above and left, and books written after structure operations in the
 * engine. Excel must load each book and read in it what the book holds, and the reader must read the file Excel saved as
 * Excel read it. Predicted before Excel ran it (`files.predicted.json`).
 */
import type { BookInput, CellInput, DataValidation } from '../../../src/contract.js'
import { editStructure } from '../../../src/index.js'
import type { ValidCase, ValidCases } from './cases.js'

/** Sheet S: a column of values under each rule (row 1 holds the column's title), the bounds and lists beside them. */
function rulesBook(): BookInput {
  const cells: Record<string, CellInput> = {
    H1: { value: 2 },
    M2: { value: 'apple' }, M3: { value: 2 }, M4: { value: true }, M6: { value: 'kiwi' },
    B2: { value: 'Yes' }, B3: { value: 'no' }, B4: { value: 'Perhaps' },
    C2: { value: 5 }, C3: { value: 5.5 }, C4: { value: 11 }, C5: { formula: '"5"' },
    D2: { value: 2.5 }, D3: { value: 2 }, D4: { value: 'abc' },
    E2: { value: 45700 }, E3: { value: 45000 }, E4: { value: 45658.5 },
    F2: { value: 0.25 }, F3: { value: 0.75 },
    G2: { value: 'abc' }, G3: { value: 'a' }, G4: { value: 12345 }, G5: { value: true },
    A2: { value: 1 }, A3: { value: 5 }, A4: { value: 9 },
    I2: { value: 2 }, I3: { value: 4 }, I4: { value: 3 },
    K2: { value: 'APPLE' }, K3: { value: 2 }, K4: { value: 'pear' }, K5: { value: 'kiwi' },
    L2: { value: 'apple' }, L3: { value: 'Kiwi' },
    N2: { value: 'b' }, N3: { value: 'z' },
    O2: { value: 'x' }, O9: { value: 'q' },
    P2: { value: 7 }, P4: { value: 0 },
  }
  for (const [col, title] of Object.entries({ A: 'n', B: 'list', C: 'whole', D: 'decimal', E: 'date', F: 'time', G: 'length', I: 'custom', J: 'any', K: 'range', L: 'name', N: 'sheet', O: 'column', P: 'areas' })) cells[`${col}1`] = { value: title }
  const validations: DataValidation[] = [
    { range: 'B2:B6', type: 'list', formula1: '"Yes,No,Maybe"', promptTitle: 'Answer', prompt: 'Yes, No or Maybe' },
    { range: 'C2:C6', type: 'whole', formula1: '1', formula2: '10', errorStyle: 'warning', errorTitle: 'Whole', error: 'From 1 to 10' },
    { range: 'D2:D6', type: 'decimal', operator: 'greaterThan', formula1: '$H$1', allowBlank: false },
    { range: 'E2:E6', type: 'date', formula1: 'DATE(2025,1,1)', formula2: 'DATE(2025,12,31)', errorStyle: 'information' },
    { range: 'F2:F6', type: 'time', operator: 'lessThan', formula1: '0.5', showPrompt: false },
    { range: 'G2:G6', type: 'textLength', operator: 'between', formula1: '2', formula2: '4', showError: false },
    { range: 'I2:I6', type: 'custom', formula1: 'I2>A2', error: 'Above the value on the left\nsecond line' },
    { range: 'J2:J6', type: 'any', promptTitle: 'Note', prompt: 'Anything goes' },
    { range: 'K2:K6', type: 'list', formula1: '$M$2:$M$6', dropdown: false },
    { range: 'L2:L3', type: 'list', formula1: 'Fruits' },
    { range: 'N2:N4', type: 'list', formula1: "'Lists'!$A$1:$A$3" },
    { range: 'O:O', type: 'list', formula1: '"x,y"', allowBlank: false },
    { range: 'P2:P3 P5:P6', type: 'custom', formula1: 'P2>H$1' },
  ]
  return {
    sheets: [
      { name: 'S', cells, validations },
      { name: 'Lists', cells: { A1: { value: 'a' }, A2: { value: 'b' }, A3: { value: 'c' } } },
    ],
    names: [{ name: 'Fruits', ref: "'S'!$M$2:$M$6" }],
  }
}

/** Groups of rows nested and collapsed, groups of columns, summaries above and left on the second sheet. */
function outlineBook(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 16; r++) cells[`A${r}`] = { value: r }
  return {
    sheets: [
      {
        name: 'Below',
        cells,
        rows: { '3:4': { outlineLevel: 1 }, '5:6': { outlineLevel: 2 }, 7: { outlineLevel: 1 }, '10:12': { outlineLevel: 1, hidden: true }, 13: { collapsed: true } },
        columns: { 'C:D': { outlineLevel: 1 }, 'F:G': { outlineLevel: 1, hidden: true }, H: { collapsed: true } },
      },
      {
        name: 'Above',
        cells: { ...cells },
        rows: { 2: { collapsed: true }, '3:5': { outlineLevel: 1, hidden: true }, 8: { height: 30 }, '9:10': { outlineLevel: 3 } },
        columns: { B: { collapsed: true }, 'C:D': { outlineLevel: 2, hidden: true, width: 12 } },
        outline: { summaryBelow: false, summaryRight: false },
      },
    ],
  }
}

/** Texts past Excel's limits, and a literal list of 300 characters. */
function longBook(): BookInput {
  const long = (n: number) => 'abcdefghij'.repeat(n / 10)
  return {
    sheets: [
      {
        name: 'S',
        cells: { A1: { value: 'x' } },
        validations: [
          { range: 'A1:A3', type: 'any', promptTitle: long(40), prompt: long(300), errorTitle: long(40), error: long(300) },
          { range: 'B1:B3', type: 'list', formula1: `"${Array.from({ length: 60 }, (_, i) => `i${i}`).join(',')}"` },
        ],
      },
    ],
  }
}

const after = (book: BookInput, op: Parameters<typeof editStructure>[1]): BookInput => {
  const r = editStructure(book, op)
  if (!r.ok) throw new Error(`files: ${r.message}`)
  return r.book
}

export function buildValidCases(): ValidCases {
  const rules = rulesBook()
  const outline = outlineBook()
  const books: Record<string, BookInput> = {
    rules,
    outline,
    long: longBook(),
    'after-insert': after(rules, { op: 'insertRows', sheet: 'S', row: 3 }),
    'after-delete': after(rules, { op: 'deleteRows', sheet: 'S', row: 2 }),
    'after-move': after(rules, { op: 'moveRange', from: { sheet: 'S', range: 'I2:I3' }, to: { sheet: 'S', cell: 'I12' } }),
    'after-copy': after(rules, { op: 'copyRange', from: { sheet: 'S', range: 'B2:C3' }, to: { sheet: 'S', range: 'B12:C15' } }),
    'after-outline': after(outline, { op: 'insertRows', sheet: 'Below', row: 5 }),
  }
  const cases: ValidCase[] = Object.keys(books).map((book) => ({ id: book, book }))
  return { books, cases, window: 'A1:P20' }
}
