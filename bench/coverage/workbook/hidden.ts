/**
 * SUBTOTAL and AGGREGATE over hidden rows and columns: rows hidden in the file by hand, by an autofilter (rows that
 * fail its criteria, rows that pass them hidden too, rows outside its range), by Excel's own filter and hiding
 * after opening; columns hidden under a row range and a two-dimensional one; errors in hidden rows. The data are
 * powers of two, so a sum names the rows it counted.
 */
import type { CellInput, SheetInput } from '../../../src/contract.js'
import type { Step, WorkbookCase } from './cases.js'
import { formulaCell as f } from '../stored.js'

const AGG = 'AGGREGATE'

/** Sheet S: a header in row 1, A2:A16 the values 2^0 … 2^14, B2:B16 the keys a filter reads (`x` kept, `y` not). */
const KEYS = 'xxxxxyyxxyxxxxx'
function dataSheet(err = false, name = 'S'): SheetInput {
  const cells: Record<string, CellInput> = { A1: { value: 'v' }, B1: { value: 'k' } }
  for (let i = 0; i < 15; i++) {
    cells[`A${i + 2}`] = { value: 2 ** i }
    cells[`B${i + 2}`] = { value: KEYS[i] }
  }
  // an error in row 4 (hidden by hand in the cases with errors) and one in row 12 (visible)
  if (err) Object.assign(cells, { A4: f('1/0'), A12: f('NA()') })
  return { name, cells }
}

/** Sheet H: B2:M2 and B3:M3 the values 2^0 … 2^11 and 2^12 … 2^23, read across hidden columns and a hidden row. */
function wideSheet(): SheetInput {
  const cells: Record<string, CellInput> = {}
  for (let c = 0; c < 12; c++) {
    const col = String.fromCharCode(66 + c)
    cells[`${col}2`] = { value: 2 ** c }
    cells[`${col}3`] = { value: 2 ** (c + 12) }
  }
  return { name: 'H', cells }
}

const K: Record<number, string> = { 14: '2', 15: '2', 16: '0.3', 17: '1', 18: '0.3', 19: '1' }

/** Every SUBTOTAL code and AGGREGATE function and option over `r`, and the array forms of AGGREGATE over `r*1`. */
function formulas(r: string): string[] {
  const out: string[] = []
  for (const c of [...Array.from({ length: 11 }, (_, i) => i + 1), ...Array.from({ length: 11 }, (_, i) => i + 101)]) out.push(`SUBTOTAL(${c},${r})`)
  for (let fn = 1; fn <= 13; fn++) for (let o = 0; o <= 7; o++) out.push(`${AGG}(${fn},${o},${r})`)
  for (let fn = 14; fn <= 19; fn++) for (let o = 0; o <= 7; o++) out.push(`${AGG}(${fn},${o},${r},${K[fn]})`)
  for (const o of [0, 1, 5, 6, 7]) out.push(`${AGG}(14,${o},${r}*1,1)`, `${AGG}(15,${o},${r}*1,2)`)
  return out
}

/** Sheet F: the formulas down column A. */
const formulaSheet = (fs: string[]): SheetInput => ({ name: 'F', cells: Object.fromEntries(fs.map((x, i) => [`A${i + 1}`, f(x)])) })

const MAIN = [...formulas('S!A2:A16'), 'SUBTOTAL(9,S!A2:A5,S!A8:A16)', 'SUBTOTAL(109,S!A2:A5,S!A8:A16)', 'SUBTOTAL(109,S!A:A)', `${AGG}(9,5,S!A:A)`]

const WIDE = [
  ...['SUBTOTAL(9,H!B2:M2)', 'SUBTOTAL(109,H!B2:M2)', 'SUBTOTAL(9,H!B2:M3)', 'SUBTOTAL(109,H!B2:M3)', 'SUBTOTAL(103,H!B2:M3)'],
  ...Array.from({ length: 8 }, (_, o) => [`${AGG}(9,${o},H!B2:M2)`, `${AGG}(9,${o},H!B2:M3)`]).flat(),
  `${AGG}(14,5,H!B2:M3,1)`,
  `${AGG}(15,5,H!B2:M3,1)`,
]

const read = (n: number) => [{ sheet: 'F', range: `A1:A${n}` }]

/** A case over S (and H) with the formulas of F. */
function bookCase(id: string, about: string, extra: Partial<WorkbookCase> = {}, err = false, wide = false): WorkbookCase {
  const fs = wide ? [...MAIN, ...WIDE] : MAIN
  return { id, about, book: { sheets: [dataSheet(err), ...(wide ? [wideSheet()] : []), formulaSheet(fs)] }, read: read(fs.length), ...extra }
}

const FILTER = { range: 'A1:B16', keep: { 1: ['x'] } }

export function buildWorkbookCases(): WorkbookCase[] {
  const recalc: Step[] = ['read', 'calc', 'read', 'full', 'read']
  return [
    bookCase('none', 'nothing hidden', {}, false, true),
    bookCase('manual', 'rows 3:4 and 16 hidden by hand; columns D:E and row 3 of H hidden', { file: { sheets: { S: { hiddenRows: ['3:4', '16'] }, H: { hiddenColumns: ['D:E'], hiddenRows: ['3'] } } } }, false, true),
    bookCase('filter', 'an autofilter keeping x hides rows 7, 8 and 11', { file: { sheets: { S: { filter: FILTER, hiddenRows: ['7:8', '11'] } } } }),
    bookCase('filter-manual', 'the autofilter over A1:B12, row 4 (kept by it) and row 14 (outside it) hidden too', { file: { sheets: { S: { filter: { ...FILTER, range: 'A1:B12' }, hiddenRows: ['4', '7:8', '11', '14'] } } } }),
    bookCase('filter-outside', 'an autofilter over A5:B12 keeping x, rows 3 (above it), 7:8 (failing it) and 15 (below it) hidden', {
      file: { sheets: { S: { filter: { range: 'A5:B12', keep: { 1: ['x'] } }, hiddenRows: ['3', '7:8', '15'] } } },
    }),
    bookCase('filter-nomode', 'the autofilter keeping x without the sheet\'s filter mode, rows 7, 8 and 11 hidden', { file: { sheets: { S: { filter: { ...FILTER, mode: false }, hiddenRows: ['7:8', '11'] } } } }),
    {
      id: 'filter-other',
      about: 'S filtered (rows 7, 8 and 11), rows 3:4 of T hidden by hand: the formulas read T',
      book: { sheets: [dataSheet(), dataSheet(false, 'T'), formulaSheet(MAIN.map((x) => x.replaceAll('S!', 'T!')))] },
      file: { sheets: { S: { filter: FILTER, hiddenRows: ['7:8', '11'] }, T: { hiddenRows: ['3:4'] } } },
      read: read(MAIN.length),
    },
    bookCase('excel-filter-outside', 'Excel filters A5:B12 to x, then rows 3 and 15 are hidden by hand', {
      steps: ['full', { filter: { sheet: 'S', range: 'A5:B12', field: 2, criterion: 'x' } }, 'read', { hide: { sheet: 'S', rows: '3', hidden: true } }, { hide: { sheet: 'S', rows: '15', hidden: true } }, 'read'],
    }),
    bookCase('filter-blank', 'an autofilter without criteria over rows 7, 8 and 11 hidden', { file: { sheets: { S: { filter: { range: 'A1:B16', keep: {} }, hiddenRows: ['7:8', '11'] } } } }),
    bookCase('manual-err', 'rows 3:4 hidden by hand, an error in row 4 and one in row 12', { file: { sheets: { S: { hiddenRows: ['3:4'] } } } }, true),
    bookCase('excel-filter', 'Excel filters to x after opening, then rows 3:4 are hidden by hand', {
      steps: ['full', 'read', { filter: { sheet: 'S', range: 'A1:B16', field: 2, criterion: 'x' } }, ...recalc, { hide: { sheet: 'S', rows: '3:4', hidden: true } }, ...recalc],
    }),
    bookCase('hide-after', 'rows 3:4 hidden after opening, then shown again', {
      steps: ['full', 'read', { hide: { sheet: 'S', rows: '3:4', hidden: true } }, ...recalc, { hide: { sheet: 'S', rows: '3:4', hidden: false } }, 'read', 'calc', 'read'],
    }),
  ]
}
