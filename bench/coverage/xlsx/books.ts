/**
 * Books for the xlsx round trip: the writer writes each one, Excel opens the file (a repair window is the
 * writer's fault), computes it and saves it as xlsx, the reader reads what Excel saved, and the book that comes
 * back is compared with the one written, Excel's stored values with the engine's (`compare.ts`). No real data:
 * the books are made up here, the synthetic book comes from its seeded generator, and the books of the structure
 * and spill sets are taken as they are.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import { synthBook } from '../../../src/synth.js'
import type { SpillSet } from '../spill/cases.js'
import type { StructSet } from '../struct/cases.js'
import type { TableSet } from '../tables/cases.js'
import { filterBook, Y_ROWS } from './filters.js'
import { formulaCell as f, takenBook, valueCell as v } from '../stored.js'

/** The book with an AutoFilter over the data of its sheet S. */
const withFilter = (b: BookInput): BookInput => ({ ...b, sheets: [{ ...b.sheets[0], autoFilter: { range: 'A1:B11' } }, ...b.sheets.slice(1)] })

/** The clock the round trip computes its books with. */
export const XLSX_CLOCK = { now: '2026-10-07T12:00:00Z', seed: 7 }

export interface RoundTrip {
  id: string
  what: string
  book: BookInput
  /** Written with formulas and their values (the default), or as values only. */
  mode?: 'formulas' | 'values'
  /** Volatile functions in it: Excel's values of them are its own. */
  volatile?: true
}

/** The sets of the structure and spill steps whose books the round trip takes. */
const STRUCT_SETS = ['arrays', 'cube', 'edge', 'move', 'moves', 'rows', 'sheets', 'spill', 'spill-more', 'styles']
const SPILL_SETS = ['at', 'base', 'cse', 'held', 'let', 'names', 'order']
const TABLE_SETS = ['refs', 'refs-held', 'ops', 'ops-held']

const values: BookInput = {
  sheets: [
    {
      name: 'Values',
      cells: {
        A1: v(0), A2: v(-1.5), A3: v(0.1 + 0.2), A4: v(1e-300), A5: v(1.7976931348623157e308), A6: v(123456789012345), A7: v(-0),
        B1: v('text'), B2: v('  leading and trailing  '), B3: v('line\nbreak'), B4: v('tab\there'), B5: v('cr\r\nlf'), B6: v('quotes " \' and & < >'),
        B7: v('Кириллица и ёлка'), B8: v('emoji 😀 and 中文'), B9: v('=not a formula'), B10: v("'apostrophe first"), B11: v(''),
        C1: v(true), C2: v(false), D1: v({ date: '2025-01-31' }), D2: v({ date: '1900-02-28' }), D3: v({ date: '2025-06-30T18:30:00' }),
        E1: { value: 1, numberFormat: '0.00' }, E2: { value: 0.25, numberFormat: '0%' }, E3: { value: 45688, numberFormat: 'dd.mm.yyyy' }, E4: { value: 45688.75, numberFormat: 'm/d/yyyy h:mm' },
        E5: { value: 1234.5, numberFormat: '#,##0.00 "₽"' }, E6: { value: -5, numberFormat: '0;[Red]-0' }, E7: { value: 'x', numberFormat: '@' },
      },
    },
  ],
}

/** Text that XML cannot hold as it is, and text that reads like Excel's own escapes. */
const escapes: BookInput = {
  sheets: [
    {
      name: 'Escapes',
      cells: {
        A1: v('_x0041_'), A2: v('a_x000D_b'), A3: v('_x005F_x0041_'), A4: v('vertical\u000btab'), A5: v('nul\u0001and\u001fother'), A6: v('￾'),
        A7: v('x'.repeat(32_767)), A8: v('lone \uD800 surrogate'),
        B1: f('"_x0041_"&"!"'), B2: f('LEN(A4)'), B3: f('UNICODE(MID(A4,9,1))'), B4: f('A1=A3'), B5: f('LEN(A7)'),
      },
    },
  ],
}

const formulas: BookInput = {
  sheets: [
    {
      name: 'Data',
      cells: {
        A1: v(10), A2: v(20), A3: v(30), A4: v('x'), B1: v('a'), B2: v('b'), B3: v('c'),
        C1: f('SUM(A1:A3)'), C2: f('A1/0'), C3: f('XLOOKUP(20,A1:A3,B1:B3)'), C4: f('IFS(A1>5,"big",TRUE,"small")'), C5: f('TEXTJOIN("-",TRUE,B1:B3)'),
        C6: f('CONCAT(B1:B3)'), C7: f('XMATCH(30,A1:A3)'), C8: f('IFNA(MATCH(99,A1:A3,0),"none")'), C9: f('SWITCH(A2,10,"ten",20,"twenty")'),
        C10: f('STDEV.S(A1:A3)'), C11: f('CEILING.MATH(2.5)'), C12: f('A4+1'), C13: f('rate*2'), C14: f("'Other sheet'!A1+1"), C15: f('SUM(A:A)'),
        C16: f('LET(x,A1,y,A2,x*y)'), C17: f('IF(A1>5,"yes")'), C18: f('ROUND(1/3,10)'), C19: f('1E+21*1'), C20: f('0.000001*1'),
        C21: f('"a""b"'), C22: f('TRUE'), C23: f('#N/A'), C24: f('IFERROR(1/0,"div")'), C25: f('local*1'), C26: f('NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2025,1,31),1)'), C27: f('""'), C28: f('B4&""'),
      },
    },
    { name: 'Other sheet', cells: { A1: v(5), A2: f('local+Data!A1'), B1: f('SUM(Data!A1:A3)') } },
  ],
  names: [
    { name: 'rate', ref: "'Data'!$A$1" },
    { name: 'local', ref: "'Other sheet'!$A$1", scope: 'Other sheet' },
    { name: 'local', ref: "'Data'!$A$2", scope: 'Data' },
    { name: 'column', ref: "'Data'!$A:$A" },
  ],
}

const spills: BookInput = {
  sheets: [
    {
      name: 'Spill',
      cells: {
        A1: f('SEQUENCE(3,2)'), D1: f('A1#*10'), F1: f('SORT({3;1;2})'), G1: f('UNIQUE({1;1;2})'), H1: f('FILTER(A1:A3,A1:A3>1,"none")'),
        I1: f('SEQUENCE(3)'), I2: v(5), J1: f('ROWS(A1#)'), K1: f('TRANSPOSE(A1:A3)'), K3: f('{1,2;3,4}'), N1: f('SEQUENCE(2)'),
        A6: f('FILTER(A1:A3,A1:A3>100)'), B6: f('SUM(A1#)'), C6: f('A1:A3*2'), D6: { style: 0 },
      },
      merges: ['N2:O2'],
    },
  ],
  styles: [{ fill: { color: '#FFFF00' } }],
}

const styles: BookInput = {
  sheets: [
    {
      name: 'Styles',
      cells: {
        A1: { value: 'bold red', style: 0 }, A2: { value: 'fill and border', style: 1 }, A3: { value: 'aligned', style: 2 }, A4: { value: 1.5, style: 3, numberFormat: '0.000' },
        A5: { style: 1 }, A6: { value: 'Arial 14', style: 4 }, A7: { value: 'all borders', style: 5 }, B1: { value: 'merged' }, D5: { value: 'in a styled column' },
      },
      columns: { A: { width: 24 }, 'B:C': { width: 12.5, style: 1 }, D: { style: 2, numberFormat: '0.0' }, F: { width: 0 } },
      rows: { 2: { height: 30 }, '4:5': { height: 20, style: 0 }, 9: { height: 0 }, 10: { numberFormat: '0%' } },
      merges: ['B1:C2', 'E7:F8'],
    },
  ],
  styles: [
    { font: { bold: true, italic: true, color: '#FF0000', underline: 'double', strike: true } },
    { fill: { color: '#DDEBF7' }, border: { bottom: { style: 'thick', color: '#1F4E79' }, top: { style: 'dashed' } } },
    { alignment: { horizontal: 'center', vertical: 'top', wrap: true, indent: 2 } },
    { font: { size: 8 }, alignment: { horizontal: 'right' } },
    { font: { name: 'Arial', size: 14, underline: 'single' } },
    { border: { top: { style: 'thin' }, right: { style: 'medium', color: '#00B050' }, bottom: { style: 'double' }, left: { style: 'dotted' } } },
  ],
}

const sheetNames: BookInput = {
  sheets: [
    { name: "It's a sheet", cells: { A1: v(1), B1: f("'Лист 2'!A1+'_x0041_'!A1") } },
    { name: 'Лист 2', cells: { A1: v(2), B1: f("'It''s a sheet'!A1*2") } },
    { name: '_x0041_', cells: { A1: v(3) } },
    { name: 'Empty', cells: {} },
  ],
}

const ru: BookInput = {
  sheets: [{ name: 'Лист1', cells: { A1: v('1,5'), B1: f('A1*2'), C1: f('TEXT(1234.5,"# ##0,00")'), D1: f('"1.5"+0'), E1: f('TRUE&""') } }],
  locale: 'ru-RU',
}

const en: BookInput = { sheets: [{ name: 'Sheet1', cells: { A1: v('1.5'), B1: f('A1*2'), C1: f('TRUE&""') } }], locale: 'en-US' }

/** References to a deleted sheet as the engine keeps them (`#REF!A1`, as Excel shows them once the sheet is gone), in cells and a name. */
const lostSheet: BookInput = {
  sheets: [{ name: 'Plain', cells: { A1: f('#REF!E2*2'), A2: f('SUM(#REF!E2:E4)'), A3: f('SUM(#REF!E2#)'), A4: f('IFERROR(total,0)+1'), B1: v('3') } }],
  names: [{ name: 'total', ref: '#REF!$B$5' }],
}

/** Dates typed in and computed in the 1904 date system, a negative date and time among them. */
const date1904: BookInput = {
  dateSystem: 1904,
  sheets: [
    {
      name: 'Dates',
      cells: {
        A1: { value: { date: '2025-01-31' }, numberFormat: 'yyyy-mm-dd' }, A2: { value: { date: '1904-01-01T06:00:00' }, numberFormat: 22 }, A3: { value: 0, numberFormat: 14 }, A4: { value: -1.5, numberFormat: 'yyyy-mm-dd h:mm' },
        B1: f('DATE(2025,1,31)'), B2: f('YEAR(A1)'), B3: f('A1-A2'), B4: f('TEXT(A1,"yyyy-mm-dd")'), B5: f('DATEVALUE("2025-01-31")'), B6: f('WEEKDAY(A3)'),
        B7: f('EDATE(A1,1)'), B8: f('DAY(0)'), B9: f('MONTH(A3)'), B10: f('TEXT(A4,"yyyy-mm-dd h:mm")'), B11: f('TEXT(-0.5,"h:mm")'), B12: f('A1=B1'),
      },
    },
  ],
}

/** Cycles that reach their fixed point exactly, so that Excel, going on from the values written, keeps them. */
const cycles = (iteration: BookInput['iteration']): BookInput => ({
  iteration,
  sheets: [{ name: 'Cycle', cells: { A1: f('MIN(B1+1,5)'), B1: f('A1'), C1: f('IF(D1>=3,3,D1+1)'), D1: f('C1*1'), E1: f('A1+C1') } }],
})

/** Every built-in format by its number, on cells, a row and a column. */
const builtin = (locale?: string): BookInput => ({
  ...(locale ? { locale } : {}),
  sheets: [
    {
      name: 'Formats',
      cells: { ...Object.fromEntries(Array.from({ length: 49 }, (_, i) => [`A${i + 1}`, { value: 45688.75, numberFormat: i + 1 }])), B1: { value: -1234.5, numberFormat: 38 }, B2: { value: 0, numberFormat: 0 } },
      rows: { 51: { numberFormat: 14 } },
      columns: { C: { numberFormat: 44 } },
    },
  ],
})

/** The books of the structure and spill sets, each once, as the engine takes them. */
async function setBooks(): Promise<RoundTrip[]> {
  const out: RoundTrip[] = []
  for (const name of STRUCT_SETS) {
    const { books } = ((await import(`../struct/${name}.ts`)) as StructSet).buildStructCases()
    for (const [key, book] of Object.entries(books)) out.push({ id: `struct-${name}-${key}`, what: `the book ${key} of the structure set ${name}`, book: takenBook(book) })
  }
  for (const name of SPILL_SETS)
    for (const c of ((await import(`../spill/${name}.ts`)) as SpillSet).buildSpillCases())
      out.push({ id: `spill-${name}-${c.id}`, what: c.about, book: takenBook(c.book), ...(c.volatile ? { volatile: true as const } : {}) })
  for (const name of TABLE_SETS) for (const c of ((await import(`../tables/${name}.ts`)) as TableSet).buildTableCases()) out.push({ id: `tables-${name}-${c.id}`, what: c.about, book: takenBook(c.book) })
  const seen = new Set<string>()
  // a whole column spilled: Excel saves its million cells, 5.6 MB a file
  const huge = new Set(['spill-base-col-free', 'spill-base-col-format'])
  return out.filter((r) => !huge.has(r.id) && !seen.has(JSON.stringify(r.book)) && !!seen.add(JSON.stringify(r.book)))
}

export async function buildRoundTrips(): Promise<RoundTrip[]> {
  return [
    ...ownBooks(),
    ...(await setBooks()),
  ]
}

function ownBooks(): RoundTrip[] {
  return [
    { id: 'values', what: 'numbers, text with spaces, breaks and quotes, empty text, booleans, dates, number formats', book: values },
    { id: 'escapes', what: "Excel's _xHHHH_ escapes, control characters, the longest text", book: escapes },
    { id: 'formulas', what: 'functions newer than Excel 2007, names global and local, sheets, errors, LET, empty text results', book: formulas },
    { id: 'formulas-values', what: 'the same book as values only', book: formulas, mode: 'values' },
    { id: 'spills', what: 'dynamic arrays: spills, spill references, a blocked spill, an empty FILTER, a merged area in the way', book: spills },
    { id: 'spills-values', what: 'the spills as values only', book: spills, mode: 'values' },
    { id: 'styles', what: 'fonts, fills, borders, alignment, number formats, widths, heights, column and row styles, merged areas', book: styles },
    { id: 'sheet-names', what: 'sheet names with an apostrophe, Cyrillic, an escape-like name, an empty sheet', book: sheetNames },
    { id: 'ru', what: 'a ru-RU book: the locale kept as a custom property', book: ru },
    { id: 'en', what: 'an en-US book with its locale given', book: en },
    { id: 'lost-sheet', what: 'references to a deleted sheet in cells and a name, which a file holds as #REF!', book: lostSheet },
    { id: 'synth', what: 'the synthetic book (340 rows): SUMIFS over whole columns, names, XLOOKUP, Cyrillic', book: synthBook({ seed: 1, rows: 340 }).book },
    {
      id: 'hidden',
      what: 'rows and columns hidden with and without their sizes and looks; SUBTOTAL and AGGREGATE over them',
      book: filterBook({ rows: { 3: { hidden: true }, 5: { height: 30, hidden: true }, 6: { hidden: true, numberFormat: '0.0' } }, columns: { E: { width: 12, hidden: true }, G: { hidden: true } } }),
    },
    { id: 'zero-sizes', what: 'a row of height 0 and a column of width 0, not marked hidden', book: filterBook({ rows: { 3: { height: 0 } }, columns: { E: { width: 0 } } }) },
    { id: 'filtered', what: 'a filtered sheet: its hidden rows taken as rows a filter hid', book: filterBook({ rows: Y_ROWS, filtered: true }) },
    { id: 'filtered-table', what: "a filtered sheet with a table, which holds the sheet's filter mode", book: filterBook({ rows: Y_ROWS, filtered: true, table: true }) },
    { id: 'autofilter', what: "a sheet's AutoFilter without the filter mode", book: withFilter(filterBook({})) },
    { id: 'autofilter-filtered', what: "a filtered sheet's AutoFilter, which holds the sheet's filter mode", book: withFilter(filterBook({ rows: Y_ROWS, filtered: true })) },
    { id: 'date1904', what: 'the 1904 date system: dates typed in and computed, negative ones', book: date1904 },
    { id: 'iteration', what: 'iterative calculation with a count and a change of its own over cycles that settle', book: cycles({ maxIterations: 7, maxChange: 0.5 }) },
    { id: 'iteration-default', what: "iterative calculation with Excel's defaults", book: cycles({}) },
    { id: 'builtin', what: 'the built-in number formats 1 to 49 by number on cells, a row and a column', book: builtin() },
    { id: 'builtin-ru', what: 'the built-in number formats by number in a ru-RU book', book: builtin('ru-RU') },
  ]
}
