/**
 * Books for the xlsx round trip: the writer writes each one, Excel opens the file (a repair window is the
 * writer's fault), computes it and saves it as xlsx, the reader reads what Excel saved, and the book that comes
 * back is compared with the one written, Excel's stored values with the engine's (`compare.ts`). No real data:
 * the books are made up here, the synthetic book comes from its seeded generator.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import { synthBook } from '../../../src/synth.js'

/** The clock the round trip computes its books with. */
export const XLSX_CLOCK = { now: '2026-10-07T12:00:00Z', seed: 7 }

export interface RoundTrip {
  id: string
  what: string
  book: BookInput
  /** Written with formulas and their values (the default), or as values only. */
  mode?: 'formulas' | 'values'
}

const v = (value: CellInput['value']): CellInput => ({ value })
const f = (formula: string): CellInput => ({ formula })

const values: BookInput = {
  sheets: [
    {
      name: 'Values',
      cells: {
        A1: v(0), A2: v(-1.5), A3: v(0.1 + 0.2), A4: v(1e-300), A5: v(1.7976931348623157e308), A6: v(123456789012345), A7: v(-0),
        B1: v('text'), B2: v('  leading and trailing  '), B3: v('line\nbreak'), B4: v('tab\there'), B5: v('cr\r\nlf'), B6: v('quotes " \' and & < >'),
        B7: v('Кириллица и ёлка'), B8: v('emoji 😀 and 中文'), B9: v('=not a formula'), B10: v("'apostrophe first"),
        C1: v(true), C2: v(false), D1: v({ date: '2025-01-31' }), D2: v({ date: '1900-02-28' }), D3: v({ date: '2025-06-30T18:30:00' }),
        E1: { value: 1, format: '0.00' }, E2: { value: 0.25, format: '0%' }, E3: { value: 45688, format: 'dd.mm.yyyy' }, E4: { value: 45688.75, format: 'm/d/yyyy h:mm' },
        E5: { value: 1234.5, format: '#,##0.00 "₽"' }, E6: { value: -5, format: '0;[Red]-0' }, E7: { value: 'x', format: '@' },
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
        C21: f('"a""b"'), C22: f('TRUE'), C23: f('#N/A'), C24: f('IFERROR(1/0,"div")'), C25: f('local*1'), C26: f('NETWORKDAYS.INTL(DATE(2025,1,1),DATE(2025,1,31),1)'),
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
        A1: { value: 'bold red', style: 0 }, A2: { value: 'fill and border', style: 1 }, A3: { value: 'aligned', style: 2 }, A4: { value: 1.5, style: 3, format: '0.000' },
        A5: { style: 1 }, A6: { value: 'Arial 14', style: 4 }, A7: { value: 'all borders', style: 5 }, B1: { value: 'merged' }, D5: { value: 'in a styled column' },
      },
      columns: { A: { width: 24 }, 'B:C': { width: 12.5, style: 1 }, D: { style: 2, format: '0.0' }, F: { width: 0 } },
      rows: { 2: { height: 30 }, '4:5': { height: 20, style: 0 }, 9: { height: 0 }, 10: { format: '0%' } },
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

export function buildRoundTrips(): RoundTrip[] {
  return [
    { id: 'values', what: 'numbers, text with spaces, breaks and quotes, booleans, dates, number formats', book: values },
    { id: 'escapes', what: "Excel's _xHHHH_ escapes, control characters, the longest text", book: escapes },
    { id: 'formulas', what: 'functions newer than Excel 2007, names global and local, sheets, errors, LET', book: formulas },
    { id: 'formulas-values', what: 'the same book as values only', book: formulas, mode: 'values' },
    { id: 'spills', what: 'dynamic arrays: spills, spill references, a blocked spill, an empty FILTER, a merged area in the way', book: spills },
    { id: 'spills-values', what: 'the spills as values only', book: spills, mode: 'values' },
    { id: 'styles', what: 'fonts, fills, borders, alignment, number formats, widths, heights, column and row styles, merged areas', book: styles },
    { id: 'sheet-names', what: 'sheet names with an apostrophe, Cyrillic, an escape-like name, an empty sheet', book: sheetNames },
    { id: 'ru', what: 'a ru-RU book: the locale kept as a custom property', book: ru },
    { id: 'en', what: 'an en-US book with its locale given', book: en },
    { id: 'synth', what: 'the synthetic book (340 rows): SUMIFS over whole columns, names, XLOOKUP, Cyrillic', book: synthBook({ seed: 1, rows: 340 }).book },
  ]
}
