/**
 * Excel cases for CHAR and CODE where the code page does not matter: codes 1–127 (ASCII in Windows-1252 and
 * Windows-1251, which the engine follows, and in Mac Roman and Mac Cyrillic, which Excel for Mac uses), codes
 * outside 1–255, characters outside all four tables, and the coercion of the arguments. Codes 128–255 are
 * checked against the published tables instead (`test/codepage.test.ts`). Computed as dynamic array formulas,
 * in Excel en-US (`test/excel/cases/char-code.json`) and ru-RU (`test/excel/cases/ru-RU/char-code.json`).
 */
import type { Probe, SheetCells } from '../probes.js'
import { xl } from './arrays/common.js'

/** G: codes and texts read by reference. */
export const SHEETS: SheetCells = { G: { A1: 65, A2: 97, A3: 'B', A4: 'b', B1: '65', B2: '' } }

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add = (fn: string, formula: string) => out.push({ id: `${fn}#${out.filter((p) => p.fn === fn).length + 1}`, group: fn, fn, formula: xl(formula) })
  for (let n = 1; n <= 127; n++) add('CHAR', `CHAR(${n})`)
  for (let n = 1; n <= 127; n++) add('CODE', `CODE(CHAR(${n}))`)
  // the code taken whole, outside 1–255, numbers and text and blanks as arguments
  for (const a of ['0', '256', '-1', '65.9', '65.1', '0.5', '0.99999', '255.5', '1E+300', '-0.5', '"65"', '" 65 "', '"6.5E+1"', '"1,5"', '"abc"', '""', 'TRUE', 'FALSE', 'D!A1', 'G!B1', 'G!B2', 'NA()', '1/0'])
    add('CHAR', `CHAR(${a})`)
  for (const a of ['{65,66,67}', 'G!A1:A2', '{1;127}']) add('CHAR', `TEXTJOIN("|",FALSE,CHAR(${a}))`)
  add('CHAR', 'LEN(CHAR(200))')
  add('CHAR', 'ISTEXT(CHAR(255))')
  // the first character's code; text, numbers, blanks and errors as arguments
  for (const a of ['"A"', '"abc"', '" "', '"~"', '""', '65', '1.5', '-1', '0.5', 'D!A1', 'G!B2', 'G!B1', 'NA()', 'CHAR(10)&"x"'])
    add('CODE', `CODE(${a})`)
  // characters in none of Windows-1252, Windows-1251, Mac Roman and Mac Cyrillic
  for (const ch of ['中', '😀', 'ā', 'Ā', 'ŋ', 'ğ', 'ʃ']) add('CODE', `CODE("${ch}")`)
  for (const a of ['{"A","b"}', 'G!A3:A4']) add('CODE', `SUM(CODE(${a}))`)
  return out
}
