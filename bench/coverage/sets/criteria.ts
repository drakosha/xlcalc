/**
 * Excel cases for an error as a criterion, in every function with criteria: an error value (from a cell),
 * its text in any case (`"#n/a"`), with `=`, `<>` and the order operators, over ranges holding errors, text
 * that reads like an error, blanks and whole columns, and errors in the sum range. The data sheet `D` holds
 * the errors (`D!A7` #DIV/0!, `D!A8` #N/A, `D!E3` #DIV/0!); `CR` holds the values to sum and text cells.
 * Computed as dynamic array formulas (Excel 365).
 */
import type { Probe } from '../probes.js'

export const DYNAMIC = true

export const SHEETS: Record<string, Record<string, number | string | boolean>> = {
  CR: {
    // 2^0 … 2^14 beside D!A1:A15: a sum tells which rows were taken
    ...Object.fromEntries(Array.from({ length: 15 }, (_, i) => [`A${i + 1}`, 2 ** i])),
    ...Object.fromEntries(['a', 'b', 'a', 'b', 'a', 'b', 'a', 'b'].map((v, i) => [`B${i + 1}`, v])),
    // text, not errors
    C1: '#DIV/0!',
    C2: '#N/A',
    C3: '#NUM!',
    C4: 'x',
    ...Object.fromEntries([10, 20, 30, 40].map((v, i) => [`D${i + 1}`, v])),
    // criteria ranges for the database D!A6:A8 (header 45688, records #DIV/0! and #N/A)
    E1: 45688,
    E2: '#N/A',
    F1: 45688,
    F2: '#DIV/0!',
    G1: 45688,
    G2: '<>#DIV/0!',
  },
}

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const count = new Map<string, number>()
  const add = (fn: string, formula: string) => {
    const n = (count.get(fn) ?? 0) + 1
    count.set(fn, n)
    out.push({ id: `${fn}#${n}`, group: fn, fn, formula })
  }
  const X = '_xlfn.'
  for (const c of ['D!A7', 'D!A8', '1/0', '"#DIV/0!"', '"=#DIV/0!"', '"<>#DIV/0!"', '"#N/A"', '"<>#N/A"', '"=#N/A"', '"#n/a"', '"#DIV*"', '">#DIV/0!"', '"<#N/A"', '"#NUM!"', '"#NULL!"', '"#VALUE!"', '" #N/A"', '"#DIV/0"'])
    add('COUNTIF', `COUNTIF(D!A1:A15,${c})`)
  for (const f of ['COUNTIF(D!A:A,D!A7)', 'COUNTIF(D!A:A,"<>#N/A")', 'COUNTIF(CR!C1:C4,"#DIV/0!")', 'COUNTIF(CR!C1:C4,"#N/A")', 'COUNTIF(CR!C1:C4,"#NUM!")', 'COUNTIF(CR!C1:C4,D!A7)', 'COUNTIF(CR!C1:C4,"<>#N/A")', 'COUNTIF(D!E1:E8,D!A7)', 'COUNTIF(D!E1:E8,"<>#DIV/0!")'])
    add('COUNTIF', f)
  add('ISTEXT', 'ISTEXT(CR!C1)')
  for (const c of ['D!A7', 'D!A8', '"#N/A"', '"<>#DIV/0!"', '"=#DIV/0!"', '">=#DIV/0!"']) add('SUMIF', `SUMIF(D!A1:A15,${c},CR!A1:A15)`)
  for (const f of ['SUMIF(CR!C1:C4,"#N/A",CR!D1:D4)', 'SUMIF(D!E1:E8,D!A7,D!B1:B8)', 'SUMIF(D!A1:A15,D!A7)', 'SUMIF(CR!B1:B8,"a",D!E1:E8)', 'SUMIF(CR!B1:B8,"b",D!E1:E8)'])
    add('SUMIF', f)
  for (const f of ['SUMIFS(CR!A1:A15,D!A1:A15,D!A7)', 'SUMIFS(CR!A1:A15,D!A1:A15,"<>#N/A",D!A1:A15,"<>#DIV/0!")', 'SUMIFS(D!E1:E8,CR!B1:B8,"a")', 'SUMIFS(D!B1:B8,D!E1:E8,"#DIV/0!")'])
    add('SUMIFS', f)
  for (const f of ['COUNTIFS(D!A1:A15,"#N/A",CR!A1:A15,">100")', 'COUNTIFS(D!A1:A15,D!A7,D!A1:A15,"<>#N/A")', 'COUNTIFS(D!E1:E8,D!A7)', 'COUNTIFS(D!A:A,"<>#DIV/0!",D!A:A,"<>#N/A")'])
    add('COUNTIFS', f)
  for (const f of ['AVERAGEIF(D!A1:A15,"#N/A",CR!A1:A15)', 'AVERAGEIF(D!A1:A15,D!A7)', 'AVERAGEIF(CR!B1:B8,"b",D!E1:E8)', 'AVERAGEIF(D!E1:E8,D!A7,D!B1:B8)']) add('AVERAGEIF', f)
  for (const f of ['AVERAGEIFS(CR!A1:A15,D!A1:A15,D!A8)', 'AVERAGEIFS(D!B1:B8,D!E1:E8,"#DIV/0!")']) add('AVERAGEIFS', f)
  for (const f of [`${X}MAXIFS(CR!A1:A15,D!A1:A15,"<>#N/A")`, `${X}MAXIFS(D!E1:E8,CR!B1:B8,"b")`, `${X}MAXIFS(D!E1:E8,CR!B1:B8,"a")`]) add('MAXIFS', f)
  for (const f of [`${X}MINIFS(CR!A1:A15,D!A1:A15,"#DIV/0!")`, `${X}MINIFS(CR!A1:A15,D!A1:A15,">#DIV/0!")`]) add('MINIFS', f)
  for (const c of ['D!A6:A7', 'D!A6:A8', 'CR!E1:E2', 'CR!F1:F2', 'CR!G1:G2', 'CR!E1:E1']) add('DCOUNTA', `DCOUNTA(D!A6:A8,1,${c})`)
  add('DCOUNTA', 'DCOUNTA(D!A5:A8,1,CR!E1:E2)')
  add('DCOUNT', 'DCOUNT(D!A6:A8,1,CR!E1:E2)')
  for (const c of ['CR!E1:E2', 'D!A6:A7', 'CR!G1:G2']) add('DGET', `DGET(D!A6:A8,1,${c})`)
  return out
}
