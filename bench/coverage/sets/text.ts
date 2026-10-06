/**
 * Excel cases of the text, lookup, database, logical and information functions beyond the coverage
 * measurement: wildcards, case and Cyrillic, approximate search over sorted data, errors on input, number
 * text, database criteria; the data sheets and the cases by group are in `./text/`. Computed as dynamic array
 * formulas (Excel 365).
 */
import type { Probe } from '../probes.js'
import { SHEETS, type Add } from './text/data.js'
import { lookup } from './text/lookup.js'
import { database, logical } from './text/other.js'
import { bytes, formats, strings, tildes } from './text/strings.js'

export { SHEETS }

/** Functions written with the `_xlfn.` prefix in a workbook (newer than Excel 2007). */
const XLFN = /\b(IFNA|SWITCH|XLOOKUP|XMATCH|NUMBERVALUE|UNICHAR|TEXTJOIN|TEXTBEFORE|TEXTAFTER|VALUETOTEXT|ARRAYTOTEXT)\(/g

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add: Add = (fn, ...formulas) => {
    for (const f of formulas) out.push({ id: `${fn}#${out.filter((p) => p.fn === fn).length + 1}`, group: fn, fn, formula: f.replace(XLFN, '_xlfn.$1(') })
  }
  for (const part of [strings, lookup, database, logical, bytes, formats, tildes]) part(add)
  // the & operator: Excel cuts its result at 32,767 characters
  for (const [i, f] of ['LEN(REPT("a",20000)&REPT("b",20000))', 'RIGHT(REPT("a",20000)&REPT("b",20000),1)', 'LEN(REPT("a",32767)&"x")', 'LEN("x"&REPT("a",32767))'].entries())
    out.push({ id: `&#${i + 1}`, group: '&', formula: f })
  return out
}
