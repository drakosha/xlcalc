/**
 * Excel cases of the text, lookup, database, logical and information functions beyond the coverage
 * measurement: wildcards, case and Cyrillic, approximate search over sorted data, errors on input, number
 * text, database criteria; the data sheets and the cases by group are in `./text/`.
 */
import type { Probe } from '../probes.js'
import { SHEETS, type Add } from './text/data.js'
import { lookup } from './text/lookup.js'
import { database, logical } from './text/other.js'
import { bytes, strings } from './text/strings.js'

export { SHEETS }

/** Functions written with the `_xlfn.` prefix in a workbook (newer than Excel 2007). */
const XLFN = /\b(IFNA|SWITCH|XLOOKUP|XMATCH|NUMBERVALUE|UNICHAR|TEXTJOIN|TEXTBEFORE|TEXTAFTER|VALUETOTEXT|ARRAYTOTEXT)\(/g

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add: Add = (fn, ...formulas) => {
    for (const f of formulas) out.push({ id: `${fn}#${out.filter((p) => p.fn === fn).length + 1}`, group: fn, fn, formula: f.replace(XLFN, '_xlfn.$1(') })
  }
  for (const part of [strings, lookup, database, logical, bytes]) part(add)
  return out
}
