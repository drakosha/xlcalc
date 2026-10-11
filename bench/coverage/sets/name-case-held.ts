/**
 * Held out: words of a formula in a file whose letters JavaScript's upper case changes otherwise than one for one: the
 * long s (`ſ`, upper case `S`), the sharp s (`ß`, `SS`), the ligature `ﬁ` (`FI`), the dotless `ı` (`I`) and the dotted
 * `İ` (lower case `i̇`). The engine compared function names, TRUE and FALSE and LET's names by that upper case, so
 * `falſe` read as FALSE and `ſum(1,2)` as SUM(1,2), and in some places by an ASCII pattern, so `{falſe}` was refused.
 * LET's names are written as a file writes them (`_xlpm.`). Answers: `test/excel/cases/name-case-held.json`;
 * `PREDICTED`, the engine's answers, was committed before Excel ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { collector } from './arrays/common.js'

const CASES: Record<string, [fn: string, formulas: string[]]> = {
  bool: ['IF', ['falſe', 'FALſE', 'TYPE(falſe)', 'IF(falſe,1,2)', 'ISLOGICAL({falſe})', 'NOT(falſe)', 'false']],
  let: ['LET', ['_xlfn.LET(_xlpm.falſe,1,_xlpm.falſe)', '_xlfn.LET(_xlpm.x,1,falſe)', '_xlfn.LET(_xlpm.ſ,1,_xlpm.ſ+1)']],
  text: ['IF', ['IF("falſe",1,2)', 'IF("false",1,2)']],
  fn: [
    'SUM',
    ['ſum(1,2)', 'ſUM(1,2)', 'ﬁnd("a","a")', 'ﬁxed(1)', 'gauß(1)', 'addreß(1,1)', 'ıf(TRUE,1,2)', 'İF(TRUE,1,2)', 'ısnumber(1)', 'İSNUMBER(1)', 'if(TRUE,1,2)', 'sum(1,2)'],
  ],
}

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [group, [fn, formulas]] of Object.entries(CASES)) add(group, fn, ...formulas)
  return out
}

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./name-case-held.predicted.json', import.meta.url), 'utf8'))
