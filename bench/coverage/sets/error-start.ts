/**
 * Formulas that begin with an error constant, held out: a host's books hold `#REF!-R` with the value `#VALUE!` where
 * the engine gives `#REF!`. Each error constant before and after the operators, minus, plus, times and `&`, over a
 * number, a text, a number as text, an empty cell and an error in a cell, and inside parentheses and functions.
 * Computed in Excel en-US on the Mac and Windows (`test/excel/cases/error-start.json`,
 * `test/excel/win/cases/error-start.json`); `error-start.predicted.json`, the engine's answers, was committed before
 * Excel ran the set.
 */
import type { Probe, SheetCells } from '../probes.js'

/** The cells the formulas read: a number, a text, a number as text, an empty cell (A4), an error. */
export const SHEETS: SheetCells = { E: { A1: 5, A2: 'abc', A3: '7', A5: { formula: '1/0' } } }

const ERRORS = ['#REF!', '#N/A', '#VALUE!', '#DIV/0!', '#NAME?', '#NUM!', '#NULL!']
const R = ['E!A1', 'E!A2', 'E!A3', 'E!A4', 'E!A5']

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add = (group: string, formula: string) => out.push({ id: `${group}#${out.filter((p) => p.group === group).length + 1}`, group, formula })
  for (const e of ERRORS) {
    for (const r of R) for (const op of ['-', '+', '*', '&']) add('first', `${e}${op}${r}`)
    for (const r of R) add('second', `${r}-${e}`)
    add('number', `${e}+3`)
    add('number', `3-${e}`)
    add('text', `${e}&"x"`)
    add('unary', `-${e}`)
    add('unary', `+${e}`)
    add('unary', `--${e}`)
    add('paren', `(${e})-E!A1`)
    add('paren', `SUM(${e}-E!A1)`)
    add('paren', `ISERROR(${e}-E!A1)`)
    add('paren', `ERROR.TYPE(${e}-E!A1)`)
  }
  add('two', '#REF!-#N/A')
  add('two', '#N/A-#REF!')
  add('two', '#REF!-E!A5')
  add('two', 'E!A5-#REF!')
  add('range', '#REF!-E!A1:A3')
  add('range', 'SUM(#REF!-E!A1:A3)')
  return out
}
