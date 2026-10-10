/**
 * Held out from the set `typed-keys`, typed on the keyboard into Excel for Mac and predicted by the engine before Excel
 * ran it: what Excel's editor of a percent cell adds to other texts, more formulas begun by `+` and `-`, and texts Excel
 * does not take as formulas, whose dialog the step answers with OK and then Escape, leaving the cell as it was. Only
 * ASCII is typed: the keyboard layout (ABC) has no ₽, so `-5 ₽` of the predictions is not run.
 */
import { collector, type EntryCase } from './cases.js'

export const keys = true

const PERCENT = ['abc', 'abc5', '=1+2', "'5", '.5', ',5', '+5', '$5', '(5)', ' 5', '5 ', '15.03.2026', '3/15/2026', '18:30', 'TRUE', '1,234.5', '1 234,5', '1.5E+3', '1,5E+3', '-0.5', '-0,5', '1/2']
const SIGNS = ['+1*2', '+(1+2)', '+C1+1', '+10-1', '+2^3', '++5', '-(1+2)', '-C1*2', '+-C1', '-1 -2', '+1 1/2']
const REFUSED = ['-1.5', '-1,5', '- hello world', '-15.03.2026', '=1+', '=', '-', '+', '==1', '+C1+', '-$5']

export function buildEntryCases(): EntryCase[] {
  const { out, add } = collector()
  for (const t of PERCENT) add('keyh-pct', t, '0%')
  for (const t of SIGNS) add('keyh-sign', t)
  for (const t of REFUSED) add('keyh-no', t)
  return out
}
