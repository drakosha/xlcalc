/**
 * Held out from the sets `typed-keys` and `typed-keys-held`, typed on the keyboard into Excel for Mac and predicted by
 * the engine after their rules were taken in, before Excel ran it: more numbers typed into a percent cell, more texts
 * begun by `+` and `-`. ASCII only, no text that would bring a window Excel's step does not answer.
 */
import { collector, type EntryCase } from './cases.js'

export const keys = true

const PERCENT = ['1/4', '-1/2', '2 1/2', '5.5', '5,5', '+.5', '-.5', '1e2', '007', '5 5', '1,234', '1 234', '-(5)', '5-']
const SIGNS = ['+.5+1', '+0', '+1.5', '+1,5', '++C1', '+ 5', '+(5)', '-0', '- 1', '-1e2', '+1e2+1', '-C1%', '+5%+1', '--C1', '+-+5']

export function buildEntryCases(): EntryCase[] {
  const { out, add } = collector()
  for (const t of PERCENT) add('key2-pct', t, '0%')
  for (const t of SIGNS) add('key2-sign', t)
  return out
}
