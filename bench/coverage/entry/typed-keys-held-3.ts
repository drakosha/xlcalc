/**
 * Held out, typed on the keyboard into Excel for Mac and predicted by the engine before Excel ran it: texts with a sign
 * and a percent sign typed into a percent cell, where Excel's percent editor adds a second `%` (`5%` is the text `5%%`)
 * and a text begun by a sign that is no constant is a formula. In place (F2, the set `ui/edit-held`) Excel made `-5%`
 * the formula `-5%%`. ASCII only, no text that would bring a window Excel's step does not answer.
 */
import { collector, type EntryCase } from './cases.js'

export const keys = true

const PERCENT = ['-5%', '+5%', '-0.5%', '(5%)', '-5', '+5']

export function buildEntryCases(): EntryCase[] {
  const { out, add } = collector()
  for (const t of PERCENT) add('key3-pct', t, '0%')
  return out
}
