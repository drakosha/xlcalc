/**
 * Texts typed on the keyboard into Excel for Mac (System Events, an ASCII keyboard layout), against the same texts set
 * through `Formula2Local` (the set `typed` and its answers on Windows): what the object model may take otherwise than
 * typing, a leading `+` or `-` before what is no number and plain numbers typed into a percent cell (Excel's automatic
 * percent entry), and a few of each other kind. No text here makes Excel show a window: typed, a formula Excel does not
 * take would stop the run on its dialog.
 */
import { collector, type EntryCase } from './cases.js'

export const keys = true

const GENERAL = ['5', '+5', '-5', '+1+2', '-1-2', '-5+1', '--5', '+-5', '+5%', '-5%', '-C1', '+C1', '-qq', '- 5', '1 1/2', '1/2', "'007", '5%', '18:30', '=1+2', '1.5E+3', '1,5E+3', '3/15/2026', '15.03.2026', '1,234.5', '1 234,5', 'TRUE']
const PERCENT = ['5', '50', '0.5', '0,5', '-5', '1', '0', '5%', '1 1/2']

export function buildEntryCases(): EntryCase[] {
  const { out, add } = collector()
  for (const t of GENERAL) add('key', t)
  for (const t of PERCENT) add('key-pct', t, '0%')
  for (const t of ['5', '0.5', '0,5']) add('key-pct2', t, '0.00%')
  return out
}
