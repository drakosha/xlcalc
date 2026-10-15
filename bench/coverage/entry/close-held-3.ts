/**
 * Held out after `close-held-2`, typed on the keyboard into Excel for Mac and predicted by the engine before Excel ran
 * it, with the rule corrected again: spaces before the function as signs are (`= SUM(1` is closed without a word), and an
 * array constant left open corrected only where the formula is it alone. These cases tell spaces and signs mixed,
 * arrays after an operator or a sign, functions without arguments and a few functions held out. ASCII only (the
 * keyboard's layout is ABC); the cells read C1 and C2, never column A.
 */
import { collector, type EntryCase } from './cases.js'

export const keys = true

const CASES: Readonly<Record<string, readonly string[]>> = {
  space: ['=  SUM(1', '= -SUM(1', '=- SUM(1', '= (1+2'],
  array: ['={1,2}+{3', '=-{1,2', '= {1,2', '={"a","b', '=SUM(1)+{1,2'],
  call: ['=PI(', '=NOW(', '=IF(C1>0,"a', '=ROUND(C1,1', '=AVERAGE(C1', '=CONCAT("a",C1'],
}

export function buildEntryCases(): EntryCase[] {
  const { out, add } = collector()
  for (const [group, texts] of Object.entries(CASES)) for (const t of texts) add(group, t)
  return out
}
