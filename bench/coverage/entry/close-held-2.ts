/**
 * Held out after `close-held`, typed on the keyboard into Excel for Mac and predicted by the engine before Excel ran it,
 * with the rule `close-held` corrected: Excel closes without a word only a formula begun by a function, a sign before it
 * or not, with that function's bracket the one left open and no `)` typed; it offers to correct any other left open, but
 * not an array constant left open in a function's brackets. These cases tell signs before the function, an array
 * constant closed or open beside a function's bracket, strings beside one, references, a space before the function and
 * a correction typed in lower case. ASCII only (the keyboard's layout is ABC); the cells read C1 and C2, never column A.
 */
import { collector, type EntryCase } from './cases.js'

export const keys = true

const CASES: Readonly<Record<string, readonly string[]>> = {
  sign: ['=+SUM(1', '=--SUM(1', '=-SUM(MAX(1', '=-(SUM(1', '=-FOO(1'],
  array: ['=SUM({1,2}', '=SUM(1,{1,2', '={1,2;3', '=1+{1,2', '=SUM({1,2},{3'],
  string: ['=SUM(1,"ab', '="ab"&SUM(1', '=LEN("a)b"', '=SUM("ab"'],
  ref: ['=SUM(C1:C2', '=SUM(S!C1', "=SUM('S'!C1", '=MAX(C1,C2'],
  shape: ['= SUM(1', '=1+sum(2', '=sum(max(1', '=SUM(1,2,'],
}

export function buildEntryCases(): EntryCase[] {
  const { out, add } = collector()
  for (const [group, texts] of Object.entries(CASES)) for (const t of texts) add(group, t)
  return out
}
