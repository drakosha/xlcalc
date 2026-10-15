/**
 * Held out, typed on the keyboard into Excel for Mac and predicted by the engine before Excel ran it: formulas with
 * brackets, quotes or array braces left open at their end, entered with Return. The recorded answers of Point mode
 * (`ui/point-caret-held-2`, `-3`) gave the rule predicted: Excel closes a formula begun by a function, with no `)` typed
 * and only functions' brackets open, without a word, and offers to correct any other, when it reads closed with every
 * function known. These cases tell nested and later calls, a closing bracket typed, strings, array constants, unknown
 * names and a sign before the formula. ASCII only (the keyboard's layout is ABC); the cells read C1, never column A.
 */
import { collector, type EntryCase } from './cases.js'

export const keys = true

const CASES: Readonly<Record<string, readonly string[]>> = {
  first: ['=SUM(1,2', '=sum(1,2', '=SUM(C1', '=SUM(MAX(1,2', '=SUM(1,MAX(2', '=IF(C1>0,SUM(1', '=SUM(1,', '=SUM(1+', '=ROUND(1.5,', '=SUM("(",1', '=SUM(")",1'],
  later: ['=1+SUM(2', '=-SUM(1', '=C1*SUM(2', '=SUM(1)+SUM(2', '=SUM(MAX(1),2', '=SUM(1)*(2'],
  group: ['=(1+2', '=SUM(1,(2', '=((1+2)', '=(1+2)*(3'],
  string: ['="ab', '=SUM("ab', '=LEN("ab', '="a"&"b', '=SUM("a",2'],
  array: ['={1,2', '=SUM({1,2', '=SUM({1,2},3'],
  name: ['=FOO(1', '=FOO((1', '=1+FOO(2'],
  sign: ['-SUM(1', '+SUM(1,2'],
}

export function buildEntryCases(): EntryCase[] {
  const { out, add } = collector()
  for (const [group, texts] of Object.entries(CASES)) for (const t of texts) add(group, t)
  return out
}
