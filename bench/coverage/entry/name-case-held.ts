/**
 * Held out, predicted by the engine before Excel ran it: formulas typed with words whose letters JavaScript's upper case
 * changes otherwise than one for one (`ſ` to `S`, `ß` to `SS`, `ﬁ` to `FI`, `ı` to `I`; `İ` stays): FALSE spelled with a
 * long s, as a value, in an array and as LET's name; function names; LET's names against the letter's upper case. The
 * formula Excel saves shows which word it took (a function's name in its own spelling, a LET name as declared).
 */
import { collector, type EntryCase } from './cases.js'

const GENERAL: Record<string, string[]> = {
  bool: ['=falſe', '=FALſE', '=IF(falſe,1,2)', '={falſe}', '=LET(falſe,1,falſe)', '=LET(x,1,falſe)'],
  fn: ['=ſum(1,2)', '=ﬁnd("a","a")', '=gauß(1)', '=addreß(1,1)', '=ıf(TRUE,1,2)', '=İF(TRUE,1,2)', '=if(TRUE,1,2)'],
  let: ['=LET(ſ,1,S)', '=LET(S,1,ſ+1)', '=LET(ß,1,SS)', '=LET(SS,1,ß+1)', '=LET(ı,1,I)', '=LET(I,1,ı+1)', '=LET(İ,1,i)', '=LET(i,1,İ+1)', '=LET(ﬁ,1,FI)'],
}

export function buildEntryCases(): EntryCase[] {
  const { out, add } = collector()
  for (const [group, texts] of Object.entries(GENERAL)) for (const t of texts) add(group, t)
  return out
}
