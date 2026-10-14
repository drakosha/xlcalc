/**
 * Held out: the set `name-case-held` in Excel ru-RU (`run.ts tables name-case-ru-held ru`), whether the case of defined
 * names depends on the locale: a name added in one spelling (`ſ`, `ß`, `ı`, `İ`, `ﬁ` or the one JavaScript's upper case
 * gives) and formulas over both typed in Russian. Predicted by the engine (`caseKey`) before Excel ran it
 * (`name-case-ru-held.predicted.json`).
 */
import { PAIRS } from './name-case-held.js'
import { book, inRu, on, type TableCase, typedLocal } from './cases.js'

export const local = true

export function buildTableCases(): TableCase[] {
  return PAIRS.flatMap(([id, word, upper]) =>
    ([['word', word], ['upper', upper]] as const).map(([which, name]): TableCase => ({
      id: `${id}-${which}`,
      about: `the name ${name} added in ru-RU, formulas over ${upper} and ${word} typed in Russian`,
      book: inRu(book(['S', { A1: 100 }])),
      steps: [{ name: { name, ref: '5' } }, typedLocal('C1', `=${upper}*2`), typedLocal('C2', `=${word}*2`)],
      read: on('S', 'A1:C3'),
    })),
  )
}
