/**
 * Held out: the set `name-case-held` in Excel ru-RU (`run.ts tables name-case-ru-held ru`), whether the case of defined
 * names depends on the locale: a name added in one spelling (`ſ`, `ß`, `ı`, `İ`, `ﬁ` or the one JavaScript's upper case
 * gives) and formulas over both typed in Russian. Predicted by the engine (`caseKey`) before Excel ran it
 * (`name-case-ru-held.predicted.json`).
 */
import type { BookInput } from '../../../src/contract.js'
import { book, on, type TableCase, type TableStep } from './cases.js'

export const local = true

const typed = (cell: string, text: string): TableStep => ({ edit: { sheet: 'S', cell, value: { local: text } } })
const ru = (b: BookInput): BookInput => ({ ...b, locale: 'ru-RU' })

/** The pairs: the spelling JavaScript's upper case changes, the one it gives. */
const PAIRS: [id: string, word: string, upper: string][] = [
  ['long-s', 'ſ', 'S'],
  ['sharp-s', 'ß', 'SS'],
  ['dotless-i', 'ı', 'I'],
  ['dotted-i', 'İ', 'i'],
  ['ligature-fi', 'ﬁ', 'FI'],
]

export function buildTableCases(): TableCase[] {
  return PAIRS.flatMap(([id, word, upper]) =>
    ([['word', word], ['upper', upper]] as const).map(([which, name]): TableCase => ({
      id: `${id}-${which}`,
      about: `the name ${name} added in ru-RU, formulas over ${upper} and ${word} typed in Russian`,
      book: ru(book(['S', { A1: 100 }])),
      steps: [{ name: { name, ref: '5' } }, typed('C1', `=${upper}*2`), typed('C2', `=${word}*2`)],
      read: on('S', 'A1:C3'),
    })),
  )
}
