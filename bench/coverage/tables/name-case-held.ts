/**
 * Held out: whether Excel takes two spellings for one defined name where JavaScript's upper case makes them one key: the
 * long s and S (`ſ`, `S`), the sharp s and SS (`ß`), the dotless and the dotted i (`ı` and I, `İ` and i), the ligature
 * `ﬁ` and FI. Each case adds one of the two through Excel's object model (`Names.Add`) and types formulas over both, so
 * a formula over the other spelling reads the name or gives #NAME?; one case adds both. Predicted by the engine before
 * Excel ran it (`name-case-held.predicted.json`).
 */
import { addedName, book, on, type TableCase, typedFormula } from './cases.js'

/** The pairs: the spelling JavaScript's upper case changes, the one it gives (the sets `name-case-*` take them too). */
export const PAIRS: [id: string, word: string, upper: string][] = [
  ['long-s', 'ſ', 'S'],
  ['sharp-s', 'ß', 'SS'],
  ['dotless-i', 'ı', 'I'],
  ['dotted-i', 'İ', 'i'],
  ['ligature-fi', 'ﬁ', 'FI'],
]

export function buildTableCases(): TableCase[] {
  const cases: TableCase[] = []
  for (const [id, word, upper] of PAIRS)
    for (const [which, name] of [['word', word], ['upper', upper]] as const)
      cases.push({
        id: `${id}-${which}`,
        about: `the name ${name} added, formulas over ${upper} and ${word}`,
        book: book(['S', { A1: 100 }]),
        steps: [addedName(name, '5'), typedFormula('C1', `${upper}*2`), typedFormula('C2', `${word}*2`)],
        read: on('S', 'A1:C3'),
      })
  cases.push({
    id: 'long-s-both',
    about: 'the names S and ſ added in turn: one name or two',
    book: book(['S', { A1: 100 }]),
    steps: [addedName('S', '5'), addedName('ſ', '7'), typedFormula('C1', 'S*2'), typedFormula('C2', 'ſ*2')],
    read: on('S', 'A1:C3'),
  })
  return cases
}
