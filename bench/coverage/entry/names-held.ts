/**
 * Held out, predicted by the engine before Excel ran it: which words Excel takes as LET's names in a typed formula, and
 * how it keeps an unknown name typed in lower case. LET's names shaped like a cell of the grid (`R1`, `C5`, `XFD1`),
 * past it (`XFE1`, `A0`), like R1C1 (`R`, `r`, `C`, `RC`, `R1C1`, `R1C`), of four letters (`ZZZZ1`), with `?` and `\`
 * inside, first and after a cell's shape (`a?b`, `\x`, `A1?x`); the words past the grid alone in a formula; unknown names
 * and an unknown function in lower and mixed case, beside a function typed in lower case.
 */
import { type EntryCase, entryCases } from './cases.js'

const let1 = (name: string) => `=LET(${name},1,${name})`

const GENERAL: Record<string, string[]> = {
  'let-cell': ['R1', 'C5', 'XFD1', 'XFE1', 'A0', 'ZZZZ1'].map(let1),
  'let-r1c1': ['R', 'r', 'C', 'RC', 'R1C1', 'R1C'].map(let1),
  'let-chars': ['a?b', 'x?', '\\x', 'a\\b', 'A1?x', 'R1\\x'].map(let1),
  word: ['=XFE1', '=A0', '=XFE1+1'],
  unknown: ['=foo', '=foo+1', '=Foo*2', '=SUM(foo,1)', '=fooBar', '=foo(1)', '=sum(1,2)'],
}

export const buildEntryCases = (): EntryCase[] => entryCases(GENERAL)
