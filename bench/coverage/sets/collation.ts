/**
 * Excel cases for the order of text on mixed Cyrillic and Latin letters (no punctuation): comparison operators,
 * SORT, criteria with `<` and `>`, exact matches and UNIQUE on letters with and without accents. Computed in
 * Excel en-US (`test/excel/cases/collation.json`) and ru-RU (`run.ts probes collation ru`,
 * `test/excel/cases/ru-RU/collation.json`): the order is one of the rules of a locale. The sheet `O` holds the
 * words. Computed as dynamic array formulas (Excel 365).
 */
import type { Probe, SheetCells } from '../probes.js'

/** Words of both scripts, cases and accents (`а`, `е`, `х` are Cyrillic where they stand among Cyrillic letters). */
const WORDS = ['b', 'а', 'A', 'Б', 'ё', 'е', 'z', 'я', 'Z', 'ж', '1', '10', '9', 'é', 'e', 'ea', 'éa', 'Ё', 'ъ', 'ь', 'й', 'и', 'abc', 'абв', 'Abc', 'Абв', 'ab', 'аб', 'x1', 'х1']

export const SHEETS: SheetCells = { O: Object.fromEntries(WORDS.map((w, i) => [`A${i + 1}`, w])) }

/** Letters compared pairwise with the operators. */
const LETTERS = ['a', 'а', 'b', 'Б', 'z', 'я', 'е', 'ё', 'e', 'é', 'й', 'и', 'ь', 'ъ', '1']

const q = (s: string) => `"${s}"`

export function buildProbes(): Probe[] {
  const out: Probe[] = []
  const add = (group: string, formula: string) => out.push({ id: `${group}#${out.filter((p) => p.group === group).length + 1}`, group, formula })
  for (let i = 0; i < LETTERS.length; i++) for (let j = i + 1; j < LETTERS.length; j++) add('compare', `${q(LETTERS[i])}<${q(LETTERS[j])}`)
  for (const [a, b] of [['ё', 'е'], ['Ё', 'ё'], ['é', 'e'], ['Е', 'е'], ['й', 'и'], ['ь', 'ъ'], ['абв', 'АБВ'], ['x1', 'х1']]) add('equal', `${q(a)}=${q(b)}`)
  for (let k = 1; k <= WORDS.length; k++) add('sort', `INDEX(_xlfn._xlws.SORT(O!A1:A${WORDS.length}),${k},1)`)
  for (let k = 1; k <= WORDS.length; k += 3) add('sort', `INDEX(_xlfn._xlws.SORT(O!A1:A${WORDS.length},1,-1),${k},1)`)
  for (const c of ['<м', '<b', '<я', '>z', '<ё', '<=е', '<é', '>=а', '<Б', '>9']) add('criteria', `COUNTIF(O!A1:A${WORDS.length},${q(c)})`)
  for (const x of ['Ё', 'е', 'é', 'E', 'й', 'х1', 'x1', 'АБ']) add('match', `MATCH(${q(x)},O!A1:A${WORDS.length},0)`)
  for (const list of ['{"е";"ё";"Е";"e";"é"}', '{"и";"й";"И"}', '{"x1";"х1"}']) add('unique', `ROWS(_xlfn.UNIQUE(${list}))`)
  return out
}
