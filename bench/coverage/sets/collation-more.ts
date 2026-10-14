/**
 * The order of text inside Cyrillic and Latin where a later letter could outweigh an accent, a case or `ё`: words that
 * differ in a first letter only by `ё`, `й`, a case or an accent and in a second letter otherwise (`"ёа"<"еб"`), the
 * order of `ъ`, `ы`, `ь`, and SORT over such words. Excel for Windows's answers (`bench/windows`, ru-RU) check the
 * engine's order beyond the pairs of single letters of `collation.ts`. The sheet `W` holds the words.
 */
import type { Probe, SheetCells } from '../probes.js'
import { byGroup } from './arrays/common.js'

const WORDS = ['ёа', 'еб', 'еа', 'ёб', 'Еа', 'ЁА', 'ежик', 'ёжик', 'шел', 'шёл', 'йа', 'иб', 'иа', 'йб', 'ъ', 'ы', 'ь', 'éa', 'eb', 'ea', 'éb', 'Ab', 'aB', 'а1', 'a1']

export const SHEETS: SheetCells = { W: Object.fromEntries(WORDS.map((w, i) => [`A${i + 1}`, w])) }

const q = (s: string) => `"${s}"`

export function buildProbes(): Probe[] {
  const { out, add } = byGroup()
  const pairs = [
    ['ёа', 'еб'], ['еб', 'ёа'], ['еа', 'ёб'], ['ёб', 'еа'], ['Ёа', 'еб'], ['йа', 'иб'], ['иа', 'йб'], ['йб', 'иа'],
    ['éa', 'eb'], ['eb', 'éa'], ['ea', 'éb'], ['аБ', 'Аб'], ['Аб', 'аБ'], ['ёж', 'ежа'], ['ежик', 'ёжик'], ['ёжик', 'ежик'], ['шёл', 'шел'], ['шел', 'шёл'],
    ['ъ', 'ы'], ['ы', 'ь'], ['ь', 'э'], ['э', 'ю'], ['a1', 'а1'], ['Z', 'а'], ['z', 'Ё'], ['ё', 'ж'], ['Ё', 'ж'],
  ]
  for (const [a, b] of pairs) add('compare', `${q(a)}<${q(b)}`)
  for (const [a, b] of [['аб', 'АБ'], ['ёа', 'Ёа'], ['шел', 'шёл'], ['Ab', 'aB']]) add('equal', `${q(a)}=${q(b)}`)
  for (let k = 1; k <= WORDS.length; k++) add('sort', `INDEX(_xlfn._xlws.SORT(W!A1:A${WORDS.length}),${k},1)`)
  for (const c of ['<ёб', '<=еа', '>ёа', '<иб', '>=Ёа', '<ы']) add('criteria', `COUNTIF(W!A1:A${WORDS.length},${q(c)})`)
  return out
}
