/**
 * The date codes `e` and `ee`: [MS-OI29500] 2.1.713 (f) says they become `yyyy` in Office, in all locales but Japanese and Chinese (Taiwan),
 * where the standard says `yy`. The longer runs and upper case, which the notes do not name, are predicted the same. Predicted before
 * Excel ran the set.
 */
import type { DisplayCase } from './cases.js'

const VALUE = 46096.5

export function buildDisplayCases(): DisplayCase[] {
  const cases: [string, string, number?][] = [
    ['e', '2026'],
    ['ee', '2026'],
    ['eee', '2026'],
    ['eeee', '2026'],
    ['E', '2026'],
    ['EE', '2026'],
    ['e/m/d', '2026/3/15'],
    ['ee-mm-dd', '2026-03-15'],
    ['dd.mm.e', '15.03.2026'],
    ['e "г."', '2026 г.'],
    ['e h:mm', '2026 12:00'],
    ['e', '1900', 1],
  ]
  return cases.map(([format, text, value], i) => ({ id: `e${i + 1}`, cell: { value: value ?? VALUE, format }, width: 40, predicted: { text } }))
}
