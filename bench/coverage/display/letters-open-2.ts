/**
 * Files opened with the codes the sets `links/letters-held` and `links/letters-more` answered otherwise than their rule, and with codes that
 * fail whole. In `letters-open` a file opened exactly where TEXT took the code, on all 81 codes: predicted the same before Excel for Windows
 * ran this set, with TEXT's answers as the text.
 */
import type { DisplayCase } from './cases.js'

const OPENS: readonly [string, string][] = [
  ["Byy", "6926"],
  ["Blue", "69lu2026"],
  ["Bb", "69"],
  ["bBb", "2569"],
  ["B;0", "69"],
  ["B\"x\"", "69x"],
  ["aAa", "aAa"],
  ["aaa/p", "aap"],
  ["aaaAAA", "Sunday"],
  ["h bb mm", "12 69 00"],
  ["h dd mm", "12 15 00"],
  ["0;[>1]0;0", "46097"],
  ["0;[>1]0", "46097"],
  ["0;0;0;General", "46097"],
  ["0;0;0;AM/PM", "46097"],
  ["0;0;0;[Red]", "46097"],
  [".E+0", ".E+1"],
  [".0E+0", ".5E+5"],
]

export const REFUSED_PREDICTED: readonly string[] = ["bB", "BbB", "b3", "yyb1", "B1B2", "B10", "B3", "yy3", "h1", "3yy", "g3", "e3", "[>1]0;0;[>2]0", "0;0;[>1]0;@", "0;0;0;[>1]@", "0;0;0;[h]", "0;0;0;%", "0;0;0;s", "[Red][Blue]0", "n;0", "0;n", "0b;0", "aaa0;0", "mm###;0", "0;0;n", "\"x\"B"]

export function buildDisplayCases(): DisplayCase[] {
  const open: DisplayCase[] = OPENS.map(([format, text], i) => ({ id: `o${i + 1}`, cell: { value: 46096.5, format }, width: 40, predicted: { text } }))
  return [...open, ...REFUSED_PREDICTED.map((format, i) => ({ id: `r${i + 1}`, cell: { value: 46096.5, format }, width: 40 }))]
}
