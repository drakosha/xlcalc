/**
 * TEXT on Excel for Mac with a sample of the codes only Excel for Windows answered in `letters-held` and `letters-more`, each a file of its
 * own: `B` inside and at the end of a code, runs of `a`, `m` across other codes, conditions by their section, digits beside dates. Windows's
 * answers are the prediction, committed before Excel for Mac ran the set; null is #VALUE!.
 */
import type { LinksCases } from './cases.js'
import { textCase } from './letters-held.js'

export const SAMPLE: readonly [string, string | null, string?][] = [
  ["bB", null],
  ["Bb", "69"],
  ["BbB", null],
  ["bBb", "2569"],
  ["Byy", "6926"],
  ["Blue", "69lu2026"],
  ["b3", null],
  ["yyb1", null],
  ["B1B2", null],
  ["B10", null],
  ["aAa", "aAa"],
  ["aaaA", "SunA"],
  ["aaaAAA", "Sunday"],
  ["a/p", "p"],
  ["aaa/p", "aap"],
  ["aaaa/p", "Sunp"],
  ["h bb mm", "12 69 00"],
  ["h dd mm", "12 15 00"],
  ["mm aaa ss", "03 Sun 00"],
  ["mm g ss", "00  00"],
  ["0;0;0;General", "46097"],
  ["0;0;0;AM/PM", "46097"],
  ["0;0;0;d", null],
  ["0;[>1]0;0", "46097"],
  ["0;[>1]0;[>2]0", null],
  ["0;0;[>1]0", null],
  ["[>1]0;[>2]0;0;@", "46097"],
  [".E+0", ".E+1"],
  ["yy3", null],
  ["dd1", null],
  ["B;0", "1", "-1"],
  ["0b;0", null, "-1"],
  ["mm###;0", null, "-1"],
  ["0;[>1]0", null, "-1"],
]

export function buildLinksCases(): LinksCases {
  return { cases: SAMPLE.map(([code, text, value], i) => textCase(`s${i + 1}`, code, text, value)) }
}
