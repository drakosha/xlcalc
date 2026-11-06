/**
 * TEXT with what `letters-held` left open, each case a file of its own: where `B` refuses, digits beside date codes, runs of `a`, `m` across
 * the new codes, conditions by their section, the fourth section. Predicted before Excel ran the set, by the rule as those answers left it;
 * null is #VALUE!. The value is 46096.5 (15 March 2026, 12:00) without one.
 */
import type { LinksCases } from './cases.js'
import { textCase } from './letters-held.js'

export const MORE: readonly [string, string | null, string?][] = [
  // `B` reads as `b` but at the end of its section after no other `B`
  ["bBB", "2569"],
  ["BBb", "2569"],
  ["bBbB", null],
  ["yyBb", "2669"],
  ["bByy", "6926"],
  ["B\"x\"", "69x"],
  ["\"x\"B", null],
  ["B yy", "69 26"],
  ["yyB yy", "2669 26"],
  ["B ", "69 "],
  [" B", null],
  ["B;0", null, "-1"],
  ["B.", "69."],
  ["B/", "69/"],
  ["B\\n", "69n"],
  ["B_x", "69 "],
  ["B*x", "69"],
  ["[Red]B", null],
  // a digit other than a placeholder
  ["yy3", null],
  ["d3", null],
  ["3yy", null],
  ["dd1", null],
  ["h1", null],
  ["aaa3", null],
  ["g3", null],
  ["e3", null],
  ["yy 3", null],
  ["b 3", null],
  ["1", "1"],
  ["5", "5"],
  ["0 1", "46097 1"],
  ["1 0", "1 46097"],
  // runs of `a` are of one case
  ["aaaA", "SunA"],
  ["Aaaa", "ASun"],
  ["aaaAAA", "SunSun"],
  ["aaaAA", "SunAA"],
  ["aAAA", "aSun"],
  // `m` after hours or before seconds, across the other codes
  ["h dd mm", "12 15 03"],
  ["h yy mm", "12 26 03"],
  ["h e mm", "12 2026 03"],
  ["h g mm", "12  00"],
  ["h ddd mm", "12 Sun 03"],
  ["mm aaa ss", "00 Sun 00"],
  ["mm bb ss", "00 69 00"],
  ["mm g ss", "00  00"],
  ["mm dd ss", "03 15 00"],
  // conditions in the first two sections alone, the second after the first
  ["0;[>1]0", null, "-1"],
  ["0;[>1]0;0", null],
  ["[>1]0;0;[>2]0", null],
  ["[>1]0;[>2]0;@", "46097"],
  ["0;0;[>1]0;@", null],
  ["[>1]0;0;0;[>2]@", null],
  ["0;0;0;[>1]@", null],
  ["[Red]0;[>1]0", null, "-1"],
  // a fourth section
  ["0;0;0;[h]", null],
  ["0;0;0;[Red]", "46097"],
  ["0;0;0;%", null],
  ["0;0;0;#", null],
  ["0;0;0;?", null],
  ["0;0;0;s", null],
  ["0;0;0;B", null],
  ["0;0;0;aa@", "46097"],
  // an exponent
  [".0E+0", ".5E+5"],
  ["0.0E+0", "4.6E+4"],
]

export function buildLinksCases(): LinksCases {
  return { cases: MORE.map(([code, text, value], i) => textCase(`m${i + 1}`, code, text, value)) }
}
