/**
 * Files opened with the codes of `links/letters-text`: `_` and `*` taking a section's `;`, conditions with a fourth section, the letters
 * that read as codes, and the codes TEXT refuses. Whether a file opens is another reading than TEXT's. Predicted before Excel for Windows
 * ran the set: every code TEXT refuses refuses the file too (`REFUSED_PREDICTED`), the others open and show what Excel for Mac's cell showed
 * (a fill's width is left out). A refused book tells nothing of the Mac's display step, which dies on it.
 */
import type { DisplayCase } from './cases.js'

const OPENS: readonly [string, string?][] = [
  ["[Red]\\+#,##0_ ;[Green]\\-#,##0_;;", "+46,097 "],
  ["[RED]\\+#,##0_ ;[GREEN]\\-#,##0_;;;", "+46,097 "],
  ["0_;;0", "46097 "],
  ["0*;;0"],
  ["_;0", " 46097"],
  ["0_;", "46097 "],
  ["[>=1][Green]0.0%;[>=0.8][Blue]0.0%;0.0%", "4609650.0%"],
  ["[>=1][Green]0.0%;[>=0.8][Blue]0.0%;0.0%;@", "4609650.0%"],
  ["[>1]0;0;0;@", "46097"],
  ["[>1]0;[>2]0;0;@", "46097"],
  ["0;[>1]0;[>2]0", "46097"],
  ["0;0;[>1]0", "46097"],
  ["aaa", "Sun"],
  ["aaaa", "Sunday"],
  ["AAA", "Sun"],
  ["AAAA", "Sunday"],
  ["b", "69"],
  ["bb", "69"],
  ["BB", "69"],
  ["bbb", "2569"],
  ["BBB", "2569"],
  ["g", ""],
  ["gg", ""],
  ["ggg", ""],
  ["G", ""],
  ["GG", ""],
  ["B1", ""],
  ["B2", ""],
  ["b1", ""],
  ["B1yy", "26"],
  ["yyaaa", "26Sun"],
  ["yyb", "2669"],
  ["yyg", "26"],
  ["aa", "aa"],
  ["0aa", "46097aa"],
  ["aa0", "aa46097"],
  ["r", "r"],
  ["\"n\"", "n"],
  ["\\n", "n"],
  ["yy\"n\"", "26n"],
  ["yy\\n", "26n"],
  ["General", "46096.5"],
  ["e", "2026"],
  ["ee", "2026"],
  ["blue", "69lu2026"],
]

export const REFUSED_PREDICTED: readonly string[] = ["[>=1][Green]0.0%;[>=0.8][Blue]0.0%;[>=0.5]0.0%", "[>1]0;[>2]0;[>3]0;@", "0;0;0;@;", "0;0;0;0;0", "0_", "0*", "n", "N", "nn", "NN", "nnn", "dn", "yn", "n0", "0n", "en", "ne", "n1", "N1", "yyn", "B", "yyB", "0B", "0b", "0g", "0G", "0aaa", "aaa0", "0e", "0;0;0;aaa", "0;0;0;n", "E+0", "e+0", "0e+0", "yyE", "eE"]

export function buildDisplayCases(): DisplayCase[] {
  const open: DisplayCase[] = OPENS.map(([format, text], i) => ({ id: `o${i + 1}`, cell: { value: 46096.5, format }, width: 40, ...(text === undefined ? {} : { predicted: { text } }) }))
  return [...open, ...REFUSED_PREDICTED.map((format, i) => ({ id: `r${i + 1}`, cell: { value: 46096.5, format }, width: 40 }))]
}
