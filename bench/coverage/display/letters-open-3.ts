/**
 * Files opened with three conditions in the shape of a code seen in a saved book, four sections with a colour in each: with colours, without
 * them, and in three sections; and two conditions with a colour and a third section with none, which should open. Predicted before Excel for
 * Windows ran the set: a condition past the second section refuses the file, as in `letters-open`, thresholds and digits being our own.
 */
import type { DisplayCase } from './cases.js'

const OPENS: readonly [string, string][] = [['[>=1][Green]0.0%;[>=0.8][Blue]0.0%;[Red]0.0%', '4609650.0%']]

export const REFUSED_PREDICTED: readonly string[] = ['[>=1][Green]0.0%;[>=0.8][Blue]0.0%;[>=0.5][Magenta]0.0%;[Red]0.0%', '[>=1]0.0%;[>=0.8]0.0%;[>=0.5]0.0%;0.0%', '[>=1][Green]0.0%;[>=0.8][Blue]0.0%;[>=0.5][Magenta]0.0%']

export function buildDisplayCases(): DisplayCase[] {
  const open: DisplayCase[] = OPENS.map(([format, text], i) => ({ id: `o${i + 1}`, cell: { value: 46096.5, format }, width: 40, predicted: { text } }))
  return [...open, ...REFUSED_PREDICTED.map((format, i) => ({ id: `r${i + 1}`, cell: { value: 46096.5, format }, width: 40 }))]
}
