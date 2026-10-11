/**
 * Held out, after the set `name-case-held`: FALSE typed as a value with the long s (`falſe`), which the engine takes as
 * text by the rule Excel for Windows showed in formulas (`wordKey`), beside FALSE in lower case. Predicted by the engine
 * before Excel ran it.
 */
import { collector, type EntryCase } from './cases.js'

export function buildEntryCases(): EntryCase[] {
  const { out, add } = collector()
  for (const t of ['falſe', 'FALſE', 'false', 'trUE']) add('value', t)
  return out
}
