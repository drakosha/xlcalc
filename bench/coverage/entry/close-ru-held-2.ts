/**
 * Held out after `close-ru-held`, for ru-RU alone, typed on the keyboard into Excel for Mac in the Russian layout
 * (`input-source.swift com.apple.keylayout.Russian` before the run) and predicted by the engine before Excel ran it, with
 * the rule corrected by `close-held`. No Latin letters, which that layout does not type.
 */
import { collector, type EntryCase } from './cases.js'

export const keys = true

const CASES: Readonly<Record<string, readonly string[]>> = {
  ru: ['=-СУММ(1;2', '=СУММ(1;{1;2', '=1+сумм(2', '=ЕСЛИ(1>0;"да', '=СУММ(1;2;', '=МАКС(1;МИН(2', '={1;2'],
}

export function buildEntryCases(): EntryCase[] {
  const { out, add } = collector()
  for (const [group, texts] of Object.entries(CASES)) for (const t of texts) add(group, t)
  return out
}
