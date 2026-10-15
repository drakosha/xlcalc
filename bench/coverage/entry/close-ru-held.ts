/**
 * Held out with `close-held`, for ru-RU alone: formulas with the locale's names of functions and brackets or a quote left
 * open at their end, typed on the keyboard into Excel for Mac in the Russian layout (`input-source.swift
 * com.apple.keylayout.Russian` before the run), entered with Return, and predicted by the engine before Excel ran it.
 * No Latin letters, which that layout does not type (a case `=ABS(-(1)` typed `=ФИЫ(-(1)` and was taken out).
 */
import { collector, type EntryCase } from './cases.js'

export const keys = true

const CASES: Readonly<Record<string, readonly string[]>> = {
  ru: ['=СУММ(1;2', '=сумм(1;2', '=СУММ(МАКС(1;2', '=ЕСЛИ(1>0;СУММ(1', '=1+СУММ(2', '=СУММ(1)+МАКС(2', '=СУММ((1', '=ОКРУГЛ(1,5', '=ОКРУГЛ(1,5;', '=ДЛСТР("аб', '=СУММ({1;2'],
}

export function buildEntryCases(): EntryCase[] {
  const { out, add } = collector()
  for (const [group, texts] of Object.entries(CASES)) for (const t of texts) add(group, t)
  return out
}
