/**
 * Month names of ru-RU in TEXT, held out: a host's books written by Excel for Windows hold `март` where the engine gives
 * `Март` for `TEXT(DATE(,MONTH(R),N),"ММММ")`. Every month in `ММММ`, `МММ` and `МММММ`, after a day (`Д ММММ`, the
 * genitive place), with a year, under `[$-419]` and `[$-FC19]`, and the corpus's own form over a data cell. Computed in
 * Excel ru-RU on Windows and the Mac (`test/excel/win/cases/ru-RU/month-ru-held.json`,
 * `test/excel/cases/ru-RU/month-ru-held.json`); `month-ru-held.ru.predicted.json`, the engine's answers, was committed
 * before Excel ran the set.
 */
import type { Probe, SheetCells } from '../probes.js'
import { textFormatProbes } from './text-cases.js'

const MONTHS = Array.from({ length: 12 }, (_, i) => `DATE(2025,${i + 1},15)`)

const CODES = ['ММММ', 'МММ', 'МММММ', 'Д ММММ', 'ДД ММММ ГГГГ', 'ММММ ГГГГ', '[$-419]ММММ', '[$-419]Д ММММ', '[$-FC19]ММММ', '[$-FC19]Д ММММ ГГГГ']

/** The data sheet of the corpus's form: a date per month in column A. */
export const SHEETS: SheetCells = { M: Object.fromEntries(Array.from({ length: 12 }, (_, i) => [`A${i + 1}`, 45672 + 31 * i])) }

export function buildProbes(): Probe[] {
  const out = textFormatProbes([], [], CODES, MONTHS)
  for (let i = 1; i <= 12; i++) out.push({ id: `corpus#${i}`, group: 'corpus', fn: 'TEXT', formula: `TEXT(DATE(,MONTH(M!A${i}),1),"ММММ")` })
  return out
}
