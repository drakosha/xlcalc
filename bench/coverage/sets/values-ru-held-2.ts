/**
 * DOLLAR at the 255 characters FIXED and TEXT keep, held out in Excel ru-RU alone (`probes values-ru-held-2 ru`): the
 * ruble format writes a negative amount with a minus and, as the engine takes it, pads nothing. Answers:
 * `test/excel/cases/ru-RU/values-ru-held-2.json`; the engine's answers, `values-ru-held-2.ru.predicted.json`, were
 * committed before Excel ran the set.
 */
import { heldCases } from './arrays/common.js'

export const { buildProbes } = heldCases([
  ['DOLLAR', 'LEN(DOLLAR(1E+189,0))'],
  ['DOLLAR', 'DOLLAR(1E+190,0)'],
  ['DOLLAR', 'LEN(DOLLAR(-1E+187,1))'],
  ['DOLLAR', 'DOLLAR(-1E+188,1)'],
  ['FIXED', 'LEN(FIXED(1E+191,0))'],
  ['FIXED', 'FIXED(1E+192,0)'],
])
