/** What a number's format shows it as: a date, a time of day alone or a number (`dayOrTime` of its first section). */
import type { CellInput } from '../contract.js'
import { enCode } from './builtin.js'
import { dayOrTime } from './date.js'
import { EN_CODES, parseFormat } from './parse.js'

/** What a number's format shows it as, by its first section: a date (a day, a month or a year in it), a time of day alone, or a number. */
export function shownAs(format: CellInput['numberFormat']): 'date' | 'time' | 'number' {
  if (format === undefined) return 'number'
  return dayOrTime(parseFormat(enCode(format), EN_CODES)?.[0]?.toks ?? []) ?? 'number'
}
