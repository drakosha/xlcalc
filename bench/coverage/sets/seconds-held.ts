/**
 * Held out: the time of day of a serial number at the edge of a second, where the engine had two rules. DAY, MONTH and
 * YEAR rounded the whole number to the second (`dayAt`: `round(t*86400)`), HOUR, MINUTE, SECOND and number formats the
 * fraction of the day (`round((t-INT(t))*86400)`); in floating point they part where the fraction is within a few units
 * of the last place of half a second, and before midnight one of them can go on to the next day. The values of sheet T
 * are such numbers: fractions of a second near a half by decimals (0.4994 and 0.4996 tell rounding through the
 * milliseconds), the doubles nearest a half where the two rules, a rounding to 15 digits and adding half a second part,
 * exact halves (1/256 of a day is 337.5 seconds), mid-day and before midnight, before a month's, a year's and the
 * calendar's end, near 1900-02-29 and under 1. Each value is read by DAY, MONTH, YEAR, HOUR, MINUTE, SECOND, TEXT with a
 * date and time, elapsed hours and milliseconds, and INT; `load` gives the fraction of its second, so the stored value
 * is the one Excel loaded. Answers: `test/excel/cases/seconds-held.json`; `PREDICTED`, the engine's answers, was
 * committed before Excel ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'

/** A moment of 2025-01-31: 10:20:30 and a fraction of a second, or 23:59:59 and one. */
const mid = (f: number) => 45688 + (37230 + f) / 86400
const eve = (f: number) => 45688 + (86399 + f) / 86400

/** The values of T!A1 down. */
export const VALUES: readonly number[] = [
  // 10:20:30 and 0.4 to 0.5001 of a second
  mid(0.4), mid(0.499), mid(0.4994), mid(0.4996), mid(0.4999), mid(0.49999),
  // the doubles nearest 10:20:30.5: the whole number rounds up from the second, then the fraction too, 15 digits not yet
  45688.4309085648, 45688.43090856481, 45688.43090856482, mid(0.5001),
  // within the day at other magnitudes: the whole number up, the fraction down (444615…), and 15 digits or half a second either way
  444615.13495949074, 2504553.1110706017, 2577729.1113368054, 1963420.4880381944, 2693171.0251446757, 2072685.826105324,
  // exact halves: 00:05:37.5, 00:16:52.5, 23:54:22.5
  45688 + 1 / 256, 45688 + 3 / 256, 45688 + 255 / 256,
  // before midnight: 23:59:59 and 0.4994, 0.4996, a double below half, a double above it
  eve(0.4994), eve(0.4996), 45688.99999421295, 45688.999994212965,
  // before a year's end, 1900-02-29 (serial 60) and 1 (serial 0 under it)
  46022.999994210055, 46022.99999421295, 59.99999421296295, 59.99999421296296, 0.9999942129629625,
  // larger serials before midnight: the whole number goes on to the next day, the fraction does not, or 15 digits part
  999999.9999942129, 1020183.9999942129, 1000000.999994213,
  // before 9999-12-31 and before the calendar's end
  2958464.9999942128, 2958464.999994213, 2958465.999994213, 2958465.9999942128,
]

export const SHEETS: SheetCells = { T: Object.fromEntries(VALUES.map((v, i) => [`A${i + 1}`, v])) }

/** The fraction of a second each value holds, as `load` reads it: the seconds of the day less the whole ones. */
const wholeSeconds = (t: number) => Math.floor((t - Math.floor(t)) * 86400)

const READS: [group: string, fn: string, formula: (at: string, t: number) => string][] = [
  ['DAY', 'DAY', (x) => `DAY(${x})`],
  ['MONTH', 'MONTH', (x) => `MONTH(${x})`],
  ['YEAR', 'YEAR', (x) => `YEAR(${x})`],
  ['HOUR', 'HOUR', (x) => `HOUR(${x})`],
  ['MINUTE', 'MINUTE', (x) => `MINUTE(${x})`],
  ['SECOND', 'SECOND', (x) => `SECOND(${x})`],
  ['TEXT:moment', 'TEXT', (x) => `TEXT(${x},"yyyy-mm-dd hh:mm:ss")`],
  ['TEXT:day', 'TEXT', (x) => `TEXT(${x},"d h:mm:ss")`],
  ['TEXT:elapsed', 'TEXT', (x) => `TEXT(${x},"[h]:mm:ss")`],
  ['TEXT:ms', 'TEXT', (x) => `TEXT(${x},"ss.000")`],
  ['INT', 'INT', (x) => `INT(${x})`],
  ['load', 'INT', (x, t) => `(${x}-INT(${x}))*86400-${wholeSeconds(t)}`],
]

/** Case `<group>#<n>` reads T!A<n>. */
export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [group, fn, formula] of READS) VALUES.forEach((t, i) => add(group, fn, formula(`T!A${i + 1}`, t)))
  return out
}

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./seconds-held.predicted.json', import.meta.url), 'utf8'))

const of = (group: string, ns: number[]) => ns.map((n) => `${group}#${n}`)
const TIMED = [11, 12, 13, 15, 16, 17, 18, 19, 27, 32, 35]

/**
 * The predictions Excel did not confirm, the Mac and Windows alike: Excel adds half a second and cuts the sum to the
 * second (an exact half stays below, 37.5 s is 37), where the engine rounded the fraction half up in HOUR and the formats
 * and the whole number in DAY, which part on the doubles near a half; and a moment that rounds past 9999-12-31 is #NUM!,
 * in TEXT #VALUE!.
 */
export const MISSED: readonly string[] = [
  ...of('DAY', [29, 30, 32, 34, 35]), ...of('MONTH', [30, 34, 35]), ...of('YEAR', [34, 35]), ...of('HOUR', [27, 32, 34, 35]),
  ...of('MINUTE', [27, 32, 34, 35]), ...of('SECOND', [11, 12, 13, 15, 16, 17, 18, 19, 27, 32, 34, 35]),
  ...of('TEXT:moment', TIMED), ...of('TEXT:day', TIMED), ...of('TEXT:elapsed', TIMED),
]
