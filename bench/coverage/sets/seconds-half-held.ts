/**
 * Held out: moments exactly half a second past a whole second, where two readings of Excel's time of day part. Both add
 * half a second and take the day; one divides the rest of the day by a second's length (the engine's reading through
 * 0.90.0), the other scales it to hours, minutes and seconds and floors each step (the engine's reading since three
 * cases of `SECOND`: 0.5+0.5/86400 is 0, 0.5+9.5/86400 is 9, 10.5/86400 is 11). They disagree only at small serials,
 * a time of day or a duration of a few days: on 57 % of the half seconds under 1, 13 % on serials 1 and 2, 4 % on 10,
 * 1 % on 60, and on none from 2025-01-31 up, even three units of the last place either side. The values of sheet T: on serials
 * 0, 1, 2, 10 and 60 and on 2025-01-31 the first, noon's and the last half second, twelve moments where the readings
 * disagree (none on the last day) and twelve where they agree, spread evenly over the day. Each value is read by HOUR,
 * MINUTE, SECOND, DAY, TEXT with a time, and `load`, the fraction of its second as Excel loaded it. Answers:
 * `test/excel/cases/seconds-half-held.json`; `PREDICTED`, the engine's answers, was committed before Excel ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe, SheetCells } from '../probes.js'
import { collector } from './arrays/common.js'

const SECOND = 1 / 86_400
const DAYS = [0, 1, 2, 10, 60, 45688]
const EACH = 12

/** The second of the day each reading gives, a whole day as 86400. */
function bySecondLength(t: number): number {
  const rounded = t + SECOND / 2
  return Math.floor((rounded - Math.floor(rounded)) / SECOND)
}

function byHourMinuteSecond(t: number): number {
  const rounded = t + SECOND / 2
  const hours = (rounded - Math.floor(rounded)) * 24
  const minutes = (hours % 1) * 60
  return Math.floor(hours) * 3600 + Math.floor(minutes) * 60 + Math.floor((minutes % 1) * 60)
}

/** `count` of `items`, evenly spaced from the first. */
const spread = <T>(items: readonly T[], count: number): T[] => Array.from({ length: Math.min(count, items.length) }, (_, i) => items[Math.floor((i * items.length) / count)])

function valuesOfDay(day: number): number[] {
  const moment = (second: number) => day + (second + 0.5) / 86_400
  const parted: number[] = []
  const agreed: number[] = []
  for (let second = 1; second < 86_399; second++) {
    if (second === 43_200) continue
    const t = moment(second)
    ;(bySecondLength(t) === byHourMinuteSecond(t) ? agreed : parted).push(t)
  }
  return [moment(0), moment(43_200), moment(86_399), ...spread(parted, EACH), ...spread(agreed, EACH)]
}

/** The values of T!A1 down. */
export const VALUES: readonly number[] = DAYS.flatMap(valuesOfDay)

export const SHEETS: SheetCells = { T: Object.fromEntries(VALUES.map((v, i) => [`A${i + 1}`, v])) }

/** The seconds of the day less the whole ones, as `load` reads it. */
const wholeSeconds = (t: number) => Math.floor((t - Math.floor(t)) * 86_400)

const READS: [group: string, fn: string, formula: (at: string, t: number) => string][] = [
  ['HOUR', 'HOUR', (x) => `HOUR(${x})`],
  ['MINUTE', 'MINUTE', (x) => `MINUTE(${x})`],
  ['SECOND', 'SECOND', (x) => `SECOND(${x})`],
  ['DAY', 'DAY', (x) => `DAY(${x})`],
  ['TEXT', 'TEXT', (x) => `TEXT(${x},"hh:mm:ss")`],
  ['load', 'INT', (x, t) => `(${x}-INT(${x}))*86400-${wholeSeconds(t)}`],
]

/** Case `<group>#<n>` reads T!A<n>. */
export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const [group, fn, formula] of READS) VALUES.forEach((t, i) => add(group, fn, formula(`T!A${i + 1}`, t)))
  return out
}

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./seconds-half-held.predicted.json', import.meta.url), 'utf8'))
