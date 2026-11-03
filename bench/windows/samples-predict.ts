/**
 * The samples of the Number Format list as the package would give them before Excel showed a set, written as the
 * predictions of the Windows step `samples`: `npx tsx bench/windows/samples-predict.ts <set>` writes
 * `samples/<set>.predicted.json`, each case's samples in the list's order in en-US and ru-RU. The rule `basic` was
 * predicted with: each item shows the value in its format by `displayValue` without a width, General a number in
 * General and text or an empty cell as "No specific format". The rule after `basic`, which this script writes: General
 * "No specific format" for every cell, an empty cell nothing, a date or time item's text all `#` as ten `#`, a sample
 * past 50 characters cut to 50 and "...".
 */
import { displayValue, numberFormatChoices } from '../../src/index.js'
// the package's locales loaded
import '../coverage/locales.js'
import { setModule, writePrediction } from '../coverage/stored.js'
import type { SampleCase, SampleSet } from '../coverage/samples/cases.js'

const NO_FORMAT: Record<string, string> = { 'en-US': 'No specific format', 'ru-RU': 'Формат не указан' }

const DATE_ITEMS = new Set(['shortDate', 'longDate', 'time'])

function predicted(c: SampleCase, tag: string): string[] {
  return numberFormatChoices(tag).map(({ key, format }) => {
    if (key === 'general') return NO_FORMAT[tag]
    if (c.value === null) return ''
    const text = displayValue(c.value, { numberFormat: format, locale: tag, ...(c.dateSystem ? { dateSystem: c.dateSystem } : {}) }).text
    if (DATE_ITEMS.has(key) && /^#+$/.test(text)) return '#'.repeat(10)
    return text.length > 50 ? `${text.slice(0, 50)}...` : text
  })
}

const set = process.argv[2]
const cases = (await setModule<SampleSet>('samples', set)).buildSampleCases()
const locales = Object.fromEntries(Object.keys(NO_FORMAT).map((tag) => [tag, Object.fromEntries(cases.map((c) => [c.id, predicted(c, tag)]))]))
writePrediction(import.meta.url, `samples/${set}`, { now: new Date().toISOString().slice(0, 10), keys: numberFormatChoices().map((c) => c.key), locales })
