/**
 * The samples of the Number Format list as the package would give them before Excel showed any, written as the
 * predictions of the Windows step `samples`: `npx tsx bench/windows/samples-predict.ts <set>` writes
 * `samples/<set>.predicted.json`, each case's samples in the list's order in en-US and ru-RU. The rule predicted: each
 * item shows the value in the item's format by `displayValue` without a width; General a number in General and text or
 * an empty cell as "No specific format", as the list showed for a text cell.
 */
import { displayValue, numberFormatChoices } from '../../src/index.js'
// the package's locales loaded
import '../coverage/locales.js'
import { setModule, writePrediction } from '../coverage/stored.js'
import type { SampleCase, SampleSet } from '../coverage/samples/cases.js'

const NO_FORMAT: Record<string, string> = { 'en-US': 'No specific format', 'ru-RU': 'Формат не указан' }

function predicted(c: SampleCase, tag: string): string[] {
  return numberFormatChoices(tag).map(({ key, format }) => {
    if (key === 'general' && (c.value === null || typeof c.value === 'string')) return NO_FORMAT[tag]
    if (c.value === null) return ''
    return displayValue(c.value, { numberFormat: key === 'general' ? undefined : format, locale: tag, ...(c.dateSystem ? { dateSystem: c.dateSystem } : {}) }).text
  })
}

const set = process.argv[2]
const cases = (await setModule<SampleSet>('samples', set)).buildSampleCases()
const locales = Object.fromEntries(Object.keys(NO_FORMAT).map((tag) => [tag, Object.fromEntries(cases.map((c) => [c.id, predicted(c, tag)]))]))
writePrediction(import.meta.url, `samples/${set}`, { now: new Date().toISOString().slice(0, 10), keys: numberFormatChoices().map((c) => c.key), locales })
