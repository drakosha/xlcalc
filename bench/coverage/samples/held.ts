/**
 * Held out after `basic`: text around the cut at 50 characters (with the accounting padding too), numbers whose text
 * runs past 50 characters, the calendar's edges, times and the phantom 29 February, the other boolean and errors, an
 * empty text, a line break, Cyrillic text, the 1904 date system's zero and a negative time there.
 */
import type { SampleCase } from './cases.js'

const text = (n: number) => 'abcdefghij'.repeat(6).slice(0, n)

export function buildSampleCases(): SampleCase[] {
  return [
    { id: 'text-48', value: text(48) },
    { id: 'text-49', value: text(49) },
    { id: 'text-50', value: text(50) },
    { id: 'text-51', value: text(51) },
    { id: 'text-53', value: text(53) },
    { id: 'cyrillic-60', value: 'съешь же ещё этих мягких французских булок, да выпей же чаю!' },
    { id: 'line-break', value: 'one\ntwo' },
    { id: 'empty-text', value: '', formula: '""' },
    { id: 'false', value: false },
    { id: 'na', value: { error: 'NA' }, formula: 'NA()' },
    { id: 'value-error', value: { error: 'VALUE' }, formula: '"a"+1' },
    { id: 'googol', value: 1e100 },
    { id: 'e40', value: 1.5e40 },
    { id: 'e30', value: 1e30 },
    { id: 'max-percent', value: 1e307 },
    { id: 'negative-tiny', value: -1e-10 },
    { id: 'last-day', value: 2_958_465.5 },
    { id: 'last-moment', value: 2_958_465.999_999_9 },
    { id: 'past-last-day', value: 2_958_466 },
    { id: 'phantom-day', value: 60 },
    { id: 'one', value: 1 },
    { id: 'thirds', value: 1 / 3 },
    { id: 'negative-fraction', value: -2.25 },
    { id: 'large-time', value: 1.999_999_99 },
    { id: 'percent-negative', value: -0.125, numberFormat: 10 },
    { id: 'zero-1904', value: 0, dateSystem: 1904 },
    { id: 'negative-time-1904', value: -0.25, dateSystem: 1904 },
    { id: 'past-last-day-1904', value: 2_957_003, dateSystem: 1904 },
  ]
}
