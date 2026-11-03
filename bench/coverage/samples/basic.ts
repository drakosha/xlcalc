/**
 * The Number Format list's samples over one cell of each kind of value: numbers of each size and sign, dates and times,
 * a percent, text, text of digits, a boolean, an error, an empty cell; cells already in a format other than General
 * (whether the General item shows the value in General or in the cell's format); the same numbers in a narrow and a
 * wide column (whether a sample depends on the column's width); a book of the 1904 date system.
 */
import type { SampleCase } from './cases.js'

export function buildSampleCases(): SampleCase[] {
  return [
    { id: 'fraction', value: 1234.5 },
    { id: 'integer', value: 42 },
    { id: 'negative', value: -1234.5 },
    { id: 'zero', value: 0 },
    { id: 'huge', value: 1.23456789e20 },
    { id: 'tiny', value: 0.000012345 },
    { id: 'long-digits', value: 3.14159265358979 },
    { id: 'negative-small', value: -0.5 },
    { id: 'past-calendar', value: 3_000_000 },
    { id: 'date', value: 46_000, numberFormat: 14 },
    { id: 'date-time', value: 46_000.75, numberFormat: 22 },
    { id: 'time', value: 0.6, numberFormat: 21 },
    { id: 'percent', value: 0.25, numberFormat: 9 },
    { id: 'text', value: 'East' },
    { id: 'text-digits', value: '123' },
    { id: 'long-text', value: 'A text longer than the list is wide, to see whether a sample is cut' },
    { id: 'true', value: true },
    { id: 'error', value: { error: 'DIV0' }, formula: '1/0' },
    { id: 'empty', value: null },
    { id: 'empty-formatted', value: null, numberFormat: '0.00' },
    { id: 'decimals-3', value: 1234.5, numberFormat: '0.000' },
    { id: 'currency-cell', value: 1234.5, numberFormat: 7 },
    { id: 'custom-cell', value: 1234.5, numberFormat: '"x"0.0"y"' },
    { id: 'text-format-number', value: 1234.5, numberFormat: 49 },
    { id: 'text-format-text', value: 'East', numberFormat: 49 },
    { id: 'percent-format-text', value: 'East', numberFormat: 10 },
    { id: 'narrow', value: 1234.5, width: 2 },
    { id: 'narrow-general', value: 123_456_789_012, width: 4 },
    { id: 'wide', value: 1234.5, width: 40 },
    { id: 'wide-general', value: 123_456_789_012, width: 40 },
    { id: 'date-1904', value: 46_000, numberFormat: 14, dateSystem: 1904 },
    { id: 'negative-1904', value: -1234.5, dateSystem: 1904 },
  ]
}
