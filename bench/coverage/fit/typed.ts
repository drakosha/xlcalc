/**
 * Texts typed into a column, by the object model (`Value`, `Formula2Local`) and on the keyboard: dates, money, percent,
 * grouped and General numbers, a format set before, formulas, text, in a column of the default width, one set by hand,
 * one AutoFit narrowed, and in turn into one column (does a column widen again, does it narrow).
 */
import type { NumberFormat } from '../../../src/contract.js'
import { collector, type FitCase, type Via } from './cases.js'

/** A typed text in each locale. */
type Texts = { en: string; ru: string }

const DATE: Texts = { en: '3/15/2026', ru: '15.03.2026' }
const DATE_LATE: Texts = { en: '12/31/2026', ru: '31.12.2026' }

/** One text into row 1 of a column of the default width, a cell in General or with a format set before. */
const SINGLE: [string, Texts, NumberFormat?][] = [
  ['date', DATE],
  ['date', { en: '15-Mar-2026', ru: '15-мар-2026' }],
  ['date', DATE, 'dddd, mmmm d, yyyy'],
  ['date', { en: '3/15/2026 18:30', ru: '15.03.2026 18:30' }],
  ['money', { en: '$1,234,567.50', ru: '1 234 567,50 ₽' }],
  ['money', { en: '$5', ru: '5 ₽' }],
  ['pct', { en: '12345678%', ru: '12345678%' }],
  ['group', { en: '1,234,567,890', ru: '1 234 567 890' }],
  ['general', { en: '1234567890', ru: '1234567890' }],
  ['general', { en: '12345678901', ru: '12345678901' }],
  ['general', { en: '123456789012', ru: '123456789012' }],
  ['general', { en: '1234567', ru: '1234567' }],
  ['general', { en: '-123456789', ru: '-123456789' }],
  ['general', { en: '3.14159265358979', ru: '3,14159265358979' }],
  ['general', { en: '0.000001234', ru: '0,000001234' }],
  ['sci', { en: '1.5E+300', ru: '1,5E+300' }],
  ['formatted', { en: '1234567', ru: '1234567' }, '#,##0.00'],
  ['formatted', { en: '1234567', ru: '1234567' }, '0.00%'],
  ['formula', { en: '=DATE(2026,3,15)', ru: '=ДАТА(2026;3;15)' }],
  ['formula', { en: '=1234567890*10', ru: '=1234567890*10' }],
  ['formula', { en: '=1234567', ru: '=1234567' }, '#,##0.00'],
  ['text', { en: 'abcdefghijklmnopqrstuvwxyz', ru: 'abcdefghijklmnopqrstuvwxyz' }],
  ['text', { en: 'TRUE', ru: 'ИСТИНА' }],
]

/** Typed only on the keyboard's layout (ASCII): the others are typed through the object model alone. */
const ascii = (t: string) => /^[\x20-\x7e]*$/.test(t)

export function buildFitCases(tag: string): FitCase[] {
  const { out, add } = collector()
  const lang = tag === 'ru-RU' ? 'ru' : 'en'
  for (const via of ['value', 'local', 'keys'] as Via[]) {
    const typed = (group: string, c: Omit<FitCase, 'id' | 'via'>) => {
      if (via === 'keys' && !c.typed!.every(([, t]) => ascii(t))) return
      add(`${via}-${group}`, { ...c, via })
    }
    for (const [group, texts, format] of SINGLE) typed(group, { typed: [[1, texts[lang]]], ...(format !== undefined ? { cells: { 1: { format } } } : {}) })
    // a width set by hand stays, whether narrower or wider
    typed('hand', { typed: [[1, DATE[lang]]], column: { width: 5 } })
    typed('hand', { typed: [[1, DATE[lang]]], column: { width: 8.83203125 } })
    // AutoFit narrowed the column before
    typed('fitted', { cells: { 2: { value: 'ab' } }, fitFirst: true, typed: [[1, DATE[lang]]] })
    typed('fitted', { cells: { 2: { value: 'a long text in the column' } }, fitFirst: true, typed: [[1, DATE[lang]]] })
    // in turn into one column: wider again, then a shorter date, then a long date into a formatted cell
    typed('turn', { cells: { 3: { format: 'dddd, mmmm d, yyyy' } }, typed: [[1, DATE[lang]], [2, '1'], [3, DATE_LATE[lang]], [4, '5']] })
    typed('turn', { typed: [[1, '123456789'], [2, DATE[lang]]] })
  }
  return out
}
