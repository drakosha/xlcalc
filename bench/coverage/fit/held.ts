/**
 * Held out from the sets autofit, typed and glyphs, predicted by the engine after their rules were taken in and before
 * Excel ran it: text about 90 points wide, where a thirtieth more of it starts to count, text outside ASCII, numbers in
 * General near their own text and in scientific form, formats with paddings and blanks, the texts of booleans and
 * errors, cells in hidden rows and merged areas; numbers typed into columns of the default width and fitted ones, which
 * fit exactly, drop decimals or show `#`.
 */
import type { CellInput, NumberFormat } from '../../../src/contract.js'
import { textOf } from './autofit.js'
import { collector, HIDDEN, type FitCase, type Via } from './cases.js'

const one = (v: CellInput['value'], format?: NumberFormat): Record<number, CellInput> => ({ 1: { value: v, ...(format !== undefined ? { format } : {}) } })
const formula = (f: string, format?: NumberFormat): Record<number, CellInput> => ({ 1: { formula: f, ...(format !== undefined ? { format } : {}) } })

const TEXTS = ['Quarterly revenue by region', 'Total (excl. VAT), USD', 'café au lait', 'naïve résumé', '€ 1,234.50', 'Привет, мир!', 'Итого по разделу 1', 'Москва', 'ёлка и ЁЛКА', 'Сумма к оплате, ₽', 'Mixed: Цена 100%', 'x']

const NUMBERS: [number, NumberFormat | undefined][] = [
  [-1.5e20, undefined],
  [0.0000123456789, undefined],
  [0.0001, undefined],
  [1e15, undefined],
  [123456789.5, undefined],
  [-0.000012345, undefined],
  [98765.4321, undefined],
  [1234567.5, '#,##0.00 "₽"'],
  [0.1234, '0.0%'],
  [46096, 'dd.mm.yyyy'],
  [46096.5208, 'hh:mm:ss'],
  [1.75, '[h]:mm:ss'],
  [-98765.4, '#,##0.0_);(#,##0.0)'],
  [-0.000123, '0.00E+00'],
  [3.0625, '# ??/??'],
  [3, '# ??/??'],
  [1234.5, 41],
  [-1234.5, 42],
  [0, 43],
  [0, '0.00_);[Red](0.00);"zero"_)'],
  [1234, '#,##0_W'],
  [46096, 'mmmm d, yyyy'],
]

/** Typed texts in each locale, into row 1 of a column of the default width unless a case says otherwise. */
type Texts = { en: string; ru: string }
const TYPED: [string, Texts][] = [
  ['exact', { en: '12345678', ru: '12345678' }],
  ['decimals', { en: '1234567.8', ru: '1234567,8' }],
  ['decimals', { en: '123456789.5', ru: '123456789,5' }],
  ['exact', { en: '1,234,567', ru: '1 234 567' }],
  ['fits', { en: '-1234567', ru: '-1234567' }],
  ['fits', { en: '12:30:45', ru: '12:30:45' }],
  ['money', { en: '$12,345.67', ru: '12 345,67 ₽' }],
  ['datetime', { en: '3/15/2026 18:30:15', ru: '15.03.2026 18:30:15' }],
  ['formula', { en: '=1/3', ru: '=1/3' }],
  ['formula', { en: '=10^12', ru: '=10^12' }],
  ['formula', { en: '=123456789*1000', ru: '=123456789*1000' }],
  ['pct', { en: '0.5%', ru: '0,5%' }],
]

const ascii = (t: string) => /^[\x20-\x7e]*$/.test(t)

export function buildFitCases(tag: string): FitCase[] {
  const { out, add } = collector()
  const lang = tag === 'ru-RU' ? 'ru' : 'en'
  for (const w of [84, 85, 86, 88, 89, 90, 91, 92, 95, 119, 120, 121, 179, 180, 181]) add('wide', { cells: one(textOf(w)) })
  for (const t of TEXTS) add('text', { cells: one(t) })
  for (const [v, f] of NUMBERS) add('number', { cells: one(v, f) })
  for (const f of ['FALSE', 'SQRT(-1)', 'INDEX(C1:C2,5)', 'C1:C2 E1:E2', '1/0']) add('word', { cells: formula(f) })
  add('rows', { cells: { 1: { value: 'ab' }, 3: { value: 123456789012 }, [HIDDEN]: { value: 46096, format: 'dddd, mmmm d, yyyy' } } })
  add('rows', { cells: { 1: { value: 'a merged text, wider than the rest' }, 2: { value: 'short' }, 4: { value: 'down' } }, merged: [1], mergedDown: [4, 6] })
  add('rows', { cells: { 2: { value: 'x', format: '@' }, 5: { formula: 'REPT("ab",30)' } }, column: { width: 4 } })
  add('rows', { cells: { 1: { format: '#,##0' }, 2: { formula: '""' } }, column: { width: 30 } })
  add('rows', { cells: one(textOf(1600)) })
  for (const via of ['value', 'keys'] as Via[])
    for (const [group, texts] of TYPED) {
      const t = texts[lang]
      if (via === 'keys' && !ascii(t)) continue
      add(`${via}-${group}`, { typed: [[1, t]], via })
    }
  add('keys-fitted', { cells: { 2: { value: 'abc' } }, fitFirst: true, typed: [[1, lang === 'ru' ? '1 234 567' : '1,234,567']], via: 'keys' })
  add('keys-fitted', { cells: { 2: { value: 'a wide text of the column' } }, fitFirst: true, typed: [[1, '123456789012345']], via: 'keys' })
  add('keys-hand', { typed: [[1, '123456789012']], column: { width: 9 }, via: 'keys' })
  return out
}
