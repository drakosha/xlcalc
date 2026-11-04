/**
 * Which number format codes Excel opens a file with, and what it shows of them: literals without quotes (Cyrillic, `₽`, every
 * Latin letter, words) beside digit placeholders, in dates and with `@`, and codes broken otherwise. A refused code costs
 * the whole file, as `numFmts` is read on opening, so `writeXlsx` must not write one. Excel for Windows alone: its step halves
 * a book Excel refuses down to the codes (the Mac's stops at the repair window).
 *
 * `REFUSED_PREDICTED` was committed before Excel ran the set. The guess: a letter Excel reads as a date's or an era's code
 * (`a b d e g h m n s y`, either case) beside a digit placeholder refuses the file, as in `digits-date-letters`; any other
 * letter, Cyrillic and `₽` are literals and show as written; so `0 шт`, `#,##0 руб` and `#,##0 ₽` open.
 */
import type { DisplayCase } from './cases.js'

const RUSSIAN = ['#,##0 ₽', '#,##0.00 ₽', '0 шт', '0 шт.', '#,##0 руб', '#,##0 руб.', '#,##0,0 "млн"', '#,##0,0 млн', '0 м', '0 кг', '0 г', '0 д', '0 ч', 'шт 0', '₽ 0', '0%шт', '0 шт;-0 шт', '[Red]0 шт', '[Красный]0']
const SYMBOLS = ['0 €', '0 £', '0 ¥', '$0', '0 $', '[$₽-419]0', '[$€-407]0', '0 №', '0 °', '0 °C', '0 ×', '0 µ']
const LETTERS = [...'abcdefghijklmnopqrstuvwxyz'].flatMap((c) => [`0 ${c}`, `0 ${c.toUpperCase()}`])
const WORDS = ['0 samples', '0 pcs', '0 kg', '0 USD', '0 usd', '0 km', '0 tn', '#,##0 x', '0 "pcs"', '0 \\p\\c\\s', '0 \\ш\\т']
const DATES = ['dd.mm.yyyy г.', 'yyyy "г."', 'd mmm yyyy г', 'dd шт', 'h:mm ч', 'ДД.ММ.ГГГГ', 'дд.мм.гггг', 'dd x', 'dd p', 'dd 0']
const TEXTS = ['шт', 'abc', 'xyz', '@ шт', 'шт @', '@ x', 'General шт', 'General "шт"', 'General x', '0 @', '@ 0', '0;@;0', '0;-0;0;@', '"шт"']
const BROKEN = ['0 "шт', '[Red 0', '0]', '[Foo]0', '0;0;0;0;0', '0\\', '0_', '0*', 'E+0', '0E0', '0e+0', '0E+0', '0.0.0', '0/', '0/0/0', `"${'a'.repeat(300)}"0`]

export const CODES = [...RUSSIAN, ...SYMBOLS, ...LETTERS, ...WORDS, ...DATES, ...TEXTS, ...BROKEN]

export const REFUSED_PREDICTED = [
  '0 м'.replace('м', 'm'),
  ...LETTERS.filter((c) => /^0 [abdeghmnsy]$/i.test(c) && c !== '0 m'),
  '0 samples',
  '0 kg',
  '0 USD',
  '0 usd',
  '0 km',
  '0 "шт',
  '[Red 0',
  '[Foo]0',
  '[Красный]0',
  '0;0;0;0;0',
  '0e+0',
  '0E0',
  '0\\',
  '0_',
  '0*',
  '@ 0',
  '0 @',
  'dd x',
  'dd p',
  'dd 0',
]

export function buildDisplayCases(): DisplayCase[] {
  return CODES.map((format, i) => ({ id: `c${i + 1}`, cell: { value: 1234.5, format }, width: 40 }))
}
