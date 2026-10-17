/**
 * Currency signs in number text in en-US, held out after `euro-ironcalc-held`: which signs VALUE reads besides `$`
 * before a number and `€` before or after it (and `£` it does not), before and after a number, with a space, and
 * currency codes. Answers: `test/excel/cases/currency-held.json`; `PREDICTED` was committed before Excel ran the set.
 */
import type { Val } from '../formulajs.js'
import { heldCases } from './arrays/common.js'
import { e, n } from '../vals.js'

const SIGNS = ['¥', '₽', '₹', '₩', '₪', '₺', '₴', '¢', '฿', '₫', '₦', '₱', '£']

/** Function, formula and the answer predicted before Excel ran. */
const CASES: [string, string, Val][] = [
  ...SIGNS.flatMap((c): [string, string, Val][] => [
    ['VALUE', `VALUE("${c}100")`, e('VALUE')],
    ['VALUE', `VALUE("100${c}")`, e('VALUE')],
  ]),
  ['VALUE', 'VALUE("100 ₽")', e('VALUE')],
  ['VALUE', 'VALUE("Fr100")', e('VALUE')],
  ['VALUE', 'VALUE("100 kr")', e('VALUE')],
  ['VALUE', 'VALUE("USD100")', e('VALUE')],
  ['VALUE', 'VALUE("100 USD")', e('VALUE')],
  ['VALUE', 'VALUE("100 $")', e('VALUE')],
  ['VALUE', 'VALUE("$100")', n(100)],
  ['VALUE', 'VALUE("100 €")', n(100)],
]

/** The set's cases and the answers predicted, by case id. */
export const { buildProbes, PREDICTED } = heldCases(CASES)
