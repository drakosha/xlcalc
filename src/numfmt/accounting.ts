/**
 * The Accounting Number Format button of Excel's Home tab (`$`) in a locale (`accountingFormats`): its left part applies
 * the built-in format 44, each item of its menu a currency's accounting format. Measured in Excel 16.0 for Windows
 * (`test/excel/win/ribbon/menu/accounting.json`): the menu lists the currency of the regional format first, then £, €,
 * ¥ and CHF, each in its own form (€ in the regional format's); an interface or editing language of another currency
 * adds that currency after the first, which the menu of a locale here, its language and format one, does not have.
 */
import type { AccountingFormats } from '../contract.js'
import { optionalLocale } from '../locale/load.js'
import type { AccountingCurrency } from '../locale/rules/data.js'
import type { Locale } from '../locale/types.js'
import { builtinIn } from './builtin.js'

/** The forms of the menu's codes by the symbol `s`: en-US's in parentheses, en-GB's, zh-CN's, and the symbol after the digits. */
const FORMS: Record<AccountingCurrency['form'], (s: string) => string> = {
  paren: (s) => `_(${s}* #,##0.00_);_(${s}* (#,##0.00);_(${s}* "-"??_);_(@_)`,
  gb: (s) => `_-${s}* #,##0.00_-;-${s}* #,##0.00_-;_-${s}* "-"??_-;_-@_-`,
  cn: (s) => `_ ${s}* #,##0.00_ ;_ ${s}* -#,##0.00_ ;_ ${s}* "-"??_ ;_ @_ `,
  after: (s) => `_-* #,##0.00 ${s}_-;-* #,##0.00 ${s}_-;_-* "-"?? ${s}_-;_-@_-`,
}

/** en-US's menu. */
const EN_ITEMS: readonly AccountingCurrency[] = [
  { currency: 'USD', symbol: '[$$-409]', form: 'paren', name: '$ English (United States)' },
  { currency: 'GBP', symbol: '[$£-809]', form: 'gb', name: '£ English (United Kingdom)' },
  { currency: 'EUR', symbol: '[$€-2] ', form: 'paren', name: '€ Euro (€ 123)' },
  { currency: 'CNY', symbol: '[$¥-804]', form: 'cn', name: '¥ Chinese (Simplified, Mainland China)' },
  { currency: 'CHF', symbol: '[$CHF-100C]', form: 'after', name: 'CHF French (Switzerland)' },
]

/**
 * The Accounting Number Format button in `locale` (a loaded locale's tag or a locale object; en-US without one): its
 * left part's format, the built-in 44 with its code in the locale, and its menu's currency items in order, each with
 * its name in the locale's language and the code it applies.
 */
export function accountingFormats(locale?: string | Locale): AccountingFormats {
  const loc = optionalLocale(locale, 'locale')
  const items = loc?.regional?.formats?.accounting ?? EN_ITEMS
  return { left: { format: 44, code: builtinIn(loc)(44) }, items: items.map(({ currency, symbol, form, name }) => ({ currency, name, code: FORMS[form](symbol) })) }
}
