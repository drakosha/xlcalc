/**
 * `accountingFormats`: the Accounting Number Format button's left part and menu in en-US and ru-RU, against the menu
 * Excel for Windows showed (`test/excel/win/ribbon/menu/accounting.json`, four runs: the interface language and the
 * regional format en-US or ru-RU each); the codes' category; its input checks.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { EN_TYPED } from '../bench/coverage/ui/cases.js'
import { accountingFormats, InputError, LOCALES, numberFormatCategory } from '../src/index.js'
import { ruRU } from '../src/locales/ru-RU.js'
import './locales.js'

interface Run {
  ui: string
  format: string
  left: { code: string }
  menu: string[]
  items: { name: string; code: string }[]
}
const RUNS = (JSON.parse(readFileSync(new URL('./excel/win/ribbon/menu/accounting.json', import.meta.url), 'utf8')) as { runs: Run[] }).runs
const run = (ui: string, format: string) => RUNS.find((r) => r.ui === ui && r.format === format)!

describe('accountingFormats', () => {
  it("is Excel's menu in a locale whose interface and format are its own: ru-RU's as Excel showed it", () => {
    const got = accountingFormats('ru-RU').items
    expect(got.map(({ name, code }) => ({ name, code }))).toEqual(run('ru-RU', 'ru-RU').items.map(({ name, code }) => ({ name, code })))
    expect(got.map((i) => i.currency)).toEqual(['RUB', 'GBP', 'EUR', 'CNY', 'CHF'])
    expect(accountingFormats(ruRU)).toEqual(accountingFormats('ru-RU'))
  })

  it("en-US's: the machine's menu without the ruble, which its editing language ru-RU added after the dollar", () => {
    // in each run the menu is the format's currency, the interface's and the preferred editing language's (ru-RU) once each, then £ € ¥ CHF
    expect(RUNS.map((r) => r.menu.slice(0, -1).map((n) => n.split(' ')[0]))).toEqual([['$', '₽', '£', '€', '¥', 'CHF'], ['$', '₽', '£', '€', '¥', 'CHF'], ['₽', '$', '£', '€', '¥', 'CHF'], ['₽', '£', '€', '¥', 'CHF']])
    const en = run('en-US', 'en-US').items.filter((i) => !i.name.startsWith('₽'))
    expect(accountingFormats().items.map(({ name, code }) => ({ name, code }))).toEqual(en.map(({ name, code }) => ({ name, code })))
    expect(accountingFormats('en-US').items.map((i) => i.currency)).toEqual(['USD', 'GBP', 'EUR', 'CNY', 'CHF'])
  })

  it("an item's code depends on the regional format, not on the interface's language", () => {
    for (const format of ['en-US', 'ru-RU']) {
      const codes = (ui: string) => Object.fromEntries(run(ui, format).items.map((i) => [i.name.split(' ')[0], i.code]))
      expect(codes('en-US')).toMatchObject(codes('ru-RU'))
    }
  })

  it("the left part applies the built-in 44, which Excel's NumberFormat gives in its typed syntax, the dollar bare", () => {
    for (const tag of LOCALES.keys()) expect(accountingFormats(tag).left).toEqual({ format: 44, code: tag === 'en-US' ? '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)' : '_-* #,##0.00 "₽"_-;-* #,##0.00 "₽"_-;_-* "-"?? "₽"_-;_-@_-' })
    for (const r of RUNS) expect(r.left.code, `${r.ui} ${r.format}`).toBe(EN_TYPED[44])
  })

  it("each item's code is Accounting in the locale, as Excel's Number Format box names it (test/excel/win/ribbon/ribbon-accounting.json)", () => {
    for (const tag of LOCALES.keys()) for (const { code } of [accountingFormats(tag).left, ...accountingFormats(tag).items]) expect(numberFormatCategory(code, tag).category, `${tag} ${code}`).toBe('accounting')
  })

  it('checks its input', () => {
    expect(() => accountingFormats('de-DE')).toThrow(new InputError('locale', `unknown locale "de-DE" (loaded: ${[...LOCALES.keys()].join(', ')}; the package's: en-US, ru-RU)`))
    expect(() => (accountingFormats as (...a: unknown[]) => unknown)('en-US', { locale: 'ru-RU' })).toThrow(InputError)
  })
})
