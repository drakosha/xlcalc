/**
 * The items of the Number Format list on Excel's Home tab in a locale (`numberFormatChoices`), with the names of the
 * box's categories and what the list writes under General: the data `category.ts` names categories by and
 * `samples.ts` writes samples by, apart from the rules of categories, which a host showing the list needs not.
 */
import type { FormatChoice, FormatChoiceKey, NumberFormat } from '../contract.js'
import { optionalLocale } from '../locale/load.js'
import type { FormatNames } from '../locale/rules/data.js'
import type { Locale } from '../locale/types.js'
import { builtinIn } from './builtin.js'

/** en-US's names of the box's categories and the list's items, and the code of its Currency item. */
const EN_NAMES = {
  categories: { general: 'General', number: 'Number', currency: 'Currency', accounting: 'Accounting', date: 'Date', time: 'Time', percentage: 'Percentage', fraction: 'Fraction', scientific: 'Scientific', text: 'Text', special: 'Special', custom: 'Custom' },
  choices: { general: 'General', number: 'Number', currency: 'Currency', accounting: 'Accounting', shortDate: 'Short Date', longDate: 'Long Date', time: 'Time', percentage: 'Percentage', fraction: 'Fraction', scientific: 'Scientific', text: 'Text' },
  currency: '"$"#,##0.00',
  noFormat: 'No specific format',
} satisfies FormatNames

/** The system's long date and time, as Excel keeps them (`[$-x-sysdate]` reads as F800). */
export const LONG_DATE = '[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy'
export const LONG_TIME = '[$-F400]h:mm:ss\\ AM/PM'

/** The list's items in its order, each with what it applies: a built-in format by number, or a code (the Currency item's is the locale's). */
const CHOICES: readonly [FormatChoiceKey, NumberFormat | undefined][] = [
  ['general', 0],
  ['number', 2],
  ['currency', undefined],
  ['accounting', 44],
  ['shortDate', 14],
  ['longDate', LONG_DATE],
  ['time', LONG_TIME],
  ['percentage', 10],
  ['fraction', 12],
  ['scientific', 11],
  ['text', 49],
]

export const namesOf = (loc: Locale | undefined): FormatNames => loc?.regional?.formats ?? EN_NAMES

/** What the list writes under General in a locale (`numberFormatSamples`). */
export const noFormatText = (loc: Locale | undefined): string => namesOf(loc).noFormat ?? EN_NAMES.noFormat

/** The list's items in a locale, in list order, each with the format it applies. */
export const choiceFormats = (loc: Locale | undefined): [FormatChoiceKey, NumberFormat][] => CHOICES.map(([key, format = namesOf(loc).currency]) => [key, format])

/** The items of the Number Format list in `locale`, in list order: each item's key, name, the format it applies and that format's code in the locale. */
export function numberFormatChoices(locale?: string | Locale): FormatChoice[] {
  const loc = optionalLocale(locale, 'locale')
  const names = namesOf(loc)
  const codes = builtinIn(loc)
  return choiceFormats(loc).map(([key, format]) => ({ key, name: names.choices[key], format, code: typeof format === 'number' ? codes(format) : format }))
}
