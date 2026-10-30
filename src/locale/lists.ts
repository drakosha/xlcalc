/** The custom lists of a workbook's Excel: en-US's short and long day and month names, and its locale's, which fills go on through and sorting by them orders. */
import type { BookInput } from '../contract.js'
import { EN_NAMES } from '../numfmt/style.js'
import { localeOf } from './registry.js'

/** The lists a text goes on through: en-US's short and long day and month names, and the book's locale's. */
export function listsOf(book: Pick<BookInput, 'locale'>): (readonly string[])[] {
  const en = [EN_NAMES.daysShort, EN_NAMES.days, EN_NAMES.monthsShort, EN_NAMES.months]
  const r = localeOf(book.locale).regional
  return r && r.lcid !== '409' ? [...en, r.daysShort, r.days, r.monthsShort, r.months] : en
}
