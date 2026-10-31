/** The custom lists of a workbook's Excel: en-US's short and long day and month names, and its locale's, which fills go on through and sorting by them orders. */
import type { BookInput } from '../contract.js'
import { EN_NAMES } from '../numfmt/style.js'
import { addedLocales, localeOf } from './registry.js'

/** The lists a text goes on through: en-US's short and long day and month names, and the book's locale's. */
export function listsOf(book: Pick<BookInput, 'locale'>): (readonly string[])[] {
  const en = [EN_NAMES.daysShort, EN_NAMES.days, EN_NAMES.monthsShort, EN_NAMES.months]
  const r = localeOf(book.locale).regional
  return r && r.lcid !== '409' ? [...en, r.daysShort, r.days, r.monthsShort, r.months] : en
}

/**
 * The lists sorting a PivotTable by its labels takes: the workbook's (`listsOf`), and the day and month names of every
 * other locale the host loaded, as the Excel of the workbook's author knows its own language's lists whatever the
 * workbook's locale (a corpus of workbooks of Excel ru-RU: Russian month names sorted январь to декабрь).
 */
export function sortListsOf(book: Pick<BookInput, 'locale'>): (readonly string[])[] {
  // a locale's lists come after en-US's four
  const others = addedLocales().filter((l) => l.tag !== (book.locale ?? 'en-US')).flatMap((l) => listsOf({ locale: l.tag }).slice(4))
  return [...listsOf(book), ...others]
}
