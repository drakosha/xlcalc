/**
 * The month and day names a format's locale code gives in a book of any locale, loaded or not (`[$-419]mmmm` is
 * `Январь` in en-US too): ru-RU's, which its module takes as its own; en-US's (`409`) are numfmt's. A code without
 * names here shows the book's (`rules.ts`). `[$-FC19]` is ru-RU with the months in the genitive, as Excel for Windows
 * writes `[$-FC19]d mmmm yyyy` (`15 марта 2025`).
 */
import type { Names } from '../numfmt.js'

/** ru-RU's names, as Excel for Windows writes them: the nominative with a capital, three letters short (Excel for Mac: `январь`, `февр`, `сент`). */
export const RU_NAMES: Names = {
  months: ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'],
  monthsShort: ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'],
  days: ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'],
  daysShort: ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'],
}

/** ru-RU's months in the genitive, lowercase: its long date's (`31 января 2025 г.`) and `[$-FC19]`'s. */
export const RU_GENITIVE: readonly string[] = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']

/** `[$-FC19]`'s names: ru-RU's with the months in the genitive. */
export const RU_GENITIVE_NAMES: Names = { ...RU_NAMES, months: RU_GENITIVE }
