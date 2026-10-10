/**
 * The month and day names a format's locale code gives in a book of any locale, loaded or not (`[$-419]mmmm` is
 * `Январь` in en-US too): ru-RU's, which its module takes as its own; en-US's (`409`) are numfmt's. A code without
 * names here shows the book's (`rules.ts`).
 */
import type { Names } from '../numfmt.js'

/** ru-RU's names, as Excel for Windows writes them: the nominative with a capital, three letters short (Excel for Mac: `январь`, `февр`, `сент`). */
export const RU_NAMES: Names = {
  months: ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'],
  monthsShort: ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'],
  days: ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'],
  daysShort: ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'],
}
