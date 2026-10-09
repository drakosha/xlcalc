/**
 * ru-RU's rules of computation: Excel 16.113 for Mac with the Russian interface and locale. Separators, the
 * currency and the format codes are Excel's international settings (`test/excel/locale/ru-RU.json`); the names
 * and the rest were measured (`test/excel/cases/ru-RU/locale.json`), the colours, the date separator of a cell
 * and the long date too (`test/excel/cases/ru-RU/text-format-ru-more.json`, `test/excel/display/ru-RU/`). Where the two
 * Excels' regional data part, the rules are Excel for Windows's (`test/excel/win/`): the order of text, month names,
 * the short and long time.
 */
import { ruRU } from '../ru-RU.js'
import type { RegionalData } from './data.js'

export const RU_RU: RegionalData = {
  locale: ruRU,
  decimal: ',',
  group: ' ',
  currency: { symbol: '₽', before: false, space: ' ', negative: '-' },
  // Windows's names: the nominative with a capital, three letters short (Excel for Mac: `январь`, `февр`, `сент`)
  months: ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'],
  monthsShort: ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'],
  days: ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'],
  daysShort: ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'],
  codes: { day: 'Д', month: 'М', minute: 'м', year: 'Г', hour: 'ч', second: 'с', general: 'Основной', color: 'Цвет' },
  colors: ['Черный', 'Белый', 'Красный', 'Зеленый', 'Синий', 'Желтый', 'Фиолетовый', 'Голубой'],
  dateSeparator: '.',
  lcid: '419',
  longDate: 'dddd, d mmmm yyyy "г."',
  // the hour without a leading zero, as Windows writes the time (Excel for Mac: `hh:mm:ss`)
  longTime: 'h:mm:ss',
  longMonths: ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'],
  // Latin before Cyrillic, as Windows sorts in every locale (Excel for Mac ru-RU: Cyrillic first)
  collation: 'en-US',
  // the built-in formats Excel ru-RU shows by codes of its own (the display set `builtin`, `NumberFormatLocal`)
  builtin: builtinRu(),
}

/** ru-RU's codes of built-in formats, as en-US codes: the ruble after the amount, a minus for negatives, hours without a leading zero (Windows's `ч:мм`). */
function builtinRu(): Record<number, string> {
  const [date, time, rub] = ['dd.mm.yyyy', 'h:mm:ss', '"₽"']
  const out: Record<number, string> = {
    5: `#,##0 ${rub};-#,##0 ${rub}`,
    6: `#,##0 ${rub};[Red]-#,##0 ${rub}`,
    7: `#,##0.00 ${rub};-#,##0.00 ${rub}`,
    8: `#,##0.00 ${rub};[Red]-#,##0.00 ${rub}`,
    14: date,
    15: 'dd.mmm.yy',
    16: 'dd.mmm',
    17: 'mmm.yy',
    20: 'h:mm',
    21: time,
    22: `${date} h:mm`,
    36: date,
    37: '#,##0;-#,##0',
    38: '#,##0;[Red]-#,##0',
    39: '#,##0.00;-#,##0.00',
    40: '#,##0.00;[Red]-#,##0.00',
    41: '_-* #,##0_-;-* #,##0_-;_-* "-"_-;_-@_-',
    42: `_-* #,##0 ${rub}_-;-* #,##0 ${rub}_-;_-* "-" ${rub}_-;_-@_-`,
    43: '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-',
    44: `_-* #,##0.00 ${rub}_-;-* #,##0.00 ${rub}_-;_-* "-"?? ${rub}_-;_-@_-`,
  }
  for (let id = 27; id <= 31; id++) out[id] = date
  for (let id = 32; id <= 35; id++) out[id] = time
  return out
}
