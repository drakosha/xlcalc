/**
 * ru-RU's rules of computation: Excel 16.113 for Mac with the Russian interface and locale. Separators, the
 * currency and the format codes are Excel's international settings (`test/excel/locale/ru-RU.json`); the names
 * and the rest were measured (`test/excel/cases/ru-RU/locale.json`), the colours, the date separator of a cell
 * and the long date and time too (`test/excel/cases/ru-RU/text-format-ru-more.json`, `test/excel/display/ru-RU/`).
 */
import { ruRU } from '../ru-RU.js'
import type { RegionalData } from './data.js'

export const RU_RU: RegionalData = {
  locale: ruRU,
  decimal: ',',
  group: ' ',
  currency: { symbol: '₽', before: false, space: ' ', negative: '-' },
  months: ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь'],
  monthsShort: ['янв', 'февр', 'март', 'апр', 'май', 'июнь', 'июль', 'авг', 'сент', 'окт', 'нояб', 'дек'],
  days: ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'],
  daysShort: ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'],
  codes: { day: 'Д', month: 'М', minute: 'м', year: 'Г', hour: 'ч', second: 'с', general: 'Основной', color: 'Цвет' },
  colors: ['Черный', 'Белый', 'Красный', 'Зеленый', 'Синий', 'Желтый', 'Фиолетовый', 'Голубой'],
  dateSeparator: '.',
  lcid: '419',
  longDate: 'dddd, d mmmm yyyy "г."',
  longTime: 'hh:mm:ss',
  longMonths: ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'],
  collation: 'ru-RU',
  // the built-in formats Excel ru-RU shows by codes of its own (the display set `builtin`, `NumberFormatLocal`)
  builtin: builtinRu(),
}

/** ru-RU's codes of built-in formats, as en-US codes: the ruble after the amount, a minus for negatives, two-digit hours. */
function builtinRu(): Record<number, string> {
  const [date, time, rub] = ['dd.mm.yyyy', 'hh:mm:ss', '"₽"']
  const out: Record<number, string> = {
    5: `#,##0 ${rub};-#,##0 ${rub}`,
    6: `#,##0 ${rub};[Red]-#,##0 ${rub}`,
    7: `#,##0.00 ${rub};-#,##0.00 ${rub}`,
    8: `#,##0.00 ${rub};[Red]-#,##0.00 ${rub}`,
    14: date,
    15: 'dd.mmm.yy',
    16: 'dd.mmm',
    17: 'mmm.yy',
    20: 'hh:mm',
    21: time,
    22: `${date} hh:mm`,
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
