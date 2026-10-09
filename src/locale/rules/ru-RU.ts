/**
 * ru-RU's rules of computation: Excel 16.113 for Mac with the Russian interface and locale. Separators, the
 * currency and the format codes are Excel's international settings (`test/excel/locale/ru-RU.json`); the names
 * and the rest were measured (`test/excel/cases/ru-RU/locale.json`), the colours, the date separator of a cell
 * and the long date too (`test/excel/cases/ru-RU/text-format-ru-more.json`, `test/excel/display/ru-RU/`), the name of a
 * new table column (`test/excel/tables/ru-RU/locale.json`). Where the two Excels' regional data part, the rules are
 * Excel for Windows's (`test/excel/win/`): the order of text, month names, the short and long time, CHAR and CODE.
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
  // Windows-1251, as Excel for Windows computes CHAR and CODE (Excel for Mac: Mac Cyrillic); `test/codepage.test.ts` checks it
  codePage: String.fromCharCode(
    0x0402, 0x0403, 0x201a, 0x0453, 0x201e, 0x2026, 0x2020, 0x2021, 0x20ac, 0x2030, 0x0409, 0x2039, 0x040a, 0x040c, 0x040b, 0x040f,
    0x0452, 0x2018, 0x2019, 0x201c, 0x201d, 0x2022, 0x2013, 0x2014, 0x0098, 0x2122, 0x0459, 0x203a, 0x045a, 0x045c, 0x045b, 0x045f,
    0x00a0, 0x040e, 0x045e, 0x0408, 0x00a4, 0x0490, 0x00a6, 0x00a7, 0x0401, 0x00a9, 0x0404, 0x00ab, 0x00ac, 0x00ad, 0x00ae, 0x0407,
    0x00b0, 0x00b1, 0x0406, 0x0456, 0x0491, 0x00b5, 0x00b6, 0x00b7, 0x0451, 0x2116, 0x0454, 0x00bb, 0x0458, 0x0405, 0x0455, 0x0457,
    ...Array.from({ length: 64 }, (_, i) => 0x0410 + i),
  ),
  columnWord: 'Столбец',
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
