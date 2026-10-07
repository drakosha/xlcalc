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
}
