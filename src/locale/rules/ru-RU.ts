/**
 * ru-RU's rules of computation: Excel 16.113 for Mac with the Russian interface and locale. Separators, the
 * currency and the format codes are Excel's international settings (`docs/locale/excel-ru-RU.json`); the names
 * and the rest were measured (`docs/excel-cases/ru-RU/locale.json`).
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
  codes: { day: 'Д', month: 'М', minute: 'м', year: 'Г', hour: 'ч', second: 'с', general: 'Основной' },
  collation: 'ru-RU',
}
