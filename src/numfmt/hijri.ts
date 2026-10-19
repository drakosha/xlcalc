/**
 * The Hijri calendar of `B2` (ECMA-376 numFmts), as Excel for Mac shows it in every locale: the tabular calendar of
 * 30-year cycles from 1 Muharram 1 AH, 622-07-15 (Hijri 9/25/1420 is 2000-01-01, Excel's help), and its names in
 * Arabic as Excel writes them, a no-break space between words; `mmm` is the whole name and `ddd` the whole day.
 */
import type { Names } from './style.js'

const MONTHS = ['محرم', 'صفر', 'ربيع الأول', 'ربيع الثاني', 'جمادى الأولى', 'جمادى الثانية', 'رجب', 'شعبان', 'رمضان', 'شوال', 'ذو القعدة', 'ذو الحجة']
const DAYS = ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت']

export const HIJRI_NAMES: Names = { months: MONTHS, monthsShort: MONTHS, days: DAYS, daysShort: DAYS }

/** The day before 1 Muharram 1 AH, counted from 1899-12-30 (the day before serial 1 of the 1900 system). */
const EPOCH = -466_581

/** Days from the epoch to the first of a Hijri month. */
const start = (y: number, m: number) => 29 * (m - 1) + Math.floor((6 * m - 1) / 11) + (y - 1) * 354 + Math.floor((3 + 11 * y) / 30)

/** The Hijri date of a serial of the 1900 system: Excel's serials before 61 count the phantom 1900-02-29, so they are a day short. */
export function hijri(serial: number): { y: number; m: number; d: number } {
  const days = (serial < 61 ? serial + 1 : serial) - EPOCH
  const y = Math.floor((30 * (days - 1) + 10646) / 10631)
  const m = Math.min(12, Math.floor((11 * (days - start(y, 1) - 1) + 330) / 325))
  return { y, m, d: days - start(y, m) }
}
