/**
 * Date and time functions: this engine's own, in Excel's 1900 date system (`src/dates.ts`).
 */
import type { Cx, FnDef } from '../args.js'
import { addMonths, daysInMonth, MAX_SERIAL, serialOf, ymd } from '../dates.js'
import { E } from '../values.js'

export const functions: Record<string, FnDef> = {
  // years 0–1899 are added to 1900; month and day roll over through Excel's calendar, so DATE(1900,2,29) is 60
  DATE: {
    args: 'i i i',
    fn: (y: number, m: number, d: number) => {
      const year = y >= 0 && y < 1900 ? y + 1900 : y
      if (year < 1900 || year > 9999) return E.NUM
      const first = addMonths(year, m, 0)
      if (first.y < 1899 || first.y > 9999) return E.NUM
      const serial = serialOf(first.y, first.m, 1) + d - 1
      return serial < 0 || serial > MAX_SERIAL ? E.NUM : serial
    },
  },
  // the start month in Excel's calendar (serial 0 is January 1900, 60 is February); the last day by the real one
  EOMONTH: {
    args: 'd i',
    fn: (start: number, months: number) => {
      const s = ymd(start)
      const t = addMonths(s.y, s.m, months)
      if (t.y < 1900 || t.y > 9999) return E.NUM
      return serialOf(t.y, t.m, daysInMonth(t.y, t.m))
    },
  },
  NOW: { args: '', cx: true, volatile: true, fn: (cx: Cx) => cx.now },
  TODAY: { args: '', cx: true, volatile: true, fn: (cx: Cx) => Math.floor(cx.now) },
}
