/** Time of day and the clock functions. */
import type { Cx, FnDef } from '../../args.js'
import { MAX_SERIAL, momentOf } from '../../dates.js'
import { E, isErr, type CErr, type V } from '../../values.js'
import { textDate } from './calendar.js'

/**
 * `f` of the second of the day of a moment, as HOUR, MINUTE and SECOND read it (`momentOf`), in the book's date system
 * (`epoch`); a moment past 9999-12-31 once rounded is #NUM! (the set `seconds-held`).
 */
function clock(t: number, epoch: number, f: (second: number) => number): number | CErr {
  const { day, second } = momentOf(t)
  return day + epoch > MAX_SERIAL ? E.NUM : f(second)
}

export const functions: Record<string, FnDef> = {
  HOUR: { args: 'h', cx: true, fn: (cx: Cx, t: number) => clock(t, cx.rules.epoch, (s) => Math.floor(s / 3600)) },
  MINUTE: { args: 'h', cx: true, fn: (cx: Cx, t: number) => clock(t, cx.rules.epoch, (s) => Math.floor(s / 60) % 60) },
  NOW: { args: '', cx: true, volatile: true, fn: (cx: Cx) => cx.now },
  SECOND: { args: 'h', cx: true, fn: (cx: Cx, t: number) => clock(t, cx.rules.epoch, (s) => s % 60) },
  // INT of each part, each at most 32767, wrapped to a day; a negative total is #NUM!
  TIME: {
    args: 'n n n',
    fn: (h: number, m: number, s: number) => {
      const [hh, mm, ss] = [Math.floor(h), Math.floor(m), Math.floor(s)]
      if (hh > 32767 || mm > 32767 || ss > 32767) return E.NUM
      const total = hh * 3600 + mm * 60 + ss
      return total < 0 ? E.NUM : (total % 86_400) / 86_400
    },
  },
  TIMEVALUE: {
    args: 'v',
    cx: true,
    fn: (cx: Cx, v: V) => {
      const t = textDate(v, cx.rules)
      return isErr(t) ? t : t - Math.floor(t)
    },
  },
  TODAY: { args: '', cx: true, volatile: true, fn: (cx: Cx) => Math.floor(cx.now) },
}
