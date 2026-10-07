/** Time of day and the clock functions. */
import type { Cx, FnDef } from '../../args.js'
import { E, isErr, type V } from '../../values.js'
import { textDate } from './calendar.js'

/** The day fraction to the nearest second, as Excel's HOUR, MINUTE and SECOND read it. */
function clock(t: number): { h: number; m: number; s: number } {
  const secs = Math.round((t - Math.floor(t)) * 86_400) % 86_400
  return { h: Math.floor(secs / 3600), m: Math.floor(secs / 60) % 60, s: secs % 60 }
}

export const functions: Record<string, FnDef> = {
  HOUR: { args: 'h', fn: (t: number) => clock(t).h },
  MINUTE: { args: 'h', fn: (t: number) => clock(t).m },
  NOW: { args: '', cx: true, volatile: true, fn: (cx: Cx) => cx.now },
  SECOND: { args: 'h', fn: (t: number) => clock(t).s },
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
