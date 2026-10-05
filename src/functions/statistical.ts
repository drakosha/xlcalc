/**
 * Statistical functions.
 *
 * COUNTIFS is this engine's own.
 */
import type { Cx, FnDef } from '../args.js'
import { isErr, type V } from '../values.js'

export const functions: Record<string, FnDef> = {
  COUNTIFS: {
    args: 'R a',
    rep: 2,
    cx: true,
    fn: (cx: Cx, rest: V[]) => {
      const pairs: [V, V][] = []
      for (let i = 0; i < rest.length; i += 2) pairs.push([rest[i], rest[i + 1]])
      const r = cx.select(pairs)
      return isErr(r) ? r : r.sel.length
    },
  },
}
