/**
 * Logical functions.
 *
 * IF and IFERROR (evaluated lazily: the branch not taken is not computed), TRUE and FALSE are this engine's own.
 */
import type { FnDef } from '../args.js'
import { isErr, toBool, type V } from '../values.js'

export const functions: Record<string, FnDef> = {
  FALSE: { args: '', fn: () => false },
  IF: {
    args: 'a a a?',
    lazy: true,
    fn: (cond: () => V, then: () => V, otherwise?: () => V) => {
      const c = toBool(cond())
      if (isErr(c)) return c
      return c ? then() : otherwise ? otherwise() : false
    },
  },
  IFERROR: {
    args: 'a a',
    lazy: true,
    fn: (value: () => V, alt: () => V) => {
      const v = value()
      return isErr(v) ? alt() : v
    },
  },
  TRUE: { args: '', fn: () => true },
}
