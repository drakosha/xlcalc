/**
 * Logical functions. AND, IFS, NOT, OR and XOR are ported from formula.js 4.6.1 (src/logical.js), Copyright
 * (c) 2014 Sutoiku, Inc., MIT License (THIRD_PARTY.md); IF, IFERROR, IFNA, SWITCH, TRUE and FALSE are this
 * engine's own (lazy: a branch not taken is not computed).
 */
import { eachValue, type FnDef, type Segment } from '../args.js'
import { compareValues, E, isErr, sameType, toBool, type CErr, type V } from '../values.js'

/** The booleans of AND, OR, XOR: numbers and booleans count, text and blanks are skipped (direct text too); none at all is #VALUE!. */
function booleans(segs: Segment[]): boolean[] | CErr {
  const out: boolean[] = []
  const e = eachValue(segs, (v) => {
    if (isErr(v)) return v
    if (typeof v === 'boolean' || typeof v === 'number') out.push(v !== 0 && v !== false)
  })
  if (e) return e
  return out.length ? out : E.VALUE
}

const logical = (f: (bs: boolean[]) => boolean): FnDef => ({
  args: 'V',
  rep: 1,
  fn: (segs: Segment[]) => {
    const bs = booleans(segs)
    return isErr(bs) ? bs : f(bs)
  },
})

export const functions: Record<string, FnDef> = {
  AND: logical((bs) => bs.every(Boolean)),
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
  IFNA: {
    args: 'a a',
    lazy: true,
    fn: (value: () => V, alt: () => V) => {
      const v = value()
      return isErr(v) && v.code === 'NA' ? alt() : v
    },
  },
  IFS: {
    args: 'a a',
    rep: 2,
    fn: (rest: V[]) => {
      for (let i = 0; i < rest.length; i += 2) {
        const c = toBool(rest[i])
        if (isErr(c)) return c
        if (c) return rest[i + 1]
      }
      return E.NA
    },
  },
  NOT: { args: 'b', fn: (b: boolean) => !b },
  OR: logical((bs) => bs.some(Boolean)),
  // values match within a type, text in any case, a blank as 0; a lone last argument is the default
  SWITCH: {
    args: 'a a a',
    rep: 1,
    lazy: true,
    fn: (expr: () => V, ...rest: (() => V)[]) => {
      const x = expr() ?? 0
      if (isErr(x)) return x
      for (let i = 0; i + 1 < rest.length; i += 2) {
        const v = rest[i]() ?? 0
        if (isErr(v)) return v
        if (sameType(x, v) && compareValues(x, v) === 0) return rest[i + 1]()
      }
      return rest.length % 2 ? rest[rest.length - 1]() : E.NA
    },
  },
  TRUE: { args: '', fn: () => true },
  XOR: logical((bs) => bs.filter(Boolean).length % 2 === 1),
}
