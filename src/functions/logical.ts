/**
 * Logical functions. AND, IFS, NOT, OR and XOR are ported from formula.js 4.6.1 (src/logical.js), Copyright
 * (c) 2014 Sutoiku, Inc., MIT License (THIRD_PARTY.md); IF, IFERROR, IFNA, SWITCH, TRUE and FALSE are this
 * engine's own (lazy: a branch not taken is not computed).
 */
import { directError, eachValue, type Arg, type Cx, type FnDef, type Segment } from '../args.js'
import { arr } from '../arrays/value.js'
import type { Rules } from '../locale/rules.js'
import { Areas, E, isErr, sameType, type CErr, type V } from '../values.js'
import { NEWER_MAX_ARGS } from './common.js'

/**
 * The booleans of AND, OR, XOR: numbers and booleans count, a direct text argument that is the locale's word for
 * TRUE or FALSE counts as it (the logical-text set), a direct blank as FALSE (`AND(TRUE,)`), other text and
 * blanks in ranges are skipped; none at all is #VALUE!.
 */
function booleans(segs: Segment[], rules: Rules): boolean[] | CErr {
  const out: boolean[] = []
  const e = directError(segs) ?? eachValue(segs, (v, inRange) => {
    if (isErr(v)) return v
    if (typeof v === 'boolean' || typeof v === 'number' || (v === null && !inRange)) out.push(v !== 0 && v !== false && v !== null)
    else if (typeof v === 'string' && !inRange) {
      const b = rules.toBool(v)
      if (typeof b === 'boolean') out.push(b)
    }
  })
  if (e) return e
  return out.length ? out : E.VALUE
}

const logical = (f: (bs: boolean[]) => boolean): FnDef => ({
  args: 'V',
  rep: 1,
  cx: true,
  fn: (cx: Cx, segs: Segment[]) => {
    const bs = booleans(segs, cx.rules)
    return isErr(bs) ? bs : f(bs)
  },
})

export const functions: Record<string, FnDef> = {
  AND: logical((bs) => bs.every(Boolean)),
  FALSE: { args: '', fn: () => false },
  // the branch taken as it is, a reference too (SUM(IF(TRUE,A2)) skips text in A2)
  IF: {
    args: 'a P P?',
    lazy: true,
    cx: true,
    fn: (cx: Cx, cond: () => V, then: () => V, otherwise?: () => V) => {
      const c = cx.rules.toBool(cond())
      if (isErr(c)) return c
      return c ? then() : otherwise ? otherwise() : false
    },
  },
  IFERROR: {
    args: 'a p',
    lazy: true,
    fn: (value: () => V, alt: () => V) => {
      const v = value()
      return isErr(v) ? alt() : v
    },
  },
  IFNA: {
    args: 'a p',
    lazy: true,
    fn: (value: () => V, alt: () => V) => {
      const v = value()
      return isErr(v) && v.code === 'NA' ? alt() : v
    },
  },
  IFS: {
    args: 'a P',
    rep: 2,
    cx: true,
    array: true,
    fn: (cx: Cx, rest: V[]) => {
      for (let i = 0; i < rest.length; i += 2) {
        const c = cx.rules.toBool(rest[i])
        // a condition of text that is no boolean: Excel's #VALUE! is an array of one there (ROWS counts 1), an error given as one is the answer itself
        if (isErr(c)) return isErr(rest[i]) ? c : arr(1, 1, [c])
        if (c) return rest[i + 1]
      }
      return E.NA
    },
  },
  NOT: { args: 'b', fn: (b: boolean) => !b },
  OR: logical((bs) => bs.some(Boolean)),
  // values match within a type, text in any case, a blank as 0; a lone last argument is the default;
  // the expression and the values to match lift, the results do not; a result is answered as it is, a
  // reference too (the values to match come the same way and are read as values); 126 pairs at most
  SWITCH: {
    args: 'a P P',
    rep: 1,
    max: 254,
    lazy: true,
    cx: true,
    liftAt: (i, n) => i === 0 || (i % 2 === 1 && i < n - 1),
    fn: (cx: Cx, expr: () => V, ...rest: (() => Arg)[]) => {
      const x = expr() ?? 0
      if (isErr(x)) return x
      for (let i = 0; i + 1 < rest.length; i += 2) {
        // a reference across sheets compared is #REF!, as in Excel
        const c = rest[i]()
        const v = c instanceof Areas && c.cube ? E.REF : (cx.scalar(c) ?? 0)
        if (isErr(v)) return v
        if (sameType(x, v) && cx.rules.compare(x, v) === 0) return rest[i + 1]()
      }
      return rest.length % 2 ? rest[rest.length - 1]() : E.NA
    },
  },
  TRUE: { args: '', fn: () => true },
  XOR: { ...logical((bs) => bs.filter(Boolean).length % 2 === 1), max: NEWER_MAX_ARGS },
}
