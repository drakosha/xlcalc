/**
 * Function calls over arrays (`docs/arrays-incell.md`): a one-value argument given an array or a range of
 * several cells lifts the call element by element; the lazy functions (IF, IFERROR, IFNA, SWITCH) stay
 * lazy unless an argument that lifts them is an array, and then every argument is computed.
 */
import type { Arg, Cx, FunctionSpec } from '../args.js'
import { Areas, Rng, type V } from '../values.js'
import { first, liftN, lifts } from './lift.js'
import { holdsReferences, keepReferences, placesOf, referenceAt } from './references.js'
import { arr, Arr } from './value.js'

type Call = (...a: any[]) => Arg

/** A one-cell reference is its value (an array, a 1 × 1 one too, lifts the call). */
const single = (x: Arg, cx: Cx): Arg => (x instanceof Rng && !lifts(x) ? cx.scalar(x) : x)

/** The top-left cell of a reference (N and T take it from a range of several cells). */
const topLeft = (g: Rng, cx: Cx): V => cx.scalar(new Rng(g.s, g.r1, g.c1, g.r1, g.c1))

/** The caller of a function whose one-value arguments may get arrays (the generated code calls it only then). */
export function liftCaller(spec: FunctionSpec, cell: boolean, call: Call, cx: Cx): Call {
  const off = cell ? 3 : 0
  return (...a: Arg[]) => {
    let many = false
    for (let i = off; i < a.length; i++) {
      const p = i - off
      const mode = spec.mode(p)
      const x = a[i]
      // an array of references where the function takes a reference lifts the call together with the arrays of values
      if (spec.referencesAt(p) && holdsReferences(x)) many = true
      if (mode === 'whole' || mode === 'pass' || !(x instanceof Arr || x instanceof Rng)) continue
      if (x instanceof Rng && mode === 'array') continue
      if (!lifts(x)) a[i] = single(x, cx)
      else if (x instanceof Arr) many = true
      else {
        const how = spec.rangeAt(p)
        if (how === 'first') a[i] = topLeft(x, cx)
        else if (how === 'lift') many = true
      }
    }
    if (!many) return call(...a)
    // lifted: the arrays, the liftable ranges and the passed-through arguments take part
    const at: number[] = []
    const xs: (V | Arr)[] = []
    // the arrays of references among them, by their place in `at`
    const byRef: (Arr | undefined)[] = []
    for (let i = off; i < a.length; i++) {
      const p = i - off
      const mode = spec.mode(p)
      const x = a[i]
      if (spec.referencesAt(p) && holdsReferences(x)) {
        byRef[at.length] = x
        at.push(i)
        xs.push(placesOf(x))
        continue
      }
      if (!(x instanceof Arr || x instanceof Rng) || mode === 'whole') continue
      if (x instanceof Rng && (mode === 'array' || (mode === 'lift' && spec.rangeAt(p) !== 'lift'))) continue
      at.push(i)
      xs.push(cx.view(x))
    }
    const args = a.slice()
    // VLOOKUP's index: each lifted call reads the index's top-left number (an invalid one elsewhere does not count)
    for (const p of spec.liftFirst) {
      const x = args[off + p]
      if (x instanceof Arr || x instanceof Rng) args[off + p] = arr(1, 1, [x instanceof Arr ? x.at(0, 0) : topLeft(x, cx)])
    }
    // OFFSET and INDIRECT keep the references they find for the functions that take one (`references.ts`); CELL and
    // GETPIVOTDATA, dynamic too, answer values, which are no references
    const found: Arg[] | undefined = spec.dynamic ? [] : undefined
    let refs = false
    const out = liftN(xs, (els) => {
      for (let k = 0; k < at.length; k++) args[at[k]] = byRef[k] ? (referenceAt(byRef[k]!, els[k]) as Arg) : els[k]
      const r = call(...args)
      if (found) {
        found.push(r)
        if (r instanceof Rng || r instanceof Areas) refs = true
      }
      return first(r, (g) => topLeft(g, cx))
    }, true)
    return refs ? keepReferences(out, found!) : out
  }
}

/**
 * A lazy function with arguments that may be arrays: the ones that lift it (`spec.liftAt`, among `check`,
 * the positions the code generator could not prove scalar) are computed first; an array among them computes
 * every argument and lifts the call over all of them, else the call stays lazy.
 */
export function liftLazy(fn: Call, spec: FunctionSpec, check: readonly number[], thunks: (() => Arg)[], cx: Cx): Arg {
  const n = thunks.length
  const done: (Arg | undefined)[] = new Array(n)
  let many = false
  for (const i of check) {
    if (i >= n || !spec.liftAt(i, n)) continue
    const v = thunks[i]()
    if (lifts(v)) many = true
    done[i] = v
  }
  if (!many) return fn(...thunks.map((t, i) => (done[i] === undefined ? t : () => single(done[i]!, cx))))
  const all = thunks.map((t, i) => (done[i] === undefined ? t() : done[i]!))
  const xs = all.map((x) => cx.view(x))
  return liftN(xs, (els) => first(fn(...els.map((e) => () => e)), (g) => topLeft(g, cx)), true)
}
