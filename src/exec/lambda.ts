/**
 * LAMBDA's calls (`Lambda` in `values.ts`): its arguments as its calculation reads them, the depth of calls, and the call
 * of a value as a function (`LAMBDA(x,x*2)(3)`).
 */
import type { Arg } from '../args.js'
import { E, isErr, Lambda } from '../values.js'
import { ArraySizeError, spend } from './budget.js'

/** The value of a parameter left out of the call (`f(1,)`): blank where it is read, and ISOMITTED tells it from a blank given. */
export const OMITTED = (): Arg => null

/** The values of a call's arguments for a LAMBDA of `count` parameters, as its calculation reads them. */
export const parameters = (args: readonly (Arg | undefined)[], count: number): (() => Arg)[] =>
  Array.from({ length: count }, (_, i) => {
    const a = args[i]
    return a === undefined ? OMITTED : () => a
  })

/**
 * How deep LAMBDAs call one another: past it a call is #NUM!, as Excel answers too many recursive calls, before the
 * JavaScript stack would run out (a stack that does run out is #NUM! too).
 */
export const MAX_CALL_DEPTH = 1024

let depth = 0

/**
 * A LAMBDA of `count` parameters whose calculation `run` computes: a call gives each of them, one left out between commas
 * (`f(1,)`) too, as Excel takes no fewer; each call counts as a unit of the computation's work.
 */
export function lambdaOf(count: number, run: (args: readonly (Arg | undefined)[]) => Arg): Lambda {
  return new Lambda(count, count, (args) => {
    if (depth >= MAX_CALL_DEPTH) return E.NUM
    spend(1)
    depth++
    try {
      return run(args as readonly (Arg | undefined)[])
    } catch (e) {
      if (e instanceof RangeError && !(e instanceof ArraySizeError)) return E.NUM
      throw e
    } finally {
      depth--
    }
  })
}

/** A value called with `args`: a function within its counts, #VALUE! for wrong counts or a value that is no function, an error as it is. */
export function callValue(f: Arg, args: readonly (Arg | undefined)[]): Arg {
  if (f instanceof Lambda) return args.length < f.min || args.length > f.max ? E.VALUE : (f.run(args) as Arg)
  return isErr(f) ? f : E.VALUE
}
