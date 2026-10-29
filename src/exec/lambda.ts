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
 * How deep LAMBDAs may call one another, by the heights of the calculations of the calls under way (how deep each one's
 * code nests, 5 for `IF(n<=0,0,1+g(g,n-1))`): past it a call is #NUM!, as Excel answers too many recursive calls. Excel's
 * own limit is its stack's: 4,000 calls of that calculation, 3,000 of a heavier one, 2,000 with a LET in it (the set
 * `lambda-depth`). This one is where the JavaScript stack of a process that has not warmed up holds with room to spare
 * (500 calls of that calculation, 300 with a LET, in Node), so that a recursion ends in the same place in every process
 * and both execution modes; a stack that does run out first is #NUM! too.
 */
export const CALL_BUDGET = 1000

let depth = 0

/**
 * A LAMBDA of `count` parameters whose calculation `run` computes: a call gives each of them, one left out between commas
 * (`f(1,)`) too, as Excel takes no fewer; each call counts as a unit of the computation's work.
 */
export function lambdaOf(count: number, height: number, run: (args: readonly (Arg | undefined)[]) => Arg): Lambda {
  return new Lambda(count, count, (args) => {
    if (depth + height > CALL_BUDGET) return E.NUM
    spend(1)
    depth += height
    try {
      return run(args as readonly (Arg | undefined)[])
    } catch (e) {
      if (e instanceof RangeError && !(e instanceof ArraySizeError)) return E.NUM
      throw e
    } finally {
      depth -= height
    }
  })
}

/** A value called with `args`: a function within its counts, #VALUE! for wrong counts or a value that is no function, an error as it is. */
export function callValue(f: Arg, args: readonly (Arg | undefined)[]): Arg {
  if (f instanceof Lambda) return args.length < f.min || args.length > f.max ? E.VALUE : (f.run(args) as Arg)
  return isErr(f) ? f : E.VALUE
}
