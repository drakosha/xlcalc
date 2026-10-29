/**
 * Functions as values: a LAMBDA, or a built-in function by its name (`BYROW(A1:C3,SUM)`). Wherever a value goes one is
 * #CALC!, as Excel shows a LAMBDA left without a call, so it is an error to every other function; a call gives it its
 * arguments (`LAMBDA(x,x*2)(3)`), MAP and its kin call it for each value (`functions/lambda.ts`).
 */
import type { Arg } from '../args.js'
import { CErr, E, isErr } from '../values.js'
import { ArraySizeError, spend } from './budget.js'

export class Lambda extends CErr {
  constructor(
    /** The fewest and the most arguments it takes; a parameter left out at the end is omitted (ISOMITTED). */
    readonly min: number,
    readonly max: number,
    /** The call with arguments within min…max, undefined for one left out. */
    readonly run: (args: readonly (Arg | undefined)[]) => Arg,
  ) {
    super('CALC')
  }
}

/** The value of a parameter left out of the call: blank where it is read, and ISOMITTED tells it from a blank given. */
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

/** A LAMBDA of `count` parameters whose calculation `run` computes: each call counts as a unit of the computation's work. */
export function lambdaOf(count: number, run: (args: readonly (Arg | undefined)[]) => Arg): Lambda {
  return new Lambda(0, count, (args) => {
    if (depth >= MAX_CALL_DEPTH) return E.NUM
    spend(1)
    depth++
    try {
      return run(args)
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
  if (f instanceof Lambda) return args.length < f.min || args.length > f.max ? E.VALUE : f.run(args)
  return isErr(f) ? f : E.VALUE
}
