/**
 * The check of a formula's calls by the functions' declared arguments (`function-arity.ts`) for the calls that check
 * formulas without the engine (the root's and `xlcalc/entry`'s). A host's bundler drops the engine with its registration
 * of the specs (`functions.ts`, `"sideEffects": false`), so each such call gives the table itself, unless the engine
 * gave its specs first.
 */
import { aritiesOf, useAritiesUnlessGiven } from './arity.js'
import { FUNCTION_ARITY } from './function-arity.js'

const declared = () => aritiesOf(FUNCTION_ARITY)

/** `f`, by its name and number of parameters, checking a formula's calls by the declared arguments where the engine gave none. */
export function withDeclaredArities<F extends (...args: never[]) => unknown>(f: F): F {
  const g = function (this: unknown, ...args: never[]) {
    useAritiesUnlessGiven(declared)
    return f.apply(this, args)
  }
  return Object.defineProperties(g, { name: { value: f.name }, length: { value: f.length } }) as unknown as F
}
