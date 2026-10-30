/**
 * The check of a formula's calls by the functions' declared arguments (`function-arity.ts`) for the calls that check
 * formulas without the engine (the root's and `xlcalc/entry`'s), and of the functions the engine has not loaded. A host's
 * bundler drops the engine with its registration of the specs (`functions.ts`, `"sideEffects": false`), so each such
 * call gives the table itself.
 */
import { aritiesOf, SYNTAX, useDeclaredArities } from './arity.js'
import { before } from './strict-args.js'
import { FUNCTION_ARITY } from './function-arity.js'

/** The whitelist: every function a formula can call, by its English name, in alphabetical order, whichever groups are loaded. */
export const FUNCTION_NAMES: readonly string[] = /* @__PURE__ */ [...Object.keys(FUNCTION_ARITY), ...SYNTAX].sort()

const declared = () => aritiesOf(FUNCTION_ARITY)

/** `f`, by its name and number of parameters, checking a formula's calls by the declared arguments of the functions the engine has not. */
export const withDeclaredArities = <F extends (...args: never[]) => unknown>(f: F): F => before(f, () => useDeclaredArities(declared))
