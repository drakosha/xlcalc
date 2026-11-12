/** Formula idioms Excel treats apart, recognized by their text alone: shared by legacy conversion and the CSE plan. */
import type { Ast } from './ast.js'

const inner = (n: Ast): Ast => (n.k === 'paren' ? inner(n.a) : n)

/** The call of one of `names` that `n` is, through parentheses. */
export const callOf = (n: Ast, ...names: string[]) => {
  const x = inner(n)
  return x.k === 'fn' && names.includes(x.name) ? x : undefined
}

/**
 * `N(IF(…))` or `T(IF(…))` as a lookup's key or place: the old idiom that hands the lookup every value at once, in
 * a legacy array formula and in a formula of the old kind alike (measured).
 */
export const dereferenced = (n: Ast) => {
  const x = callOf(n, 'N', 'T')
  return x !== undefined && x.args.length === 1 && callOf(x.args[0], 'IF') !== undefined
}
