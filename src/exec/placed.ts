/**
 * Where a reference found as the formula computes is used for its place alone (`ROW(INDIRECT("1:9"))`,
 * `ROWS(A1:OFFSET(A1,9,0))`, OFFSET's own reference): the cells there are not read, so the graph does not learn them
 * (`graph/dynamic.ts`), as Excel finds no circular reference in such a formula over its own cell. The walk compiles
 * such an INDIRECT, OFFSET or range operator as its quiet form (`QUIET`).
 */
import { FUNCTIONS } from '../functions.js'
import type { Ast } from '../syntax.js'

/** The suffix of a function's name in its quiet form, which no formula can spell. */
export const QUIET = '#'

/** The range operator in its quiet form. */
export const QUIET_RANGE = '::'

/** The arguments of functions that read where a reference is, not its cells. */
const PLACE: Readonly<Record<string, (i: number, n: Extract<Ast, { k: 'fn' }>) => boolean>> = {
  ROW: (i) => i === 0,
  COLUMN: (i) => i === 0,
  ROWS: (i) => i === 0,
  COLUMNS: (i) => i === 0,
  AREAS: (i) => i === 0,
  ISREF: (i) => i === 0,
  SHEET: (i) => i === 0,
  SHEETS: (i) => i === 0,
  OFFSET: (i) => i === 0,
  CELL: (i, n) => i === 1 && n.args[0]?.k === 'str' && /^(address|row|col)$/i.test(n.args[0].v),
}

/** The argument at `i` passes on the reference the call answers: INDEX's, TAKE's, DROP's and TRIMRANGE's array, IF's and CHOOSE's values. */
function passes(n: Extract<Ast, { k: 'fn' }>, i: number): boolean {
  if (n.name === 'INDEX' || n.name === 'TAKE' || n.name === 'DROP' || n.name === 'TRIMRANGE') return i === 0
  const spec = FUNCTIONS[n.name]
  return !!spec?.refs && spec.mode(i) === 'pass'
}

/** The nodes of a formula whose reference is used for its place alone; names and LET's values are not looked into. */
export function placedOnly(ast: Ast): Set<Ast> {
  const out = new Set<Ast>()
  const visit = (n: Ast, place: boolean): void => {
    if (place) out.add(n)
    switch (n.k) {
      case 'paren':
      case 'at':
        return visit(n.a, place)
      case 'bin':
        visit(n.a, place && n.op === ':')
        return visit(n.b, place && n.op === ':')
      case 'un':
      case 'pct':
        return visit(n.a, false)
      case 'union':
        return n.items.forEach((x) => visit(x, place))
      case 'fn':
        return n.args.forEach((a, i) => visit(a, PLACE[n.name]?.(i, n) || (place && passes(n, i))))
      case 'let':
        n.binds.forEach((b) => visit(b.value, false))
        if (n.body) visit(n.body, place)
        return
    }
  }
  visit(ast, false)
  return out
}
