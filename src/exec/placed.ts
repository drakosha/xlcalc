/**
 * Where a formula uses a reference for its place alone (`ROW(A1)`, `ROWS(A1:OFFSET(A1,9,0))`, OFFSET's own reference,
 * CELL's address or width) or for the formula there (FORMULATEXT, ISFORMULA): the cells' values are not read, so the
 * order of computation does not wait for them and a formula over its own cell is no circular reference, as Excel
 * computes it (the set `place-cycles-held`). The graph keeps such a reference as a listener alone (`FormulaInfo.near`);
 * an INDIRECT, OFFSET or range operator used for its place alone is compiled in its quiet form (`QUIET`), which tells
 * the graph nothing it finds.
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
  CELL: (i, n) => i === 1 && n.args[0]?.k === 'str' && /^(address|row|col|width|format|color|parentheses|protect|filename)$/i.test(n.args[0].v),
}

/** The arguments of functions that read the formula a cell holds, not its value. */
const FORMULA: Readonly<Record<string, (i: number) => boolean>> = { FORMULATEXT: (i) => i === 0, ISFORMULA: (i) => i === 0 }

/** How a node's reference is used: for its place alone, or for the formulas there. */
export type Use = 'place' | 'formula'

/**
 * The argument at `i` passes on the reference the call answers: INDEX's, TAKE's and DROP's array, IF's and CHOOSE's
 * values (TRIMRANGE reads its range's cells to trim it).
 */
function passes(n: Extract<Ast, { k: 'fn' }>, i: number): boolean {
  if (n.name === 'INDEX' || n.name === 'TAKE' || n.name === 'DROP') return i === 0
  const spec = FUNCTIONS[n.name]
  return !!spec?.refs && spec.mode(i) === 'pass'
}

/** The nodes of a formula whose reference is used for its place or its formulas alone; names and LET's values are not looked into. */
export function placedOnly(ast: Ast): Map<Ast, Use> {
  const out = new Map<Ast, Use>()
  const visit = (n: Ast, use: Use | undefined): void => {
    if (use) out.set(n, use)
    switch (n.k) {
      case 'paren':
      case 'at':
        return visit(n.a, use)
      case 'bin':
        visit(n.a, n.op === ':' ? use : undefined)
        return visit(n.b, n.op === ':' ? use : undefined)
      case 'un':
      case 'pct':
        return visit(n.a, undefined)
      case 'union':
        return n.items.forEach((x) => visit(x, use))
      case 'fn':
        return n.args.forEach((a, i) => visit(a, PLACE[n.name]?.(i, n) ? 'place' : FORMULA[n.name]?.(i) ? 'formula' : use && passes(n, i) ? use : undefined))
      case 'let':
        n.binds.forEach((b) => visit(b.value, undefined))
        if (n.body) visit(n.body, use)
        return
    }
  }
  visit(ast, undefined)
  return out
}

/** Whether a call reads references for their place or formulas, so that the walk looks for such uses (`placedOnly`). */
export const placing = (name: string) => name in PLACE || name in FORMULA
