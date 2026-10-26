/** Formulas of a file in the engine's spelling, with what the engine cannot take reported on their cell or name. */
import { FormulaError, type Ast } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'
import { fromFile, type Found } from '../formulas/from-file.js'
import type { Issues, Where, XlsxMessage } from '../issues.js'
import type { XlsxIssueCode } from '../types.js'

const FOUND: Record<Found, [XlsxIssueCode, XlsxMessage]> = {
  userFunction: ['userFunction', { key: 'userFunction' }],
  table: ['table', { key: 'tableRefWithoutTable' }],
  externalLink: ['externalLink', { key: 'formulaExternal' }],
  letName: ['formula', { key: 'letPrefix' }],
}

/** Found kinds that already say why the engine does not parse the formula. */
const EXPLAINED = new Set<Found>(['table', 'externalLink', 'letName'])

export interface ReadFormula {
  formula: string
  /** The syntax tree, when the engine parses the formula. */
  ast?: Ast
  found: Found[]
  /** The user-defined functions it calls, as `fromFile` finds them. */
  userFunctions?: string[]
}

/** How a formula of the file is read. */
export interface FormulaOptions {
  /** A name's formula, which may be a union of references without parentheses. */
  union?: boolean
  /** The numbers `[N]` of the book's links to other workbooks (`readLinks`): a formula the engine parses that refers to these alone is in it. */
  links?: ReadonlySet<number>
  /** References to sheets read under other names than the file's (`sheet-names.ts`) rewritten. */
  rename?: (formula: string) => string
}

/**
 * Reports what the engine cannot take of a formula read, at `where`: a shared formula's copies report their master's;
 * each user-defined function on its own (`XlsxIssue.fn`).
 */
export function reportFormula(f: ReadFormula, issues: Issues, where: Where): void {
  for (const x of f.found) {
    if (x === 'userFunction') for (const fn of f.userFunctions ?? [undefined]) issues.add(...FOUND[x], { ...where, fn })
    else issues.add(...FOUND[x], where)
  }
  if (!f.ast && !f.found.some((x) => EXPLAINED.has(x))) issues.add('formula', { key: 'formulaUnparsed' }, where)
}

/** A formula of a file in the engine's spelling, parsed when the engine parses it, with nothing reported yet. */
export function parseFormula(text: string, opts: FormulaOptions = {}): ReadFormula {
  const { links, ...c } = fromFile(text)
  if (opts.rename) c.formula = opts.rename(c.formula)
  let ast: Ast
  try {
    ast = parse(c.formula, { union: opts.union })
  } catch (e) {
    if (!(e instanceof FormulaError)) throw e
    return c
  }
  const linked = links?.every((n) => opts.links?.has(n))
  return { ...c, ast, ...(linked ? { found: c.found.filter((x) => x !== 'externalLink') } : {}) }
}

/** A formula read at `where`. */
export function readFormula(text: string, issues: Issues, where: Where, opts?: FormulaOptions): ReadFormula {
  const f = parseFormula(text, opts)
  reportFormula(f, issues, where)
  return f
}
