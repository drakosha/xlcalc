/** Formulas of a file in the engine's spelling, with what the engine cannot take reported on their cell or name. */
import { FormulaError, type Ast } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'
import { fromFile, type Found } from '../formulas/from-file.js'
import type { Issues, Where } from '../issues.js'
import type { XlsxIssueCode } from '../types.js'

const FOUND: Record<Found, [XlsxIssueCode, string]> = {
  userFunction: ['userFunction', 'a user-defined function (a macro or an add-in) is not in the engine'],
  table: ['table', 'a structured reference without a table name is not in a file: the formula is read as written and computes to an error'],
  externalLink: ['externalLink', 'references to other workbooks are not in the engine: the formula is read as written and computes to an error'],
  letName: ['formula', 'a LET name that reads as a reference or a boolean once its _xlpm. prefix goes: the formula changes its meaning'],
}

/** Found kinds that already say why the engine does not parse the formula. */
const EXPLAINED = new Set<Found>(['table', 'externalLink', 'letName'])

export interface ReadFormula {
  formula: string
  /** The syntax tree, when the engine parses the formula. */
  ast?: Ast
  found: Found[]
}

/** Reports what the engine cannot take of a formula read, at `where`: a shared formula's copies report their master's. */
export function reportFormula(f: ReadFormula, issues: Issues, where: Where): void {
  for (const x of f.found) issues.add(...FOUND[x], where)
  if (!f.ast && !f.found.some((x) => EXPLAINED.has(x))) issues.add('formula', 'a formula the engine does not parse is read as written (its cell computes to an error)', where)
}

/** A formula of a file in the engine's spelling, parsed when the engine parses it, with nothing reported yet. */
export function parseFormula(text: string, union = false): ReadFormula {
  const c = fromFile(text)
  try {
    return { ...c, ast: parse(c.formula, { union }) }
  } catch (e) {
    if (!(e instanceof FormulaError)) throw e
    return c
  }
}

/** A formula read at `where`; `union` for a name's formula, which may be a union of references without parentheses. */
export function readFormula(text: string, issues: Issues, where: Where, union = false): ReadFormula {
  const f = parseFormula(text, union)
  reportFormula(f, issues, where)
  return f
}
