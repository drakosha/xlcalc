/** Formulas of a file in the engine's spelling, with what the engine cannot take reported on their cell or name. */
import { FormulaError, type Ast } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'
import { fromFile, type Found } from '../formulas/from-file.js'
import type { Issues, Where } from '../issues.js'
import type { XlsxIssueCode } from '../types.js'

const FOUND: Record<Found, [XlsxIssueCode, string]> = {
  singleOperator: ['singleOperator', 'the @ operator (_xlfn.SINGLE) is not in the engine: the formula computes to #NAME?'],
  userFunction: ['userFunction', 'a user-defined function (a macro or an add-in) is not in the engine'],
  table: ['table', 'structured references to tables are not in the engine: the formula is read as written and computes to an error'],
  externalLink: ['externalLink', 'references to other workbooks are not in the engine: the formula is read as written and computes to an error'],
  reference3d: ['reference3d', 'references across sheets (Sheet1:Sheet3!A1) are not in the engine: the formula is read as written and computes to an error'],
  letName: ['formula', 'a LET name that reads as a reference or a boolean once its _xlpm. prefix goes: the formula changes its meaning'],
}

/** Found kinds that already say why the engine does not parse the formula. */
const EXPLAINED = new Set<Found>(['table', 'externalLink', 'reference3d', 'letName'])

export interface ReadFormula {
  formula: string
  /** The syntax tree, when the engine parses the formula. */
  ast?: Ast
}

export function readFormula(text: string, issues: Issues, where: Where): ReadFormula {
  const c = fromFile(text)
  for (const f of c.found) issues.add(...FOUND[f], where)
  try {
    return { formula: c.formula, ast: parse(c.formula) }
  } catch (e) {
    if (!(e instanceof FormulaError)) throw e
    if (!c.found.some((f) => EXPLAINED.has(f))) issues.add('formula', 'a formula the engine does not parse is read as written (its cell computes to an error)', where)
    return { formula: c.formula }
  }
}
