/**
 * Formulas of a file in the engine's spelling, with what the engine cannot take reported on their cell or name. Of the
 * formulas of one shape (`syntax/shape.ts`), a formula copied down a column, one is parsed for a file.
 */
import { MAX_FORMULA_LENGTH } from '../../contract.js'
import { FormulaError, type Ast } from '../../syntax/ast.js'
import { parseTokens } from '../../syntax/parse.js'
import { shapeOf } from '../../syntax/shape.js'
import { tokenize, type Tok } from '../../syntax/tokens.js'
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
  /** The engine parses the formula. */
  ok: boolean
  /** The syntax tree, when the engine parses the formula and it was asked for. */
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
  if (!f.ok && !f.found.some((x) => EXPLAINED.has(x))) issues.add('formula', { key: 'formulaUnparsed' }, where)
}

/** The shapes of the formulas read with one options object, a file's: true for one that parsed, false for one parsed formula by formula. */
const shapesOf = new WeakMap<FormulaOptions, Map<string, boolean>>()

/** At most this many shapes are kept for a file, as the engine keeps (`prepare.ts`). */
const SHAPES = 4096

/**
 * Whether the formula parses, and its tree where `tree` asks for it: a formula of a shape a formula of the same file
 * (`opts`) parsed before parses too, and is not parsed again.
 */
function parsed(text: string, opts: FormulaOptions | undefined, tree: boolean): { ok: boolean; ast?: Ast } {
  const own = (toks: Tok[] | null) => {
    try {
      return { ok: true, ast: parseTokens(text, toks, { union: opts?.union }) }
    } catch (e) {
      if (e instanceof FormulaError) return { ok: false }
      throw e
    }
  }
  // a name's formula, and a text too long or empty, which fails before it is read
  if (tree || !opts || opts.union || text.length > MAX_FORMULA_LENGTH || text.trim() === '') return own(null)
  let toks: Tok[]
  try {
    toks = tokenize(text)
  } catch (e) {
    if (e instanceof FormulaError) return { ok: false }
    throw e
  }
  const shape = shapeOf(toks)
  if (!shape) return own(toks)
  let shapes = shapesOf.get(opts)
  if (!shapes) shapesOf.set(opts, (shapes = new Map()))
  if (shapes.get(shape.key)) return { ok: true }
  // a shape that failed is parsed formula by formula, as the engine prepares it
  const r = own(toks)
  if (shapes.size < SHAPES) shapes.set(shape.key, r.ok)
  return r
}

/**
 * A formula of a file in the engine's spelling, parsed when the engine parses it, with nothing reported yet; its tree
 * left out where `tree` is false and its shape parsed before.
 */
export function parseFormula(text: string, opts?: FormulaOptions, tree = true): ReadFormula {
  const { links, ...c } = fromFile(text)
  if (opts?.rename) c.formula = opts.rename(c.formula)
  const r = parsed(c.formula, opts, tree)
  if (!r.ok) return { ...c, ok: false }
  const linked = links?.every((n) => opts?.links?.has(n))
  return { ...c, ok: true, ...(r.ast ? { ast: r.ast } : {}), ...(linked ? { found: c.found.filter((x) => x !== 'externalLink') } : {}) }
}

/** A formula read at `where`, with its tree where `tree` asks for it (`parseFormula`). */
export function readFormula(text: string, issues: Issues, where: Where, opts?: FormulaOptions, tree = false): ReadFormula {
  const f = parseFormula(text, opts, tree)
  reportFormula(f, issues, where)
  return f
}
