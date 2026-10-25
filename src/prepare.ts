/**
 * A formula prepared for the engine: checked against the workbook (its sheets, names, tables, functions and links), its
 * canonical text, its AST with linked workbooks' names put in their place (`expandLinkNames`). While a workbook's state
 * is built, the formulas of one shape (`syntax/shape.ts`) are prepared once.
 */
import { MAX_FORMULA_LENGTH, type CheckCode, type ErrorCode } from './contract.js'
import type { EngineState, Invalid } from './engine.js'
import { expandLinkNames } from './links.js'
import { nameFor } from './names.js'
import { FormulaError, printFormula, sheetKey, tokenize, type Ast, type PrintOptions, type Tok } from './syntax.js'
import { validateTokens } from './syntax/check.js'
import { instance, ready, shapeOf, type Template } from './syntax/shape.js'
import { tableAt, tableNames } from './tables/model.js'

export type Prepared = { ok: true; text: string; ast: Ast } | { ok: false; invalid: Invalid }

export function invalidError(code: CheckCode): ErrorCode {
  if (code === 'unknown_function' || code === 'unknown_name') return 'NAME'
  if (code === 'unknown_sheet' || code === 'external_ref') return 'REF'
  return 'ERROR'
}

type Named = (name: string, sheet: string | undefined) => ReturnType<typeof nameFor>

/** The names a formula of sheet `s` sees, by a name and the sheet it is seen from (undefined: `s`). */
const namedFrom = (state: EngineState, s: number): Named => (name, sheet) => nameFor(state.names, sheet === undefined ? s : state.sheetIdx.get(sheetKey(sheet))!, name)

/** Sheets and names spelled as the workbook has them, as Excel shows them (`ставка*2` is `Ставка*2`). */
const spelling = (state: EngineState, named: Named): PrintOptions => ({
  sheetOf: (sh) => state.sheetNames[state.sheetIdx.get(sheetKey(sh))!],
  nameOf: (name, sheet) => named(name, sheet)?.name ?? name,
})

/** The check of a formula in the table named `enclosing` (none: outside tables), over its tokens when they are read (`toks`). */
function checked(state: EngineState, enclosing: string | undefined, text: string, toks: Tok[] | null, named: Named): { ok: true; ast: Ast } | { ok: false; invalid: Invalid } {
  const known = (name: string, sheet: string | undefined) => !!named(name, sheet)
  const r = validateTokens(text, toks, { sheets: state.sheetNames, known, externalFunctions: state.externalFunctions, tables: tableNames(state.tables), externalBooks: state.input.externalBooks ?? [], ...(enclosing ? { table: enclosing } : {}) })
  return r.ok ? r : { ok: false, invalid: { code: r.code, message: r.message, at: r.at, error: invalidError(r.code) } }
}

/**
 * How `validate` checks a formula of sheet `s` against the state's book: its sheets, names, functions and tables, and
 * the table the formula is in (`enclosing`). Without the book's linked workbooks: a conditional format's or a data
 * validation's formula reads none of them, so one that names them fails the check.
 */
export const validateIn = (state: EngineState, s: number, enclosing?: string) => ({
  sheets: state.sheetNames,
  known: (name: string, sheet: string | undefined) => !!namedFrom(state, s)(name, sheet),
  externalFunctions: state.externalFunctions,
  tables: tableNames(state.tables),
  ...(enclosing ? { table: enclosing } : {}),
})

/** A formula's AST checked, with its text and linked workbooks' names put in their place. */
const accepted = (state: EngineState, text: string, ast: Ast): Prepared => ({ ok: true, text, ast: expandLinkNames(ast, state.links, state.input.externalBooks) })

/** A formula of sheet `s` at `row`, `col` prepared on its own. */
export function prepare(state: EngineState, s: number, row: number, col: number, text: string, toks: Tok[] | null = null): Prepared {
  const named = namedFrom(state, s)
  const r = checked(state, tableAt(state.tables, s, row, col)?.name, text, toks, named)
  return r.ok ? accepted(state, printFormula(r.ast, spelling(state, named)), r.ast) : r
}

/** The shapes of one build of a state, by their key: a template, or null for a shape prepared formula by formula. */
export type Shapes = Map<string, Template | null>

/** At most this many shapes are kept for a build: a workbook of formulas that share none keeps no more. */
const SHAPES = 4096

/** `prepare` through the formula's shape: the shape's template taken when `shapes` holds it, else the formula prepared and its shape kept. */
export function prepareShaped(state: EngineState, shapes: Shapes, s: number, row: number, col: number, text: string): Prepared {
  // a text too long fails before it is read; a text that does not read fails as on its own
  if (text.length > MAX_FORMULA_LENGTH) return prepare(state, s, row, col, text)
  let toks: Tok[]
  try {
    toks = tokenize(text)
  } catch (e) {
    if (e instanceof FormulaError) return prepare(state, s, row, col, text)
    throw e
  }
  const shape = shapeOf(toks)
  if (!shape) return prepare(state, s, row, col, text, toks)
  const table = tableAt(state.tables, s, row, col)?.name
  const key = `${s}:${table === undefined ? '-' : `${table.length}:${table}`}:${shape.key}`
  const t = shapes.get(key)
  if (t && (t.at || ready(t, spelling(state, namedFrom(state, s))))) {
    const x = instance(t, shape.holes)
    return accepted(state, x.text, x.ast)
  }
  if (t !== undefined) {
    // a rejected shape, or a template its formula's AST or text cannot take holes in
    shapes.set(key, null)
    return prepare(state, s, row, col, text, toks)
  }
  const named = namedFrom(state, s)
  const r = checked(state, table, text, toks, named)
  if (shapes.size < SHAPES) shapes.set(key, r.ok ? { ast: r.ast, holes: shape.holes } : null)
  return r.ok ? accepted(state, printFormula(r.ast, spelling(state, named)), r.ast) : r
}
