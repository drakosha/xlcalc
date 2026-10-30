/**
 * A formula prepared for the engine: checked against the workbook (its sheets, names, tables, functions and links), its
 * canonical text, its AST with linked workbooks' names put in their place (`expandLinkNames`). While a workbook's state
 * is built, the formulas of one shape (`syntax/shape.ts`) are prepared once.
 */
import type { EngineState, Invalid } from './engine.js'
import { expandLinkNames } from './links.js'
import { nameFor } from './names.js'
import { printFormula, sheetKey, type Ast, type PrintOptions, type Tok } from './syntax.js'
import { invalidError, validateTokens } from './syntax/check.js'
import { instance, ready, shaped, SHAPES, tableShapeKey, type Template } from './syntax/shape.js'
import { tableAt, tableNames } from './tables/model.js'

/**
 * A prepared formula; `unknown` the check's refusal of a name the workbook does not define or a function the engine
 * does not know or has not loaded, which the formula is calculated over all the same: #NAME? where it is computed, as in Excel
 * (`IF(TRUE,1,Nope)` is 1); `holds` when the formula keeps the value its file stores over such a function
 * (`unknownFunctions: 'held'`).
 */
export type Prepared = { ok: true; text: string; ast: Ast; unknown?: Invalid; holds?: true } | { ok: false; invalid: Invalid }

type Named = (name: string, sheet: string | undefined) => ReturnType<typeof nameFor>

/** The names a formula of sheet `s` sees, by a name and the sheet it is seen from (undefined: `s`). */
const namedFrom = (state: EngineState, s: number): Named => (name, sheet) => nameFor(state.names, sheet === undefined ? s : state.sheetIdx.get(sheetKey(sheet))!, name)

/** Sheets and names spelled as the workbook has them, as Excel shows them (`ставка*2` is `Ставка*2`). */
const spelling = (state: EngineState, named: Named): PrintOptions => ({
  sheetOf: (sh) => state.sheetNames[state.sheetIdx.get(sheetKey(sh))!],
  nameOf: (name, sheet) => named(name, sheet)?.name ?? name,
})

/** The check of a formula in the table named `enclosing` (none: outside tables), over its tokens when they are read (`toks`). */
function checked(state: EngineState, enclosing: string | undefined, text: string, toks: Tok[] | null, named: Named): { ok: true; ast: Ast; unknown?: Invalid } | { ok: false; invalid: Invalid } {
  const check = (lenient: boolean) =>
    validateTokens(text, toks, { sheets: state.sheetNames, known: (name, sheet) => !!named(name, sheet), externalFunctions: state.externalFunctions, tables: tableNames(state.tables), externalBooks: state.input.externalBooks ?? [], notLoaded: state.groups.notLoaded, ...(enclosing ? { table: enclosing } : {}), ...(lenient ? { unknown: true, anyName: true } : {}), ...(state.ctx.hosts ? { answered: true } : {}) })
  const r = check(false)
  if (r.ok) return r
  const invalid: Invalid = { code: r.code, message: r.message, at: r.at, error: invalidError(r.code), ...(r.fn !== undefined ? { fn: r.fn } : {}) }
  // a name the book does not define, or a function the engine does not know or has not loaded, fails only where it is
  // computed, as in Excel: checked again taking every name and function, the formula keeps the refusal
  if (r.code !== 'unknown_name' && r.code !== 'unknown_function' && r.code !== 'function_not_loaded') return { ok: false, invalid }
  const again = check(true)
  return again.ok ? { ok: true, ast: again.ast, unknown: invalid } : { ok: false, invalid }
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
  notLoaded: state.groups.notLoaded,
  ...(enclosing ? { table: enclosing } : {}),
})

/** A formula's AST checked, with its text and linked workbooks' names put in their place, and the refusal of a name it keeps. */
const accepted = (state: EngineState, text: string, ast: Ast, unknown?: Invalid): Prepared => ({ ok: true, text, ast: expandLinkNames(ast, state.links, state.input.externalBooks), ...(unknown ? { unknown } : {}) })

/** A formula of sheet `s` at `row`, `col` prepared on its own. */
export function prepare(state: EngineState, s: number, row: number, col: number, text: string, toks: Tok[] | null = null): Prepared {
  const named = namedFrom(state, s)
  const r = checked(state, tableAt(state.tables, s, row, col)?.name, text, toks, named)
  return r.ok ? accepted(state, printFormula(r.ast, spelling(state, named)), r.ast, r.unknown) : r
}

/** A formula over a function the engine does not know or has not loaded that keeps the value its file stores (`held`) under `unknownFunctions: 'held'`; otherwise `p` as it is. */
export const preparedHolding = (state: EngineState, p: Prepared, held: unknown): Prepared =>
  p.ok && (p.unknown?.code === 'unknown_function' || p.unknown?.code === 'function_not_loaded') && state.unknownHeld && held !== undefined ? { ...p, holds: true } : p

/** The shapes of one build of a state, by their key: a template, or null for a shape prepared formula by formula. */
export type Shapes = Map<string, Template | null>

/** `prepare` through the formula's shape: the shape's template taken when `shapes` holds it, else the formula prepared and its shape kept. */
export function prepareShaped(state: EngineState, shapes: Shapes, s: number, row: number, col: number, text: string): Prepared {
  // a text too long, blank or not read fails as on its own
  const read = shaped(text)
  if (read === null || read === 'unread') return prepare(state, s, row, col, text)
  const { toks, shape } = read
  if (!shape) return prepare(state, s, row, col, text, toks)
  const table = tableAt(state.tables, s, row, col)?.name
  const key = `${s}:${tableShapeKey(table, shape.key)}`
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
  // a formula over a name the book does not define is prepared on its own, with its refusal
  if (shapes.size < SHAPES) shapes.set(key, r.ok && !r.unknown ? { ast: r.ast, holes: shape.holes } : null)
  return r.ok ? accepted(state, printFormula(r.ast, spelling(state, named)), r.ast, r.unknown) : r
}
