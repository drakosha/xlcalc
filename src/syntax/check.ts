/** Checks of a formula with codes: functions and their argument counts, sheets, names; valid range names. */
import type { CheckCode, NameCheck, ValidateContext, ValidateResult } from '../contract.js'
import { FUNCTIONS, isFunction } from '../functions.js'
import { FormulaError, type Ast } from './ast.js'
import { parse } from './parse.js'
import { canonical, mapAst, sheetKey, walk } from './print.js'

/** Argument counts of each function and the arguments it cannot leave out; literal arguments of external functions. */
export function checkCalls(ast: Ast, src: string, external: (fn: string) => boolean): FormulaError | null {
  const at = (needle: string) => Math.max(0, src.toUpperCase().indexOf(needle.toUpperCase()))
  const visit = (n: Ast): FormulaError | null => {
    switch (n.k) {
      case 'paren':
      case 'un':
      case 'pct':
      case 'at':
        return visit(n.a)
      case 'union':
        for (const a of n.items) {
          const e = visit(a)
          if (e) return e
        }
        return null
      case 'bin':
        return visit(n.a) ?? visit(n.b)
      case 'let':
        for (const a of [...n.binds.map((b) => b.value), ...(n.body ? [n.body] : [])]) {
          const e = visit(a)
          if (e) return e
        }
        return null
      case 'fn': {
        if (external(n.name)) {
          for (const a of n.args) {
            const lit = a.k === 'num' || a.k === 'str' || a.k === 'bool' || (a.k === 'un' && a.a.k === 'num')
            if (!lit) return new FormulaError('unsupported', `arguments of ${n.name} must be literals`, at(n.name))
          }
          return null
        }
        const spec = FUNCTIONS[n.name]
        if (!spec) return null
        const k = n.args.length
        if (!spec.takes(k)) return new FormulaError('syntax', `${n.name} takes ${spec.counts} arguments, got ${k}`, at(n.name))
        const left = n.args.findIndex((a, i) => a.k === 'missing' && !spec.omits(i))
        if (left >= 0) return new FormulaError('syntax', `${n.name} takes a reference as its argument ${left + 1}, which cannot be left out`, at(n.name))
        for (const a of n.args) {
          const e = visit(a)
          if (e) return e
        }
        return null
      }
      default:
        return null
    }
  }
  return visit(ast)
}

/** What the engine adds to a check: a name's formula (`union`), which names a book has as a sheet sees them (`known`), the table of the formula's cell (`enclosing`). */
interface BookCheck {
  union?: boolean
  known?: (name: string, sheet: string | undefined) => boolean
  enclosing?: string
}

type Fail = Extract<ValidateResult, { ok: false }>

/**
 * Structured references bound to the tables of `ctx.tables`: a reference without a table name takes the formula's
 * table, tables and columns are spelled as named, a bare table name is the table's data (`Sales[]`, as a file stores
 * it). An unknown table or column, or a reference without a table name outside one, fails with `unknown_name`.
 */
export function bindTables(ast: Ast, text: string, ctx: ValidateContext & BookCheck): { ok: true; ast: Ast; tables: string[] } | Fail {
  const known = ctx.tables ? new Map(ctx.tables.map((t) => [sheetKey(t.name), t])) : null
  const used = new Set<string>()
  let fail: Fail | null = null
  const at = (needle: string) => Math.max(0, text.toUpperCase().indexOf(needle.toUpperCase()))
  const out = mapAst(ast, (n) => {
    if (fail) return undefined
    if (n.k === 'name' && n.sheet === undefined && known?.has(sheetKey(n.name))) {
      const t = known.get(sheetKey(n.name))!
      used.add(t.name)
      return { k: 'table', table: t.name, spec: { areas: [] } }
    }
    if (n.k !== 'table') return undefined
    const name = n.table ?? ctx.enclosing
    if (name === undefined) {
      if (ctx.tables) fail = { ok: false, code: 'unknown_name', message: 'a structured reference without a table name outside a table', at: at('[') }
      return undefined
    }
    const t = known?.get(sheetKey(name))
    if (!t) {
      if (known) fail = { ok: false, code: 'unknown_name', message: `unknown table ${name}`, at: at(name) }
      else used.add(name)
      return undefined
    }
    used.add(t.name)
    const cols = n.spec.cols?.map((c) => t.columns.find((x) => sheetKey(x) === sheetKey(c)))
    const missing = n.spec.cols?.find((c, i) => cols![i] === undefined)
    if (missing !== undefined) {
      fail = { ok: false, code: 'unknown_name', message: `the table ${t.name} has no column ${missing}`, at: at(missing) }
      return undefined
    }
    return { k: 'table', table: t.name, spec: cols ? { ...n.spec, cols: cols as [string] | [string, string] } : n.spec }
  })
  return fail ?? { ok: true, ast: out, tables: [...used] }
}

/** Formula check: parsing, whitelisted or external functions, sheets, names, argument counts. */
export function validate(text: string, ctx: ValidateContext & BookCheck = {}): ({ ok: true; ast: Ast } & Extract<ValidateResult, { ok: true }>) | Extract<ValidateResult, { ok: false }> {
  let ast: Ast
  try {
    ast = parse(text, { union: ctx.union })
  } catch (e) {
    if (e instanceof FormulaError) return { ok: false, code: e.code, message: e.message, at: e.at }
    throw e
  }
  const bound = bindTables(ast, text, ctx)
  if (!bound.ok) return bound
  ast = bound.ast
  const externalSet = new Set((ctx.externalFunctions ?? []).map((f) => f.toUpperCase()))
  const functions = new Set<string>()
  const sheets = new Set<string>()
  const names = new Set<string>()
  // names with the sheet they are seen from (undefined: the formula's own)
  const uses: [string, string | undefined][] = []
  const externals: string[] = []
  walk(ast, (n) => {
    if (n.k === 'fn') {
      functions.add(n.name)
      if (externalSet.has(n.name) && !FUNCTIONS[n.name]) externals.push(canonical(n))
    } else if (n.k === 'let') functions.add('LET')
    else if ((n.k === 'ref' && n.sheet !== null) || (n.k === 'err' && n.sheet !== undefined)) sheets.add(n.sheet!)
    else if (n.k === 'ref3d') sheets.add(n.first).add(n.last)
    else if (n.k === 'name') {
      names.add(n.name)
      uses.push([n.name, n.sheet])
      if (n.sheet !== undefined) sheets.add(n.sheet)
    }
  })
  const fail = (code: CheckCode, message: string, needle: string) => ({ ok: false as const, code, message, at: Math.max(0, text.toUpperCase().indexOf(needle.toUpperCase())) })
  for (const f of functions) if (!isFunction(f) && !externalSet.has(f)) return fail('unknown_function', `unknown function ${f}`, f)
  if (ctx.sheets) {
    const known = new Set(ctx.sheets.map(sheetKey))
    for (const s of sheets) if (!known.has(sheetKey(s))) return fail('unknown_sheet', `unknown sheet ${s}`, s)
  }
  if (ctx.known) {
    for (const [n, sheet] of uses) if (!ctx.known(n, sheet)) return fail('unknown_name', `unknown name ${n}`, n)
  } else if (ctx.names) {
    const known = new Set(ctx.names.map((n) => n.toUpperCase()))
    for (const n of names) if (!known.has(n.toUpperCase())) return fail('unknown_name', `unknown name ${n}`, n)
  }
  const calls = checkCalls(ast, text, (f) => externalSet.has(f) && !FUNCTIONS[f])
  if (calls) return { ok: false, code: calls.code, message: calls.message, at: calls.at }
  return { ok: true, ast, canonical: canonical(ast), functions: [...functions].sort(), sheets: [...sheets], names: [...names], externals: [...new Set(externals)], tables: bound.tables }
}

const NAME_RE = /^[A-Za-z_][A-Za-z0-9_.]{0,254}$/
const LOOKS_LIKE_CELL = /^[A-Za-z]{1,3}[0-9]+$|^[Rr][0-9]*[Cc][0-9]*$/

/** A valid range name: Latin letters, not a cell address, not R1C1, not TRUE/FALSE. */
export function validateName(name: string): NameCheck {
  if (!NAME_RE.test(name)) return { ok: false, message: 'a name must start with a latin letter or _ and contain only latin letters, digits, _ and .' }
  if (LOOKS_LIKE_CELL.test(name)) return { ok: false, message: 'a name must not look like a cell address' }
  if (/^(TRUE|FALSE)$/i.test(name)) return { ok: false, message: 'TRUE and FALSE are reserved' }
  return { ok: true }
}
