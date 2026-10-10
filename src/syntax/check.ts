/** Checks of a formula with codes: functions and their argument counts, sheets, names; valid range names. */
import type { CheckCode, NameCheck, ValidateContext, ValidateResult } from '../contract.js'
import { FUNCTIONS, isFunction } from '../functions.js'
import { english } from '../locale/english.js'
import type { Message, NameRule } from '../locale/messages.js'
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
            if (!lit) return new FormulaError('unsupported', { key: 'externalLiterals', fn: n.name }, at(n.name))
          }
          return null
        }
        const spec = FUNCTIONS[n.name]
        if (!spec) return null
        const k = n.args.length
        if (!spec.takes(k)) return new FormulaError('syntax', { key: 'argumentCount', fn: n.name, counts: spec.counts, got: k }, at(n.name))
        const left = n.args.findIndex((a, i) => a.k === 'missing' && !spec.omits(i))
        if (left >= 0) return new FormulaError('syntax', { key: 'argumentLeftOut', fn: n.name, position: left + 1 }, at(n.name))
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

/**
 * What the engine adds to a check: a name's formula (`union`), which names a book has as a sheet sees them (`known`), the
 * table of the formula's cell (`enclosing`), the language of messages (`say`, English without it).
 */
interface BookCheck {
  union?: boolean
  known?: (name: string, sheet: string | undefined) => boolean
  enclosing?: string
  say?: (m: Message) => string
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
  const unknown = (m: Message, needle: string): Fail => ({ ok: false, code: 'unknown_name', message: (ctx.say ?? english)(m), at: at(needle) })
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
      if (ctx.tables) fail = unknown({ key: 'tableRefOutside' }, '[')
      return undefined
    }
    const t = known?.get(sheetKey(name))
    if (!t) {
      if (known) fail = unknown({ key: 'unknownTable', table: name }, name)
      else used.add(name)
      return undefined
    }
    used.add(t.name)
    const cols = n.spec.cols?.map((c) => t.columns.find((x) => sheetKey(x) === sheetKey(c)))
    const missing = n.spec.cols?.find((c, i) => cols![i] === undefined)
    if (missing !== undefined) {
      fail = unknown({ key: 'unknownColumn', table: t.name, column: missing }, missing)
      return undefined
    }
    return { k: 'table', table: t.name, spec: cols ? { ...n.spec, cols: cols as [string] | [string, string] } : n.spec }
  })
  return fail ?? { ok: true, ast: out, tables: [...used] }
}

/** Formula check: parsing, whitelisted or external functions, sheets, names, argument counts. */
export function validate(text: string, ctx: ValidateContext & BookCheck = {}): ({ ok: true; ast: Ast } & Extract<ValidateResult, { ok: true }>) | Extract<ValidateResult, { ok: false }> {
  const say = ctx.say ?? english
  let ast: Ast
  try {
    ast = parse(text, { union: ctx.union })
  } catch (e) {
    if (e instanceof FormulaError) return { ok: false, code: e.code, message: say(e.msg), at: e.at }
    throw e
  }
  // nothing to bind without tables, but a reference without a table name, which fails
  const bound = ctx.tables?.length || text.includes('[') ? bindTables(ast, text, ctx) : { ok: true as const, ast, tables: [] }
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
  const fail = (code: CheckCode, m: Message, needle: string) => ({ ok: false as const, code, message: say(m), at: Math.max(0, text.toUpperCase().indexOf(needle.toUpperCase())) })
  for (const f of functions) if (!isFunction(f) && !externalSet.has(f)) return fail('unknown_function', { key: 'unknownFunction', fn: f }, f)
  if (ctx.sheets) {
    const known = new Set(ctx.sheets.map(sheetKey))
    for (const s of sheets) if (!known.has(sheetKey(s))) return fail('unknown_sheet', { key: 'unknownSheet', sheet: s }, s)
  }
  if (ctx.known) {
    for (const [n, sheet] of uses) if (!ctx.known(n, sheet)) return fail('unknown_name', { key: 'unknownName', name: n }, n)
  } else if (ctx.names) {
    const known = new Set(ctx.names.map((n) => n.toUpperCase()))
    for (const n of names) if (!known.has(n.toUpperCase())) return fail('unknown_name', { key: 'unknownName', name: n }, n)
  }
  const calls = checkCalls(ast, text, (f) => externalSet.has(f) && !FUNCTIONS[f])
  if (calls) return { ok: false, code: calls.code, message: say(calls.msg), at: calls.at }
  return { ok: true, ast, canonical: canonical(ast), functions: [...functions].sort(), sheets: [...sheets], names: [...names], externals: [...new Set(externals)], tables: bound.tables }
}

const NAME = /^[\p{L}_\\][\p{L}\p{N}_.\\?]{0,254}$/u

/**
 * Why Excel would not take a word that reads as a reference or a boolean as a name (null when it would); `booleans`
 * are the book's locale's words for TRUE and FALSE (ИСТИНА and ЛОЖЬ in ru-RU, which Excel ru-RU refuses).
 */
export function reservedName(name: string, booleans: readonly string[] = []): NameRule | null {
  // a cell address in Latin letters only (`АБ12` is a name), R1C1 with `R` and `C` alone too
  if (/^[A-Za-z]{1,3}[0-9]+$/.test(name) || /^[RrCc]$|^[Rr][0-9]*[Cc][0-9]*$/.test(name)) return 'cell'
  if (/^(TRUE|FALSE)$/i.test(name)) return 'true'
  const up = name.toUpperCase()
  if (booleans.some((b) => b.toUpperCase() === up)) return 'boolean'
  return null
}

/** The rule of names a name breaks (null when Excel would take it), a defined name's or a table's, the namespace aside. */
export function nameRule(name: string, booleans: readonly string[] = []): NameRule | null {
  return NAME.test(name) ? reservedName(name, booleans) : 'chars'
}

/** Why Excel would not take a name, in English (null when it would): `nameRule`'s message. */
export function nameProblem(name: string, booleans: readonly string[] = []): string | null {
  const rule = nameRule(name, booleans)
  return rule && english({ key: 'nameRule', name, rule })
}

/** A valid defined name by Excel's rules (`nameProblem`) in a book whose locale's booleans are `booleans`. */
export function validateName(name: string, booleans: readonly string[] = []): NameCheck {
  const problem = nameProblem(name, booleans)
  return problem === null ? { ok: true } : { ok: false, message: problem }
}
