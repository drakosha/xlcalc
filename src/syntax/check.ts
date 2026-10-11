/** Checks of a formula with codes: functions and their argument counts, sheets, names; valid range names. */
import type { CheckCode, NameCheck, NameRefusal, ValidateContext, ValidateResult } from '../contract.js'
import { caseKey, wordKey } from '../collate.js'
import { FUNCTIONS, isFunction } from '../functions.js'
import { english } from '../locale/english.js'
import type { Message, NameRule } from '../locale/messages.js'
import { FormulaError, type Ast } from './ast.js'
import { boolWord, parse } from './parse.js'
import { IDENT_FIRST, IDENT_NEXT } from './scan.js'
import { readsAsCell } from './tokens.js'
import { canonical, mapAst, sheetKey, walk } from './print.js'

/** Where a message points: the first place of `needle` in `text` in any case, the start when it is not there. */
const positionOf = (text: string, needle: string) => Math.max(0, wordKey(text).indexOf(wordKey(needle)))

/** Argument counts of each function and the arguments it cannot leave out; literal arguments of external functions. */
export function checkCalls(ast: Ast, src: string, external: (fn: string) => boolean): FormulaError | null {
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
            if (!lit) return new FormulaError('unsupported', { key: 'externalLiterals', fn: n.name }, positionOf(src, n.name))
          }
          return null
        }
        const spec = FUNCTIONS[n.name]
        if (!spec) return null
        const k = n.args.length
        if (!spec.takes(k)) return new FormulaError('syntax', { key: 'argumentCount', fn: n.name, counts: spec.counts, got: k }, positionOf(src, n.name))
        const left = n.args.findIndex((a, i) => a.k === 'missing' && !spec.omits(i))
        if (left >= 0) return new FormulaError('syntax', { key: 'argumentLeftOut', fn: n.name, position: left + 1 }, positionOf(src, n.name))
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
 * language of messages (`say`, English without it).
 */
interface BookCheck {
  union?: boolean
  known?: (name: string, sheet: string | undefined) => boolean
  say?: (m: Message) => string
}

type Fail = Extract<ValidateResult, { ok: false }>

/**
 * Structured references bound to the tables of `ctx.tables`: a reference without a table name takes the formula's
 * table (`ctx.table`), tables and columns are spelled as named, a bare table name is the table's data (`Sales[]`, as a
 * file stores it). An unknown table or column, or a reference without a table name outside one, fails with `unknown_name`.
 */
export function bindTables(ast: Ast, text: string, ctx: ValidateContext & BookCheck): { ok: true; ast: Ast; tables: string[] } | Fail {
  const known = ctx.tables ? new Map(ctx.tables.map((t) => [sheetKey(t.name), t])) : null
  const used = new Set<string>()
  let fail: Fail | null = null
  const unknown = (m: Message, needle: string): Fail => ({ ok: false, code: 'unknown_name', message: (ctx.say ?? english)(m), at: positionOf(text, needle) })
  const out = mapAst(ast, (n) => {
    if (fail) return undefined
    if (n.k === 'name' && n.sheet === undefined && known?.has(sheetKey(n.name))) {
      const t = known.get(sheetKey(n.name))!
      used.add(t.name)
      return { k: 'table', table: t.name, spec: { areas: [] } }
    }
    if (n.k !== 'table') return undefined
    const name = n.table ?? ctx.table
    if (name === undefined) {
      if (ctx.tables) fail = unknown({ key: 'tableRefOutside' }, '[')
      return undefined
    }
    const t = known?.get(sheetKey(name))
    if (!t) {
      if (known) {
        fail = unknown({ key: 'unknownTable', table: name }, name)
        return undefined
      }
      used.add(name)
      // without the book's tables, a reference in the formula's table takes its name as given
      return n.table === null ? { ...n, table: name } : undefined
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
  const externalSet = new Set((ctx.externalFunctions ?? []).map(wordKey))
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
  const fail = (code: CheckCode, m: Message, needle: string) => ({ ok: false as const, code, message: say(m), at: positionOf(text, needle) })
  for (const f of functions) if (!isFunction(f) && !externalSet.has(f)) return fail('unknown_function', { key: 'unknownFunction', fn: f }, f)
  if (ctx.sheets) {
    const known = new Set(ctx.sheets.map(sheetKey))
    for (const s of sheets) if (!known.has(sheetKey(s))) return fail('unknown_sheet', { key: 'unknownSheet', sheet: s }, s)
  }
  if (ctx.known) {
    for (const [n, sheet] of uses) if (!ctx.known(n, sheet)) return fail('unknown_name', { key: 'unknownName', name: n }, n)
  } else if (ctx.names) {
    const known = new Set(ctx.names.map(caseKey))
    for (const n of names) if (!known.has(caseKey(n))) return fail('unknown_name', { key: 'unknownName', name: n }, n)
  }
  const calls = checkCalls(ast, text, (f) => externalSet.has(f) && !FUNCTIONS[f])
  if (calls) return { ok: false, code: calls.code, message: say(calls.msg), at: calls.at }
  return { ok: true, ast, canonical: canonical(ast), functions: [...functions].sort(), sheets: [...sheets], names: [...names], externals: [...new Set(externals)], tables: bound.tables }
}

/** A name's characters (an identifier's) and its length, 255 at most. */
const NAME = new RegExp(`^${IDENT_FIRST}${IDENT_NEXT}{0,254}$`, 'u')

/**
 * Why Excel would not take a word that reads as a reference or a boolean as a name (null when it would); `booleans`
 * are the book's locale's words for TRUE and FALSE (ИСТИНА and ЛОЖЬ in ru-RU, which Excel ru-RU refuses).
 */
export function reservedName(name: string, booleans: readonly string[] = []): NameRule | null {
  // a cell of the sheet in Latin letters only (`АБ12`, `XFE1` and `A0` are names: the set `tables/paste-names-held`), R1C1
  // with `R` and `C` alone too
  if (readsAsCell(name) || /^[RrCc]$|^[Rr][0-9]*[Cc][0-9]*$/.test(name)) return 'cell'
  if (boolWord(name) !== undefined) return 'true'
  const up = wordKey(name)
  if (booleans.some((b) => wordKey(b) === up)) return 'boolean'
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

/** The code of `validateName`'s refusal for each rule a name breaks: TRUE and FALSE are booleans in every locale. */
const NAME_REFUSAL: Record<NameRule, NameRefusal> = { chars: 'chars', cell: 'reference', true: 'boolean', boolean: 'boolean' }

/** A valid defined name by Excel's rules (`nameRule`) in a book whose locale's booleans are `booleans`. */
export function validateName(name: string, booleans: readonly string[] = []): NameCheck {
  const rule = nameRule(name, booleans)
  return rule === null ? { ok: true } : { ok: false, code: NAME_REFUSAL[rule], message: english({ key: 'nameRule', name, rule }) }
}
