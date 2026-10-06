/** Checks of a formula with codes: functions and their argument counts, sheets, names; valid range names. */
import type { CheckCode, NameCheck, ValidateContext, ValidateResult } from '../contract.js'
import { FUNCTIONS } from '../functions.js'
import { FormulaError, type Ast } from './ast.js'
import { parse } from './parse.js'
import { canonical, sheetKey, walk } from './print.js'

/** Argument counts of each function; literal arguments of external functions. */
export function checkCalls(ast: Ast, src: string, external: (fn: string) => boolean): FormulaError | null {
  const at = (needle: string) => Math.max(0, src.toUpperCase().indexOf(needle.toUpperCase()))
  const visit = (n: Ast): FormulaError | null => {
    switch (n.k) {
      case 'paren':
      case 'un':
      case 'pct':
        return visit(n.a)
      case 'bin':
        return visit(n.a) ?? visit(n.b)
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

/** Formula check: parsing, whitelisted or external functions, sheets, names, argument counts. */
export function validate(text: string, ctx: ValidateContext = {}): ({ ok: true; ast: Ast } & Extract<ValidateResult, { ok: true }>) | Extract<ValidateResult, { ok: false }> {
  let ast: Ast
  try {
    ast = parse(text)
  } catch (e) {
    if (e instanceof FormulaError) return { ok: false, code: e.code, message: e.message, at: e.at }
    throw e
  }
  const externalSet = new Set((ctx.externalFunctions ?? []).map((f) => f.toUpperCase()))
  const functions = new Set<string>()
  const sheets = new Set<string>()
  const names = new Set<string>()
  const externals: string[] = []
  walk(ast, (n) => {
    if (n.k === 'fn') {
      functions.add(n.name)
      if (externalSet.has(n.name) && !FUNCTIONS[n.name]) externals.push(canonical(n))
    } else if ((n.k === 'ref' && n.sheet !== null) || (n.k === 'err' && n.sheet !== undefined)) sheets.add(n.sheet!)
    else if (n.k === 'name') names.add(n.name)
  })
  const fail = (code: CheckCode, message: string, needle: string) => ({ ok: false as const, code, message, at: Math.max(0, text.toUpperCase().indexOf(needle.toUpperCase())) })
  for (const f of functions) if (!FUNCTIONS[f] && !externalSet.has(f)) return fail('unknown_function', `unknown function ${f}`, f)
  if (ctx.sheets) {
    const known = new Set(ctx.sheets.map(sheetKey))
    for (const s of sheets) if (!known.has(sheetKey(s))) return fail('unknown_sheet', `unknown sheet ${s}`, s)
  }
  if (ctx.names) {
    const known = new Set(ctx.names.map((n) => n.toUpperCase()))
    for (const n of names) if (!known.has(n.toUpperCase())) return fail('unknown_name', `unknown name ${n}`, n)
  }
  const calls = checkCalls(ast, text, (f) => externalSet.has(f) && !FUNCTIONS[f])
  if (calls) return { ok: false, code: calls.code, message: calls.message, at: calls.at }
  return { ok: true, ast, canonical: canonical(ast), functions: [...functions].sort(), sheets: [...sheets], names: [...names], externals: [...new Set(externals)] }
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
