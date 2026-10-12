/**
 * Defined names prepared for the engine: checked, their scope resolved, their formula parsed (a reference, a
 * constant, a formula, a union); problems go to name issues. A name's formula is computed where it is used
 * (`exec/walk.ts`).
 */
import type { ExternalBook, NameInput, NameIssue } from './contract.js'
import { FormulaError, nameProblem, parse, sheetKey, validate, walk, type Ast } from './syntax.js'
import { caseKey } from './collate.js'

export interface PreparedName {
  name: string
  /** Index of the scope sheet; -1 for a global name. Its formula sees the global names, as a file's formula of a name does. */
  scope: number
  /** The formula, or the error it gives where Excel would not compute it (`#REF!` for a lost reference). */
  ast: Ast
}

/** Names in the book's order; an invalid one is left out with its issue. `booleans`: the book's locale's words for TRUE and FALSE, no names. */
export function prepareNames(sheetIdx: Map<string, number>, names: NameInput[], externalFunctions: string[], tables: { name: string; columns: string[] }[] = [], booleans: readonly string[] = [], externalBooks: ExternalBook[] = []): { names: PreparedName[]; issues: NameIssue[] } {
  const out: PreparedName[] = []
  const issues: NameIssue[] = []
  const seen = new Set<string>()
  const sheets = [...sheetIdx.keys()]
  for (const n of names) {
    const issue = (reason: string, dropped = true) => issues.push({ name: n.name, ...(n.scope ? { scope: n.scope } : {}), reason, dropped })
    const problem = nameProblem(n.name, booleans)
    if (problem) {
      issue(problem)
      continue
    }
    const scope = n.scope === undefined ? -1 : (sheetIdx.get(sheetKey(n.scope)) ?? -2)
    if (scope === -2) {
      issue(`unknown scope sheet ${n.scope}`)
      continue
    }
    const dup = `${scope}:${caseKey(n.name)}`
    if (seen.has(dup)) {
      issue('duplicate name')
      continue
    }
    seen.add(dup)
    let ast: Ast
    try {
      ast = parse(n.ref, { union: true })
    } catch (e) {
      if (!(e instanceof FormulaError)) throw e
      issue(`bad formula: ${e.message}`)
      continue
    }
    if (withoutSheet(ast)) {
      issue('a reference in a name needs its sheet')
      continue
    }
    // a formula Excel would not compute stays a name giving its error, as a lost reference does (#REF!)
    const v = validate(n.ref, { sheets, externalFunctions, union: true, tables, externalBooks })
    if (!v.ok) {
      issue(v.message, false)
      ast = { k: 'err', code: v.code === 'unknown_sheet' || v.code === 'external_ref' ? 'REF' : v.code === 'unknown_function' || v.code === 'unknown_name' ? 'NAME' : 'ERROR' }
    } else ast = v.ast
    out.push({ name: n.name, scope, ast })
  }
  return { names: out, issues }
}

/** A reference without its sheet, which a name's formula cannot have. */
function withoutSheet(ast: Ast): boolean {
  let found = false
  walk(ast, (n) => (found ||= n.k === 'ref' && n.sheet === null))
  return found
}

/** The name the names of sheet `here` show: its own local name, else the global one (`here` -1: the global one). */
export function nameFor(names: PreparedName[], here: number, name: string): PreparedName | undefined {
  const up = caseKey(name)
  return names.find((n) => n.scope === here && caseKey(n.name) === up) ?? names.find((n) => n.scope === -1 && caseKey(n.name) === up)
}

/** `nameFor` among a book's names as given: the one `sheet` sees, its own local name, else the global one (no sheet: the global one). */
export function inputNameFor(names: readonly NameInput[], sheet: string | undefined, name: string): NameInput | undefined {
  const up = caseKey(name)
  const here = sheet === undefined ? undefined : sheetKey(sheet)
  return (here === undefined ? undefined : names.find((n) => n.scope !== undefined && sheetKey(n.scope) === here && caseKey(n.name) === up)) ?? names.find((n) => n.scope === undefined && caseKey(n.name) === up)
}
