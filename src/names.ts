/** Defined names prepared for the engine: checked, their scope resolved, their reference parsed; problems go to name issues. */
import type { NameInput, NameIssue } from './contract.js'
import { FormulaError, parse, sheetKey, validateName, type Ast } from './syntax.js'

export interface PreparedName {
  name: string
  /** Index of the scope sheet; -1 for a global name. */
  scope: number
  /** The reference (a spill reference too), or `#REF!` when Excel lost it (`'S'!#REF!`, `#REF!A1`). */
  ast: Extract<Ast, { k: 'ref' | 'err' }>
}

/** Names in the book's order; an invalid one is left out with its issue. */
export function prepareNames(sheetIdx: Map<string, number>, names: NameInput[]): { names: PreparedName[]; issues: NameIssue[] } {
  const out: PreparedName[] = []
  const issues: NameIssue[] = []
  const seen = new Set<string>()
  for (const n of names) {
    const issue = (reason: string) => issues.push({ name: n.name, ...(n.scope ? { scope: n.scope } : {}), reason })
    const ok = validateName(n.name)
    if (!ok.ok) {
      issue(ok.message)
      continue
    }
    const scope = n.scope === undefined ? -1 : (sheetIdx.get(sheetKey(n.scope)) ?? -2)
    if (scope === -2) {
      issue(`unknown scope sheet ${n.scope}`)
      continue
    }
    const dup = `${scope}:${n.name.toUpperCase()}`
    if (seen.has(dup)) {
      issue('duplicate name')
      continue
    }
    seen.add(dup)
    let ast: Ast
    try {
      ast = parse(n.ref)
    } catch (e) {
      if (!(e instanceof FormulaError)) throw e
      issue(`bad reference: ${e.message}`)
      continue
    }
    // a reference Excel lost stays a name, as in Excel: formulas using it give #REF!
    if (ast.k === 'err' && ast.code === 'REF') {
      out.push({ name: n.name, scope, ast })
      continue
    }
    if (ast.k !== 'ref' || ast.sheet === null) {
      issue('a name must reference a cell or a range with a sheet')
      continue
    }
    if (!sheetIdx.has(sheetKey(ast.sheet))) {
      issue(`unknown sheet ${ast.sheet}`)
      continue
    }
    out.push({ name: n.name, scope, ast })
  }
  return { names: out, issues }
}

/** The name visible from sheet `here`: its own local name, else the global one. */
export function nameFor(names: PreparedName[], here: number, name: string): PreparedName | undefined {
  const up = name.toUpperCase()
  return names.find((n) => n.scope === here && n.name.toUpperCase() === up) ?? names.find((n) => n.scope === -1 && n.name.toUpperCase() === up)
}
