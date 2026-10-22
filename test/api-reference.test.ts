/**
 * `docs/api.md` is the reference `test/api-reference.ts` makes from the package's types: a change of an export or of its
 * comment shows there (`npx tsx test/api-reference.ts --write`), and every function, value and type has a comment.
 */
import { join, resolve } from 'node:path'
import ts from 'typescript'
import { describe, expect, it } from 'vitest'
import { apiReference, current, ENTRIES } from './api-reference.js'

describe('docs/api.md', () => {
  const made = apiReference()

  it('is the reference the package types make', () => {
    expect(current() === made, 'docs/api.md is out of date: npx tsx test/api-reference.ts --write').toBe(true)
  }, 60_000)

  it('has a comment for every export', () => {
    const bare = made
      .split(/\n(?=### )/)
      .filter((b) => b.startsWith('### '))
      .filter((b) => !b.slice(b.indexOf('\n')).replace(/```[\s\S]*?```/g, '').trim())
      .map((b) => b.slice(0, b.indexOf('\n')))
    expect(bare).toEqual([])
  })
})

// xlcalc/xlsx keeps its own types (`ReadXlsxOptions` and the rest) outside the contract on purpose
describe('xlcalc/contract', () => {
  const root = resolve(import.meta.dirname, '..')
  const files = ENTRIES.filter(([entry]) => entry !== 'xlcalc/xlsx').map(([, f]) => join(root, f))
  const program = ts.createProgram(files, { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.NodeNext, moduleResolution: ts.ModuleResolutionKind.NodeNext, strict: true, skipLibCheck: true })
  const checker = program.getTypeChecker()
  const resolved = (s: ts.Symbol) => (s.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(s) : s)
  const exportsOf = (f: string) => checker.getExportsOfModule(checker.getSymbolAtLocation(program.getSourceFile(f)!)!).map(resolved)
  const contract = new Set(exportsOf(join(root, 'src/contract.ts')))
  const typeOf = (name: string) => checker.getDeclaredTypeOfSymbol([...contract].find((s) => s.name === name)!)

  it("names every type the root exports and every type of an entry's function signatures", () => {
    const isOwnType = (s: ts.Symbol) => !!(s.flags & (ts.SymbolFlags.Interface | ts.SymbolFlags.TypeAlias)) && !!s.declarations?.[0].getSourceFile().fileName.startsWith(join(root, 'src'))
    const missing = new Set<string>()
    for (const f of files) {
      for (const s of exportsOf(f)) {
        if (!(s.flags & ts.SymbolFlags.Value)) {
          if (f === join(root, 'src/index.ts') && !contract.has(s)) missing.add(s.name)
          continue
        }
        for (const sig of checker.getTypeOfSymbol(s).getCallSignatures()) {
          const d = sig.getDeclaration()
          const visit = (n: ts.Node): void => {
            const t = ts.isTypeReferenceNode(n) ? checker.getSymbolAtLocation(n.typeName) : undefined
            if (t && isOwnType(resolved(t)) && !contract.has(resolved(t))) missing.add(`${s.name}: ${resolved(t).name}`)
            ts.forEachChild(n, visit)
          }
          for (const node of d ? [...d.parameters.map((p) => p.type), d.type] : []) if (node) visit(node)
        }
      }
    }
    expect([...missing]).toEqual([])
  }, 60_000)

  it('types ReadCell with the fields readCells and findCells return: the address and each CellReadField', () => {
    const fields = (typeOf('CellReadField') as ts.UnionType).types.map((t) => (t as ts.StringLiteralType).value)
    const declared = typeOf('ReadCell').getProperties().filter((p) => !p.declarations?.some((d) => ts.getJSDocDeprecatedTag(d)))
    expect(declared.map((p) => p.name).sort()).toEqual(['cell', 'sheet', ...fields].sort())
  })
})
