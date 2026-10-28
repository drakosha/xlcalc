/**
 * `docs/api.md`, the API reference, made from the package's types and doc comments by the TypeScript compiler: each entry
 * of the package (`package.json` `exports`, the unstable ones left out) with its functions and values, their signatures
 * and comments, then every type an entry exports, as its source declares it. `test/api-reference.test.ts` checks that
 * the file is the one this makes; `npx tsx test/api-reference.ts --write` writes it again.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import ts from 'typescript'

const root = resolve(import.meta.dirname, '..')

/** The package's entries and their sources, in the order the reference lists them. */
export const ENTRIES: readonly [string, string][] = [
  ['xlcalc', 'src/index.ts'],
  ['xlcalc/a1', 'src/a1.ts'],
  ['xlcalc/text', 'src/text.ts'],
  ['xlcalc/entry', 'src/typed-entry.ts'],
  ['xlcalc/xlsx', 'src/xlsx.ts'],
  ['xlcalc/xlsx/keep', 'src/xlsx-keep.ts'],
  ['xlcalc/worker', 'src/worker/index.ts'],
  ['xlcalc/function-text', 'src/function-text/index.ts'],
  ['xlcalc/function-text/ru-RU', 'src/function-text/ru-RU.ts'],
  ['xlcalc/locales', 'src/locales.ts'],
  ['xlcalc/locales/ru-RU', 'src/locales/ru-RU.ts'],
  ['xlcalc/contract', 'src/contract.ts'],
]

const INTRO = `# API reference

Every function, value and type the package exports, entry by entry, with the comments of its source. Guides with
examples are the [README](../README.md) and the other files of \`docs/\`. This file is made from the package's types
(\`npx tsx test/api-reference.ts --write\`): change the comments in the source, not here.
`

/** A symbol's comment: its text and its tags but the internal ones. */
function commentOf(symbol: ts.Symbol, checker: ts.TypeChecker, signature?: ts.Signature): string {
  const parts = symbol.getDocumentationComment(checker)
  const text = ts.displayPartsToString(parts.length || !signature ? parts : signature.getDocumentationComment(checker)).trim()
  const tags = (signature ? signature.getJsDocTags() : symbol.getJsDocTags(checker))
    .filter((t) => t.name !== 'internal')
    .map((t) => `@${t.name}${t.text ? ` ${ts.displayPartsToString(t.text)}` : ''}`)
  return [text, ...tags].filter(Boolean).join('\n\n')
}

/** A type declaration as its source writes it, comments of its members kept, its own comment left out (it is the prose). */
function declarationText(node: ts.Node): string {
  const text = node.getText()
  return text.replace(/^export (declare )?/, '')
}

/** Whether a symbol is a type, an interface or a class rather than a function or a value alone. */
const isType = (s: ts.Symbol) => !!(s.flags & (ts.SymbolFlags.Interface | ts.SymbolFlags.TypeAlias)) && !(s.flags & ts.SymbolFlags.Value)

interface Item {
  name: string
  section: string
  key: string
}

export function apiReference(): string {
  const program = ts.createProgram(
    ENTRIES.map(([, f]) => join(root, f)),
    { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.NodeNext, moduleResolution: ts.ModuleResolutionKind.NodeNext, strict: true, skipLibCheck: true },
  )
  const checker = program.getTypeChecker()
  const out: string[] = [INTRO]
  // where a symbol was described first, so another entry that exports it links there
  const described = new Map<ts.Symbol, string>()
  const types = new Map<ts.Symbol, Item>()
  const toc: string[] = []
  const sections: string[] = []
  for (const [entry, file] of ENTRIES) {
    const source = program.getSourceFile(join(root, file))!
    const module = checker.getSymbolAtLocation(source)!
    const lines: string[] = []
    const seeAlso: string[] = []
    const exports = checker.getExportsOfModule(module).sort((a, b) => a.name.localeCompare(b.name))
    for (const exported of exports) {
      const symbol = exported.flags & ts.SymbolFlags.Alias ? checker.getAliasedSymbol(exported) : exported
      if (isType(symbol)) {
        if (!types.has(symbol)) types.set(symbol, { name: exported.name, section: entry, key: exported.name })
        continue
      }
      const where = described.get(symbol)
      if (where) {
        seeAlso.push(`[\`${exported.name}\`](#${where})`)
        continue
      }
      const anchor = `${slug(entry)}-${exported.name.toLowerCase()}`
      described.set(symbol, anchor)
      lines.push(`### <a id="${anchor}"></a>\`${exported.name}\``, '')
      const decl = symbol.declarations?.[0]
      const type = checker.getTypeOfSymbolAtLocation(symbol, decl ?? source)
      const signatures = type.getCallSignatures()
      if (symbol.flags & ts.SymbolFlags.Class && decl && ts.isClassDeclaration(decl)) {
        lines.push('```ts', `class ${exported.name}${decl.heritageClauses ? ' ' + decl.heritageClauses.map((h) => h.getText()).join(' ') : ''}`, '```', '')
        const text = commentOf(symbol, checker)
        if (text) lines.push(text, '')
        continue
      }
      if (signatures.length) {
        lines.push('```ts', ...signatures.map((s) => `function ${exported.name}${signatureText(checker.signatureToString(s, undefined, ts.TypeFormatFlags.NoTruncation | ts.TypeFormatFlags.WriteArrowStyleSignature))}`), '```', '')
        const text = commentOf(symbol, checker, signatures[0])
        if (text) lines.push(text, '')
      } else {
        lines.push('```ts', `const ${exported.name}: ${checker.typeToString(type, undefined, ts.TypeFormatFlags.NoTruncation)}`, '```', '')
        const text = commentOf(symbol, checker)
        if (text) lines.push(text, '')
      }
    }
    const typesHere = [...types.values()].filter((t) => t.section === entry).length
    if (!lines.length && !seeAlso.length && !typesHere) continue
    toc.push(`- [\`${entry}\`](#${slug(entry)})`)
    const body = [`## ${entry}`, '']
    if (seeAlso.length) body.push(`Described under an entry above: ${seeAlso.join(', ')}.`, '')
    if (typesHere) body.push(`Types: see [Types](#types).`, '')
    sections.push([...body, ...lines].join('\n'))
  }
  toc.push('- [Types](#types)')
  out.push(toc.join('\n'), '', ...sections, '## Types', '')
  for (const [symbol, item] of [...types].sort((a, b) => a[1].name.localeCompare(b[1].name))) {
    out.push(`### \`${item.name}\``, '')
    const text = commentOf(symbol, checker)
    if (text) out.push(text, '')
    const decl = symbol.declarations?.find((d) => ts.isInterfaceDeclaration(d) || ts.isTypeAliasDeclaration(d))
    if (decl) out.push('```ts', declarationText(decl), '```', '')
  }
  const targets = new Map<string, string>([...[...types.values()].map((t): [string, string] => [t.name, slug(`\`${t.name}\``)]), ...[...described].map(([s, a]): [string, string] => [s.name, a])])
  return linked(out.join('\n'), targets).replace(/\n{3,}/g, '\n\n').trimEnd() + '\n'
}

/** A signature as `function name(...)` writes it: its return after `:`, optional parameters without their `| undefined`. */
const signatureText = (s: string) => s.replace(/(\w\?: [^,()]*?) \| undefined(?=[,)])/g, '$1').replace(/\) => /, '): ')

/** The prose's names of types and functions the reference describes as links to them (`BookInput`), outside code and headings. */
function linked(text: string, targets: Map<string, string>): string {
  let code = false
  return text
    .split('\n')
    .map((line) => {
      if (line.startsWith('```')) code = !code
      if (code || line.startsWith('#')) return line
      return line.replace(/(?<!\[)`([A-Za-z_]\w*)((?:\.\w+)*(?:\(\))?)`/g, (m, name: string, rest: string) => {
        const anchor = targets.get(name)
        return anchor ? `[\`${name}${rest}\`](#${anchor})` : m
      })
    })
    .join('\n')
}

/** An entry's anchor: its heading as GitHub writes it. */
const slug = (heading: string) => heading.toLowerCase().replace(/[^\p{L}\p{N}\s_-]/gu, '').replace(/\s/g, '-')

if (process.argv.includes('--write')) writeFileSync(join(root, 'docs/api.md'), apiReference())
else if (import.meta.url === `file://${process.argv[1]}`) process.stdout.write(apiReference())

export const current = () => readFileSync(join(root, 'docs/api.md'), 'utf8')
