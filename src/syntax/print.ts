/** Traversal of an AST and its text: canonical, or with another engine's literals. */
import { ERROR_TEXT, type Ast } from './ast.js'
import { specText } from './table.js'

export function walk(ast: Ast, visit: (n: Ast) => void): void {
  visit(ast)
  switch (ast.k) {
    case 'fn':
      for (const a of ast.args) walk(a, visit)
      break
    case 'un':
    case 'pct':
    case 'paren':
    case 'at':
      walk(ast.a, visit)
      break
    case 'union':
      for (const a of ast.items) walk(a, visit)
      break
    case 'bin':
      walk(ast.a, visit)
      walk(ast.b, visit)
      break
    case 'let':
      for (const b of ast.binds) walk(b.value, visit)
      if (ast.body) walk(ast.body, visit)
      break
  }
}

/** The AST with nodes replaced by `f` (bottom-up; undefined keeps a node); unchanged branches are shared. */
export function mapAst(ast: Ast, f: (n: Ast) => Ast | undefined): Ast {
  const go = (n: Ast): Ast => {
    let m: Ast = n
    const list = (xs: Ast[]) => {
      const ys = xs.map(go)
      return ys.some((y, i) => y !== xs[i]) ? ys : xs
    }
    switch (n.k) {
      case 'fn': {
        const args = list(n.args)
        if (args !== n.args) m = { ...n, args }
        break
      }
      case 'un':
      case 'pct':
      case 'paren':
      case 'at': {
        const a = go(n.a)
        if (a !== n.a) m = { ...n, a }
        break
      }
      case 'union': {
        const items = list(n.items)
        if (items !== n.items) m = { ...n, items }
        break
      }
      case 'bin': {
        const [a, b] = [go(n.a), go(n.b)]
        if (a !== n.a || b !== n.b) m = { ...n, a, b }
        break
      }
      case 'let': {
        const binds = n.binds.map((x) => {
          const value = go(x.value)
          return value === x.value ? x : { ...x, value }
        })
        const body = n.body && go(n.body)
        if (body !== n.body || binds.some((x, i) => x !== n.binds[i])) m = { ...n, binds, ...(body ? { body } : {}) }
        break
      }
    }
    return f(m) ?? m
  }
  return go(ast)
}

/** Sheet names compare case-insensitively, as in Excel. */
export const sheetKey = (s: string) => s.toLocaleUpperCase('ru-RU')

export const quoteSheet = (s: string) => `'${s.replaceAll("'", "''")}'`

export interface PrintOptions {
  str?: (s: string) => string
  bool?: (b: boolean) => string
  /** The sheet name as spelled in the book. */
  sheetOf?: (s: string) => string
  /** A defined name as spelled where it is defined, seen from its sheet (undefined: the formula's own). */
  nameOf?: (name: string, sheet: string | undefined) => string
}

/** Prints an AST; canonical by default. Adapters for other engines substitute their own literals. */
export function printFormula(ast: Ast, opts: PrintOptions = {}): string {
  const str = opts.str ?? ((s: string) => `"${s.replaceAll('"', '""')}"`)
  const bool = opts.bool ?? ((b: boolean) => (b ? 'TRUE' : 'FALSE'))
  const sheetOf = opts.sheetOf ?? ((s: string) => s)
  const nameOf = opts.nameOf ?? ((name: string) => name)
  const go = (n: Ast): string => {
    switch (n.k) {
      case 'num':
        return String(n.v)
      case 'str':
        return str(n.v)
      case 'bool':
        return bool(n.v)
      case 'err':
        return (n.sheet === undefined ? '' : `${quoteSheet(sheetOf(n.sheet))}!`) + ERROR_TEXT[n.code] + (n.ref ?? '')
      case 'ref':
        return (n.sheet === null ? '' : `${quoteSheet(sheetOf(n.sheet))}!`) + n.ref.text + (n.ref.spill ? '#' : '')
      case 'ref3d':
        return `${quoteSheet(`${sheetOf(n.first)}:${sheetOf(n.last)}`)}!${n.ref.text}`
      case 'table':
        return `${n.table ?? ''}[${specText(n.spec)}]`
      case 'name':
        return (n.sheet === undefined ? '' : `${quoteSheet(sheetOf(n.sheet))}!`) + nameOf(n.name, n.sheet)
      case 'at':
        return `@${go(n.a)}`
      case 'union':
        return n.items.map(go).join(',')
      case 'fn':
        return `${n.name}(${n.args.map(go).join(',')})`
      case 'un':
        return n.op + go(n.a)
      case 'pct':
        return go(n.a) + '%'
      case 'bin':
        return go(n.a) + n.op + go(n.b)
      case 'paren':
        return `(${go(n.a)})`
      case 'arr':
        return `{${n.rows.map((r) => r.map(go).join(',')).join(';')}}`
      case 'missing':
        return ''
      case 'let':
        return `LET(${[...n.binds.flatMap((b) => [b.name, go(b.value)]), ...(n.body ? [go(n.body)] : [])].join(',')})`
      case 'local':
        return n.name
    }
  }
  return go(ast)
}

/** Canonical stored text (without `=`). */
export const canonical = (ast: Ast, sheetOf?: (s: string) => string) => printFormula(ast, { sheetOf })
