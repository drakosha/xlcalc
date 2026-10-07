/** Traversal of an AST and its text: canonical, or with another engine's literals. */
import { ERROR_TEXT, type Ast } from './ast.js'

export function walk(ast: Ast, visit: (n: Ast) => void): void {
  visit(ast)
  switch (ast.k) {
    case 'fn':
      for (const a of ast.args) walk(a, visit)
      break
    case 'un':
    case 'pct':
    case 'paren':
      walk(ast.a, visit)
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

/** Sheet names compare case-insensitively, as in Excel. */
export const sheetKey = (s: string) => s.toLocaleUpperCase('ru-RU')

export const quoteSheet = (s: string) => `'${s.replaceAll("'", "''")}'`

export interface PrintOptions {
  str?: (s: string) => string
  bool?: (b: boolean) => string
  /** The sheet name as spelled in the book. */
  sheetOf?: (s: string) => string
}

/** Prints an AST; canonical by default. Adapters for other engines substitute their own literals. */
export function printFormula(ast: Ast, opts: PrintOptions = {}): string {
  const str = opts.str ?? ((s: string) => `"${s.replaceAll('"', '""')}"`)
  const bool = opts.bool ?? ((b: boolean) => (b ? 'TRUE' : 'FALSE'))
  const sheetOf = opts.sheetOf ?? ((s: string) => s)
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
      case 'name':
        return n.name
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
