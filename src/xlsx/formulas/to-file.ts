/**
 * A formula of the book in a file's spelling, printed from its syntax tree: functions newer than Excel 2007 with
 * their prefix (`_xlfn.XLOOKUP`, `_xlfn._xlws.FILTER`), LET as `_xlfn.LET` with its names as `_xlpm.x`, a spill
 * reference `E2#` as `_xlfn.ANCHORARRAY(E2)`, `@x` as `_xlfn.SINGLE(x)`, sheets always quoted, numbers with an upper-case exponent. The rest
 * as the engine's canonical text prints it.
 */
import type { Ast } from '../../syntax/ast.js'
import { ERROR_TEXT } from '../../syntax/ast.js'
import { quoteSheet } from '../../syntax/print.js'
import { XLFN, XLWS } from './prefixes.js'

const fn = (name: string) => (XLWS.has(name) ? `_xlfn._xlws.${name}` : XLFN.has(name) ? `_xlfn.${name}` : name)

const num = (v: number) => String(v).toUpperCase()

export function toFile(ast: Ast): string {
  const go = (n: Ast): string => {
    switch (n.k) {
      case 'num':
        return num(n.v)
      case 'str':
        return `"${n.v.replaceAll('"', '""')}"`
      case 'bool':
        return n.v ? 'TRUE' : 'FALSE'
      case 'err': {
        const text = (n.sheet === undefined ? '' : `${quoteSheet(n.sheet)}!`) + ERROR_TEXT[n.code]
        // a lost spill reference (`#REF!#`) is the anchor of a lost cell
        return n.ref?.endsWith('#') ? `_xlfn.ANCHORARRAY(${text}${n.ref.slice(0, -1)})` : text + (n.ref ?? '')
      }
      case 'ref': {
        const text = (n.sheet === null ? '' : `${quoteSheet(n.sheet)}!`) + n.ref.text
        return n.ref.spill ? `_xlfn.ANCHORARRAY(${text})` : text
      }
      case 'ref3d':
        return `${quoteSheet(`${n.first}:${n.last}`)}!${n.ref.text}`
      case 'name':
        return (n.sheet === undefined ? '' : `${quoteSheet(n.sheet)}!`) + n.name
      case 'at':
        return `_xlfn.SINGLE(${go(n.a)})`
      case 'union':
        return n.items.map(go).join(',')
      case 'fn':
        return `${fn(n.name)}(${n.args.map(go).join(',')})`
      case 'un':
        return n.op + go(n.a)
      case 'pct':
        return `${go(n.a)}%`
      case 'bin':
        return go(n.a) + n.op + go(n.b)
      case 'paren':
        return `(${go(n.a)})`
      case 'arr':
        return `{${n.rows.map((r) => r.map(go).join(',')).join(';')}}`
      case 'missing':
        return ''
      case 'let':
        return `_xlfn.LET(${[...n.binds.flatMap((b) => [`_xlpm.${b.name}`, go(b.value)]), ...(n.body ? [go(n.body)] : [])].join(',')})`
      case 'local':
        return `_xlpm.${n.name}`
    }
  }
  return go(ast)
}
