/** The parser: recursive descent over the tokens, Excel's operator precedence, function calls, array constants. */
import { MAX_FORMULA_DEPTH, MAX_FORMULA_LENGTH } from '../contract.js'
import { bad, type Ast, type BinOp, type Lit } from './ast.js'
import { tokenize, type Tok } from './tokens.js'

const LEVELS: string[][] = [['=', '<>', '<', '>', '<=', '>='], ['&'], ['+', '-'], ['*', '/'], ['^']]

export function parse(src: string): Ast {
  if (src.trim() === '') bad(0, 'empty formula', 'empty')
  if (src.length > MAX_FORMULA_LENGTH) bad(MAX_FORMULA_LENGTH, `formula longer than ${MAX_FORMULA_LENGTH}`, 'too_long')
  const toks = tokenize(src)
  let p = 0
  let depth = 0
  const peek = () => toks[p]
  const enter = (at: number) => {
    if (++depth > MAX_FORMULA_DEPTH) bad(at, `nesting deeper than ${MAX_FORMULA_DEPTH}`, 'too_deep')
  }

  const binary = (level: number): Ast => {
    if (level === LEVELS.length) return unary()
    let a = binary(level + 1)
    for (;;) {
      const t = peek()
      if (t.t !== 'op' || !LEVELS[level].includes(t.v)) return a
      p++
      const b = binary(level + 1)
      a = { k: 'bin', op: t.v as BinOp, a, b }
    }
  }
  // in Excel unary minus binds tighter than ^: -2^2 = 4
  const unary = (): Ast => {
    const t = peek()
    if (t.t === 'op' && (t.v === '-' || t.v === '+')) {
      p++
      enter(t.at)
      const a = unary()
      depth--
      return { k: 'un', op: t.v, a }
    }
    return postfix()
  }
  const postfix = (): Ast => {
    let a = primary()
    while (peek().t === 'op' && (peek() as { v: string }).v === '%') {
      p++
      a = { k: 'pct', a }
    }
    return a
  }
  const primary = (): Ast => {
    const t = toks[p++]
    switch (t.t) {
      case 'num':
        return { k: 'num', v: t.v }
      case 'str':
        return { k: 'str', v: t.v }
      case 'err':
        return { k: 'err', code: t.code, ...(t.sheet !== undefined ? { sheet: t.sheet } : {}), ...(t.ref !== undefined ? { ref: t.ref } : {}) }
      case 'ref':
        return { k: 'ref', sheet: t.sheet, ref: t.ref }
      case 'ident': {
        const up = t.v.toUpperCase()
        if (up === 'TRUE' || up === 'FALSE') return { k: 'bool', v: up === 'TRUE' }
        return { k: 'name', name: t.v }
      }
      case '(': {
        enter(t.at)
        const a = binary(0)
        depth--
        expect(')')
        return { k: 'paren', a }
      }
      case '{':
        return constant(t.at)
      case 'fn': {
        expect('(')
        enter(t.at)
        const args: Ast[] = []
        if (peek().t !== ')') {
          for (;;) {
            args.push(binary(0))
            if (peek().t === ',') {
              p++
              continue
            }
            break
          }
        }
        depth--
        expect(')')
        return { k: 'fn', name: t.v, args }
      }
      default:
        return bad(t.at, t.t === 'eof' ? 'unexpected end of formula' : `unexpected ${describe(t)}`)
    }
  }
  // rows separated by `;`, values by `,`, all rows of one length, as Excel requires
  const constant = (at: number): Ast => {
    const rows: Lit[][] = [[]]
    for (;;) {
      rows.at(-1)!.push(literal())
      const t = toks[p++]
      if (t.t === ',') continue
      if (t.t === ';') {
        rows.push([])
        continue
      }
      if (t.t !== '}') bad(t.at, `expected , ; or } in an array constant, found ${describe(t)}`)
      break
    }
    if (rows.some((r) => r.length !== rows[0].length)) bad(at, 'the rows of an array constant must have the same number of values')
    return { k: 'arr', rows }
  }
  const literal = (): Lit => {
    const t = toks[p++]
    if (t.t === 'op' && (t.v === '-' || t.v === '+') && toks[p].t === 'num') {
      const n = toks[p++] as { v: number }
      return { k: 'num', v: t.v === '-' ? -n.v : n.v }
    }
    if (t.t === 'num') return { k: 'num', v: t.v }
    if (t.t === 'str') return { k: 'str', v: t.v }
    if (t.t === 'err' && t.sheet === undefined && t.ref === undefined) return { k: 'err', code: t.code }
    if (t.t === 'ident' && /^(TRUE|FALSE)$/i.test(t.v)) return { k: 'bool', v: t.v.toUpperCase() === 'TRUE' }
    return bad(t.at, `an array constant holds numbers, text, booleans and errors only, found ${describe(t)}`)
  }
  const expect = (k: '(' | ')') => {
    const t = toks[p]
    if (t.t !== k) bad(t.at, `expected ${k}, found ${describe(t)}`)
    p++
  }

  const ast = binary(0)
  const t = peek()
  if (t.t !== 'eof') bad(t.at, t.t === ',' ? 'comma outside of a function call' : `unexpected ${describe(t)}`)
  return ast
}

function describe(t: Tok): string {
  switch (t.t) {
    case 'op':
    case 'ident':
    case 'fn':
      return JSON.stringify(t.v)
    case 'eof':
      return 'end of formula'
    default:
      return t.t
  }
}
