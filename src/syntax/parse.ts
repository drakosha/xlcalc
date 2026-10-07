/**
 * The parser: recursive descent over the tokens, Excel's operator precedence, function calls with arguments left out,
 * array constants, the `@` operator, unions of references in parentheses (and at the top of a name's formula).
 */
import { MAX_FORMULA_DEPTH, MAX_FORMULA_LENGTH } from '../contract.js'
import { bad, type Ast, type BinOp, type Bind, type Lit } from './ast.js'
import { tokenize, type Tok } from './tokens.js'

const LEVELS: string[][] = [['=', '<>', '<', '>', '<=', '>='], ['&'], ['+', '-'], ['*', '/'], ['^']]

/** Excel loads a LET of 254 arguments at most (126 names and the calculation, or 127 names). */
const MAX_LET_ARGS = 254

/** `union`: a top-level comma joins references (a name's formula: `'S'!A1:A3,'S'!C1:C3`). */
export function parse(src: string, opts: { union?: boolean } = {}): Ast {
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
    let a = intersect()
    while (peek().t === 'op' && (peek() as { v: string }).v === '%') {
      p++
      a = { k: 'pct', a }
    }
    return a
  }
  // `@` takes the operand that follows it: a reference, a name, a call, a parenthesis
  const intersect = (): Ast => {
    const t = peek()
    if (t.t !== 'op' || t.v !== '@') return primary()
    p++
    enter(t.at)
    const a = intersect()
    depth--
    return { k: 'at', a }
  }
  // items joined by commas, as one union
  const joined = (first: Ast): Ast => {
    if (peek().t !== ',') return first
    const items = [first]
    while (peek().t === ',') {
      p++
      items.push(binary(0))
    }
    return { k: 'union', items }
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
        return t.last === undefined ? { k: 'ref', sheet: t.sheet, ref: t.ref } : { k: 'ref3d', first: t.sheet!, last: t.last, ref: t.ref }
      case 'table':
        return { k: 'table', table: t.table, spec: t.spec }
      case 'ident': {
        if (t.sheet !== undefined) return { k: 'name', name: t.v, sheet: t.sheet }
        const up = t.v.toUpperCase()
        if (up === 'TRUE' || up === 'FALSE') return { k: 'bool', v: up === 'TRUE' }
        // a LET name as it is declared, as Excel shows it
        const bind = scope.get(up)
        return bind ? { k: 'local', name: bind.name, slot: bind.slot } : { k: 'name', name: t.v }
      }
      case '(': {
        enter(t.at)
        const a = joined(binary(0))
        depth--
        expect(')')
        return { k: 'paren', a }
      }
      case '{':
        return constant(t.at)
      case 'fn': {
        expect('(')
        enter(t.at)
        if (t.v === 'LET') {
          const n = let_(t.at)
          depth--
          expect(')')
          return n
        }
        const args: Ast[] = []
        if (peek().t !== ')') {
          for (;;) {
            args.push(arg())
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
  // an argument, or nothing before the next `,` or `)`: left out
  const arg = (): Ast => (peek().t === ',' || peek().t === ')' ? { k: 'missing' } : binary(0))
  // LET's names in scope by upper-case name: a name is seen from the end of its value to the end of its LET
  const scope = new Map<string, Bind>()
  let slots = 0
  // LET(name, value, …, calculation); the calculation may be left off (Excel loads `LET(x,1)`: 1)
  const let_ = (at: number): Ast => {
    const binds: Bind[] = []
    const outer: [string, Bind | undefined][] = []
    let body: Ast | undefined
    for (;;) {
      const t = peek()
      if (t.t === 'ident' && toks[p + 1].t === ',' && !/^(TRUE|FALSE)$/i.test(t.v)) {
        const up = t.v.toUpperCase()
        if (binds.some((b) => b.name.toUpperCase() === up)) bad(t.at, `LET names ${t.v} twice`)
        p += 2
        const bind: Bind = { name: t.v, slot: slots++, value: arg() }
        binds.push(bind)
        outer.push([up, scope.get(up)])
        scope.set(up, bind)
        if (peek().t !== ',') break
        p++
        continue
      }
      body = arg()
      if (peek().t === ',') bad(t.at, `LET takes a name as its argument ${binds.length * 2 + 1}`)
      break
    }
    for (const [up, bind] of outer.reverse()) {
      if (bind) scope.set(up, bind)
      else scope.delete(up)
    }
    if (!binds.length) bad(at, 'LET takes a name and its value at least')
    if (binds.length * 2 + (body ? 1 : 0) > MAX_LET_ARGS) bad(at, `LET takes at most ${MAX_LET_ARGS} arguments`)
    return body ? { k: 'let', binds, body } : { k: 'let', binds }
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

  const ast = opts.union ? joined(binary(0)) : binary(0)
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
