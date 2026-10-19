/**
 * The parser: recursive descent over the tokens, Excel's operator precedence, function calls with arguments left out,
 * array constants, the `@` operator, the reference operators (`:` binds tightest, then the intersection, a space,
 * then `@`), unions of references in parentheses (and at the top of a name's formula).
 */
import { wordKey } from '../collate.js'
import { MAX_FORMULA_DEPTH, MAX_FORMULA_LENGTH } from '../contract.js'
import type { Found } from '../locale/messages.js'
import { bad, type Ast, type BinOp, type Bind, type Lit } from './ast.js'
import { readsAsCell, tokenize, type Tok } from './tokens.js'

const LEVELS: string[][] = [['=', '<>', '<', '>', '<=', '>='], ['&'], ['+', '-'], ['*', '/'], ['^']]
const REF_OPS = [' ', ':'] as const

/** Excel loads a LET of 254 arguments at most (126 names and the calculation, or 127 names). */
const MAX_LET_ARGS = 254

/**
 * A word LET does not take as a name: a boolean, or a cell that reads as R1C1's row or column too (`R2`, `C5`, `RC5`).
 * Any other cell is a name in LET, written without `$` (`A1`, `XFD1`), and so is R1C1's cell (`R1C1`, `RC`): the sets
 * `entry/names-held`, `let-grid-held` and `let-cells-held`, in Excel for Mac and Windows.
 */
export const letNameClash = (word: string) => boolWord(word) !== undefined || (readsAsCell(word) && /^(R|C|RC)[0-9]+$/i.test(word))

/** The boolean a word of formula text is (TRUE or FALSE in any case, by `wordKey`: `falſe` is a name), else undefined. */
export function boolWord(word: string): boolean | undefined {
  const key = wordKey(word)
  return key === 'TRUE' ? true : key === 'FALSE' ? false : undefined
}

/** `union`: a top-level comma joins references (a name's formula: `'S'!A1:A3,'S'!C1:C3`). */
export function parse(src: string, opts: { union?: boolean } = {}): Ast {
  return parseTokens(src, null, opts)
}

/** `parse` over the tokens of `src` read already (`given`, a formula's shape reads them first), or null to read them. */
export function parseTokens(src: string, given: Tok[] | null, opts: { union?: boolean } = {}): Ast {
  if (src.trim() === '') bad(0, { key: 'emptyFormula' }, 'empty')
  if (src.length > MAX_FORMULA_LENGTH) bad(MAX_FORMULA_LENGTH, { key: 'tooLong', max: MAX_FORMULA_LENGTH }, 'too_long')
  const toks = given ?? tokenize(src)
  let p = 0
  let depth = 0
  const peek = () => toks[p]
  const enter = (at: number) => {
    if (++depth > MAX_FORMULA_DEPTH) bad(at, { key: 'tooDeep', max: MAX_FORMULA_DEPTH }, 'too_deep')
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
    let a = single()
    while (peek().t === 'op' && (peek() as { v: string }).v === '%') {
      p++
      a = { k: 'pct', a }
    }
    return a
  }
  // `@` takes the operand that follows it: a reference, a name, a call, a parenthesis, their intersection
  const single = (): Ast => {
    const t = peek()
    if (t.t !== 'op' || t.v !== '@') return refs(0)
    p++
    enter(t.at)
    const a = single()
    depth--
    return { k: 'at', a }
  }
  // the reference operators: the intersection (a space) over the range operator (`:`) over operands
  const refs = (level: number): Ast => {
    if (level === REF_OPS.length) return primary()
    let a = refs(level + 1)
    for (;;) {
      const t = peek()
      if (t.t !== 'op' || t.v !== REF_OPS[level]) return a
      p++
      a = { k: 'bin', op: t.v, a, b: refs(level + 1) }
    }
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
        return { k: 'err', code: t.code, ...(t.sheet !== undefined ? { sheet: t.sheet } : {}), ...(t.ref !== undefined ? { ref: t.ref } : {}), ...(t.book !== undefined ? { book: t.book } : {}) }
      case 'ref': {
        // inside a LET that declares it, a cell's word is the name
        const word = cellWord(t)
        const bind = word === undefined ? undefined : scope.get(wordKey(word))
        if (bind) return { k: 'local', name: bind.name, slot: bind.slot }
        return t.last === undefined ? { k: 'ref', sheet: t.sheet, ref: t.ref } : { k: 'ref3d', first: t.sheet!, last: t.last, ref: t.ref }
      }
      case 'table':
        return { k: 'table', table: t.table, spec: t.spec }
      case 'xref':
        return { k: 'xref', book: t.book, sheet: t.sheet, ref: t.ref }
      case 'xname':
        return { k: 'xname', book: t.book, name: t.v, ...(t.sheet === undefined ? {} : { sheet: t.sheet }) }
      case 'ident': {
        if (t.sheet !== undefined) return { k: 'name', name: t.v, sheet: t.sheet }
        const bool = boolWord(t.v)
        if (bool !== undefined) return { k: 'bool', v: bool }
        // a LET name as it is declared, as Excel shows it
        const bind = scope.get(wordKey(t.v))
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
        return bad(t.at, t.t === 'eof' ? { key: 'unexpectedEnd' } : { key: 'unexpected', found: found(t) })
    }
  }
  // an argument, or nothing before the next `,` or `)`: left out
  const arg = (): Ast => (peek().t === ',' || peek().t === ')' ? { k: 'missing' } : binary(0))
  // LET's names in scope by upper-case name: a name is seen from the end of its value to the end of its LET
  const scope = new Map<string, Bind>()
  // a cell written without `$` that LET takes as a name, as typed; undefined for any other token
  const cellWord = (t: Tok): string | undefined => {
    if (t.t !== 'ref' || t.sheet !== null || t.ref.kind !== 'cell' || t.ref.spill) return undefined
    const word = src.slice(t.at, t.at + t.ref.text.length)
    return readsAsCell(word) && !letNameClash(word) ? word : undefined
  }
  let slots = 0
  // LET(name, value, …, calculation); the calculation may be left off (Excel loads `LET(x,1)`: 1)
  const let_ = (at: number): Ast => {
    const binds: Bind[] = []
    const outer: [string, Bind | undefined][] = []
    let body: Ast | undefined
    for (;;) {
      const t = peek()
      const name = t.t === 'ident' && boolWord(t.v) === undefined ? t.v : cellWord(t)
      if (name !== undefined && toks[p + 1].t === ',') {
        const up = wordKey(name)
        if (binds.some((b) => wordKey(b.name) === up)) bad(t.at, { key: 'letTwice', name })
        p += 2
        const bind: Bind = { name, slot: slots++, value: arg() }
        binds.push(bind)
        outer.push([up, scope.get(up)])
        scope.set(up, bind)
        if (peek().t !== ',') break
        p++
        continue
      }
      body = arg()
      if (peek().t === ',') bad(t.at, { key: 'letName', position: binds.length * 2 + 1 })
      break
    }
    for (const [up, bind] of outer.reverse()) {
      if (bind) scope.set(up, bind)
      else scope.delete(up)
    }
    if (!binds.length) bad(at, { key: 'letEmpty' })
    if (binds.length * 2 + (body ? 1 : 0) > MAX_LET_ARGS) bad(at, { key: 'letTooMany', max: MAX_LET_ARGS })
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
      if (t.t !== '}') bad(t.at, { key: 'arraySeparator', found: found(t) })
      break
    }
    if (rows.some((r) => r.length !== rows[0].length)) bad(at, { key: 'arrayRows' })
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
    const bool = t.t === 'ident' ? boolWord(t.v) : undefined
    if (bool !== undefined) return { k: 'bool', v: bool }
    return bad(t.at, { key: 'arrayValue', found: found(t) })
  }
  const expect = (k: '(' | ')') => {
    const t = toks[p]
    if (t.t !== k) bad(t.at, { key: 'expected', char: k, found: found(t) })
    p++
  }

  const ast = opts.union ? joined(binary(0)) : binary(0)
  const t = peek()
  if (t.t !== 'eof') bad(t.at, t.t === ',' ? { key: 'commaOutside' } : { key: 'unexpected', found: found(t) })
  return ast
}

/** A token for a message: its kind, with its text for an operator, a name or a function. */
const found = (t: Tok): Found => (t.t === 'op' || t.t === 'ident' || t.t === 'fn' ? { t: t.t, v: t.v } : t.t === 'xname' ? { t: 'ident', v: t.v } : { t: t.t === 'xref' ? 'ref' : t.t })
