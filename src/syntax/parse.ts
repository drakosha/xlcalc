/**
 * The parser: recursive descent over the tokens, Excel's operator precedence, function calls with arguments left out,
 * array constants, the `@` operator, the reference operators (`:` binds tightest, then the intersection, a space,
 * then `@`), unions of references in parentheses (and at the top of a name's formula).
 */
import { wordKey } from '../collate.js'
import { MAX_FORMULA_DEPTH, MAX_FORMULA_LENGTH } from '../contract.js'
import type { Found } from '../locale/messages.js'
import { bad, type Ast, type BinOp, type Bind, type Lit, type Param } from './ast.js'
import { IDENT_RULE } from './scan.js'
import { readsAsCell, tokenize, type Tok } from './tokens.js'

const LEVELS: string[][] = [['=', '<>', '<', '>', '<=', '>='], ['&'], ['+', '-'], ['*', '/'], ['^']]
const REF_OPS = [' ', ':'] as const

/**
 * The operators of one formula. 8,192 characters hold 4,096 of them, which the walks over
 * the tree do not survive on a default stack (from about 2,100, depending on what the JIT has done by then). Real files stay
 * far below: of 254,683 formulas in 1,123 user workbooks, the longest has 86.
 */
const MAX_OPERATORS = 1024
/** Excel loads a LET of 254 arguments at most (126 names and the calculation, or 127 names). */
const MAX_LET_ARGS = 254
/** LAMBDA takes 253 parameters at most, as Microsoft's documentation says. */
const MAX_PARAMS = 253
/** What an optional parameter's brackets hold: a name. */
const WORD = new RegExp(`^${IDENT_RULE}$`, 'u')

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
  let chained = 0
  // operators nest without a parenthesis, `1+1+1` in `(1+1)+1`, and every walk over the tree recurses
  const chain = (at: number) => {
    if (++chained > MAX_OPERATORS) bad(at, { key: 'tooManyOperators', max: MAX_OPERATORS }, 'too_deep')
  }

  const binary = (level: number): Ast => {
    if (level === LEVELS.length) return unary()
    let a = binary(level + 1)
    for (;;) {
      const t = peek()
      if (t.t !== 'op' || !LEVELS[level].includes(t.v)) return a
      chain(t.at)
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
      chain(peek().at)
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
      chain(t.at)
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
  // a value called as a function right after it: `LAMBDA(x,x*2)(3)`, `f(1)(2)`
  const primary = (): Ast => {
    let a = atom()
    while (peek().t === '(' && (a.k === 'fn' || a.k === 'let' || a.k === 'lambda' || a.k === 'call' || a.k === 'paren' || a.k === 'local')) {
      const at = peek().at
      chain(at)
      p++
      a = { k: 'call', fn: a, args: callArgs(at) }
    }
    return a
  }
  const atom = (): Ast => {
    const t = toks[p++]
    switch (t.t) {
      case 'num':
        return { k: 'num', v: t.v }
      case 'str':
        return { k: 'str', v: t.v }
      case 'err':
        return { k: 'err', code: t.code, ...(t.sheet !== undefined ? { sheet: t.sheet } : {}), ...(t.ref !== undefined ? { ref: t.ref } : {}), ...(t.book !== undefined ? { book: t.book } : {}), ...(t.table !== undefined ? { table: t.table } : {}) }
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
        if (t.v === 'LET' || t.v === 'LAMBDA') {
          const n = t.v === 'LET' ? let_(t.at) : lambda_(t.at)
          depth--
          expect(')')
          return n
        }
        // a LET name or a parameter called (`f(2)`)
        const bind = scope.get(t.v)
        const args = callArgs(t.at, false)
        return bind ? { k: 'call', fn: { k: 'local', name: bind.name, slot: bind.slot }, args } : { k: 'fn', name: t.v, args }
      }
      default:
        return bad(t.at, t.t === 'eof' ? { key: 'unexpectedEnd' } : { key: 'unexpected', found: found(t) })
    }
  }
  // the arguments of a call up to its `)`, its `(` read already; `open`: its nesting not entered yet
  const callArgs = (at: number, open = true): Ast[] => {
    if (open) enter(at)
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
    return args
  }
  // an argument, or nothing before the next `,` or `)`: left out
  const arg = (): Ast => (peek().t === ',' || peek().t === ')' ? { k: 'missing' } : binary(0))
  // LET's names and LAMBDA's parameters in scope by upper-case name: a LET name is seen from the end of its value to
  // the end of its LET, a parameter in its LAMBDA's calculation
  const scope = new Map<string, Param>()
  // a cell written without `$` that LET and LAMBDA take as a name, as typed; undefined for any other token
  const cellWord = (t: Tok): string | undefined => {
    if (t.t !== 'ref' || t.sheet !== null || t.ref.kind !== 'cell' || t.ref.spill) return undefined
    const word = src.slice(t.at, t.at + t.ref.text.length)
    return readsAsCell(word) && !letNameClash(word) ? word : undefined
  }
  let slots = 0
  // the name a LET or a LAMBDA declares at the next token, followed by `,`: a word that is no boolean, or a cell's word
  const declared = (): string | undefined => {
    const t = peek()
    const name = t.t === 'ident' && boolWord(t.v) === undefined ? t.v : cellWord(t)
    return name !== undefined && toks[p + 1].t === ',' ? name : undefined
  }
  const twice = (list: readonly Param[], name: string) => list.some((x) => wordKey(x.name) === wordKey(name))
  // a declared name seen from now on, the outer one it shadows kept for `restore`
  const see = (x: Param, outer: [string, Param | undefined][]) => {
    const up = wordKey(x.name)
    outer.push([up, scope.get(up)])
    scope.set(up, x)
  }
  // LET(name, value, …, calculation); the calculation may be left off (Excel loads `LET(x,1)`: 1)
  const let_ = (at: number): Ast => {
    const binds: Bind[] = []
    const outer: [string, Param | undefined][] = []
    let body: Ast | undefined
    for (;;) {
      const t = peek()
      const name = declared()
      if (name !== undefined) {
        if (twice(binds, name)) bad(t.at, { key: 'letTwice', name })
        p += 2
        const bind: Bind = { name, slot: slots++, value: arg() }
        binds.push(bind)
        see(bind, outer)
        if (peek().t !== ',') break
        p++
        continue
      }
      body = arg()
      if (peek().t === ',') bad(t.at, { key: 'letName', position: binds.length * 2 + 1 })
      break
    }
    restore(outer)
    if (!binds.length) bad(at, { key: 'letEmpty' })
    if (binds.length * 2 + (body ? 1 : 0) > MAX_LET_ARGS) bad(at, { key: 'letTooMany', max: MAX_LET_ARGS })
    return body ? { k: 'let', binds, body } : { k: 'let', binds }
  }
  // the names a LET or a LAMBDA declared out of scope again, the outer ones they shadowed back
  const restore = (outer: [string, Param | undefined][]) => {
    for (const [up, bind] of outer.reverse()) {
      if (bind) scope.set(up, bind)
      else scope.delete(up)
    }
  }
  // a LAMBDA's next parameter and whether it is optional (`[y]`, which a call may leave out)
  const parameter = (): [string | undefined, boolean] => {
    const t = peek()
    const name = t.t === 'table' && t.table === null && !t.spec.areas.length && t.spec.cols?.length === 1 ? t.spec.cols[0] : undefined
    return name !== undefined && WORD.test(name) && toks[p + 1].t === ',' ? [name, true] : [declared(), false]
  }
  // LAMBDA(parameter, …, calculation): the parameters are seen in the calculation
  const lambda_ = (at: number): Ast => {
    const params: Param[] = []
    for (let [name, optional] = parameter(); name !== undefined; [name, optional] = parameter()) {
      if (name.includes('.')) bad(peek().at, { key: 'lambdaDot', name })
      if (twice(params, name)) bad(peek().at, { key: 'lambdaTwice', name })
      // as Excel, which takes no required parameter after an optional one
      if (!optional && params.at(-1)?.optional) bad(peek().at, { key: 'lambdaName', position: params.length + 1 })
      params.push(optional ? { name, slot: slots++, optional } : { name, slot: slots++ })
      p += 2
    }
    if (params.length > MAX_PARAMS) bad(at, { key: 'lambdaTooMany', max: MAX_PARAMS })
    const outer: [string, Param | undefined][] = []
    for (const x of params) see(x, outer)
    const first = peek()
    const body = arg()
    restore(outer)
    if (peek().t === ',') bad(first.at, { key: 'lambdaName', position: params.length + 1 })
    if (body.k === 'missing') bad(at, { key: 'lambdaEmpty' })
    return { k: 'lambda', params, body }
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
    if (t.t === 'err' && t.sheet === undefined && t.ref === undefined && t.book === undefined) return { k: 'err', code: t.code }
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
