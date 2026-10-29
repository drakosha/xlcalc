/**
 * The parser of match mode 3's patterns, PCRE2's syntax (`pcre2pattern`, `pcre2syntax`, release 10.43 and later) with
 * UTF and UCP on, as Excel for Mac takes it: a pattern into a tree of nodes, or a `RegexSyntaxError` where PCRE2 gives a
 * compile error (`#VALUE!`). Callouts and `\C` are errors, as in Excel.
 */
import { typeClass } from '../functions/lookup/regex-syntax.js'
import { type Ascii, type CharTest, classOf, classTest, Newline, posixClass } from './chars.js'
import { H, hex, property, V } from './classes.js'

export class RegexSyntaxError extends Error {}

/** Where a back reference, a call or a condition points: group numbers, or a name resolved after the parse. */
export interface Ref {
  groups: number[]
  name?: string
}

export type Assert = 'bol' | 'mbol' | 'eol' | 'meol' | 'A' | 'z' | 'Z' | 'G' | 'b' | 'B' | 'ab' | 'aB'
export type Verb = 'accept' | 'fail' | 'commit' | 'prune' | 'skip' | 'then' | 'mark'
export type Cond = { k: 'group'; ref: Ref } | { k: 'R'; ref: Ref | null } | { k: 'define' } | { k: 'look'; look: Node & { t: 'look' } } | { k: 'true' } | { k: 'false' }

export type Node =
  | { t: 'char'; c: number; i: boolean; r: boolean }
  /** `js`: the class in JavaScript (flag `v`); a bracketed one also by its parts, for the translation's caseless characters. */
  | { t: 'set'; test: CharTest; js: string; ranges?: number[]; parts?: string[]; negated?: boolean; caseless?: boolean; r?: boolean }
  | { t: 'any'; all: boolean }
  | { t: 'seq'; items: Node[] }
  | { t: 'alt'; alts: Node[] }
  | { t: 'group'; cap: number; atomic: boolean; body: Node }
  | { t: 'look'; behind: boolean; neg: boolean; body: Node }
  | { t: 'rep'; body: Node; min: number; max: number; mode: 0 | 1 | 2 }
  | { t: 'assert'; k: Assert }
  | { t: 'bref'; ref: Ref; i: boolean; r: boolean }
  | { t: 'call'; ref: Ref }
  | { t: 'cond'; cond: Cond; yes: Node; no: Node }
  | { t: 'verb'; k: Verb; name: string | null }
  | { t: 'nl' }
  | { t: 'cluster' }
  | { t: 'empty' }

/** What precedes the pattern: `(*LIMIT_MATCH=…)`, the newline convention, `\R`'s, `(*NOTEMPTY)`. */
export interface Head {
  limit: number
  newline: Newline
  anyR: boolean
  notEmpty: boolean
  notEmptyAtStart: boolean
}

export interface Parsed {
  root: Node
  groups: number
  /** The node of each capturing group by number (1…). */
  bodies: (Node & { t: 'group' })[]
  head: Head
}

interface Flags extends Ascii {
  i: boolean
  m: boolean
  s: boolean
  x: boolean
  xx: boolean
  n: boolean
  U: boolean
  J: boolean
  r: boolean
}

const MAX_REPEAT = 65535
const MAX_NEST = 250
const empty: Node = { t: 'empty' }
const seq = (items: Node[]): Node => (items.length === 1 ? items[0] : { t: 'seq', items })
const isDigit = (c: number) => c >= 48 && c <= 57
const isWordStart = (ch: string) => /^[_\p{L}]$/u.test(ch)
const isWordPart = (ch: string) => /^[_\p{L}\p{N}]$/u.test(ch)

const HEAD_NEWLINES: Record<string, Newline> = { CR: Newline.CR, LF: Newline.LF, CRLF: Newline.CRLF, ANYCRLF: Newline.ANYCRLF, ANY: Newline.ANY, NUL: Newline.NUL }
const HEAD_IGNORED = new Set(['UTF', 'UCP', 'NO_AUTO_POSSESS', 'NO_DOTSTAR_ANCHOR', 'NO_JIT', 'NO_START_OPT'])
const ALPHA_GROUPS: Record<string, string> = {
  pla: '=', positive_lookahead: '=', nla: '!', negative_lookahead: '!', plb: '<=', positive_lookbehind: '<=', nlb: '<!', negative_lookbehind: '<!',
  napla: '=', non_atomic_positive_lookahead: '=', naplb: '<=', non_atomic_positive_lookbehind: '<=', atomic: '>',
}
const VERBS: Record<string, Verb> = { ACCEPT: 'accept', FAIL: 'fail', F: 'fail', COMMIT: 'commit', PRUNE: 'prune', SKIP: 'skip', THEN: 'then', MARK: 'mark', '': 'mark' }

class Parser {
  i = 0
  groups = 0
  readonly bodies: (Node & { t: 'group' })[] = []
  readonly names = new Map<string, number[]>()
  readonly refs: Ref[] = []
  readonly head: Head = { limit: Infinity, newline: Newline.ANYCRLF, anyR: true, notEmpty: false, notEmptyAtStart: false }
  depth = 0

  constructor(readonly p: string) {}

  fail(): never {
    throw new RegexSyntaxError(`invalid pattern at ${this.i}`)
  }

  peek(k = 0): string {
    return this.p.charAt(this.i + k)
  }

  eat(s: string): boolean {
    if (!this.p.startsWith(s, this.i)) return false
    this.i += s.length
    return true
  }

  /** The code point at the cursor, consumed. */
  next(): number {
    if (this.i >= this.p.length) this.fail()
    const c = this.p.codePointAt(this.i)!
    this.i += c > 0xffff ? 2 : 1
    return c
  }

  parseHead() {
    for (;;) {
      const m = /^\(\*([A-Z_]+)(?:=(\d+))?\)/.exec(this.p.slice(this.i))
      if (!m) return
      const [all, name, value] = m
      if (value !== undefined) {
        if (name === 'LIMIT_MATCH') this.head.limit = Math.min(this.head.limit, Number(value))
        else if (name !== 'LIMIT_DEPTH' && name !== 'LIMIT_HEAP' && name !== 'LIMIT_RECURSION') this.fail()
      } else if (name in HEAD_NEWLINES) this.head.newline = HEAD_NEWLINES[name]
      else if (name === 'BSR_ANYCRLF' || name === 'BSR_UNICODE') this.head.anyR = name === 'BSR_UNICODE'
      else if (name === 'NOTEMPTY') this.head.notEmpty = true
      else if (name === 'NOTEMPTY_ATSTART') this.head.notEmptyAtStart = true
      else if (!HEAD_IGNORED.has(name)) return
      this.i += all.length
    }
  }

  /** Alternatives up to `)` or the end; option settings carry from one alternative to the next. */
  parseAlts(f: Flags, reset = false): Node {
    const alts: Node[] = []
    const base = this.groups
    let most = base
    for (;;) {
      if (reset) this.groups = base
      alts.push(this.parseSeq(f))
      most = Math.max(most, this.groups)
      if (!this.eat('|')) break
    }
    this.groups = most
    return alts.length === 1 ? alts[0] : { t: 'alt', alts }
  }

  skipSpace(f: Flags) {
    if (!f.x) return
    for (;;) {
      const c = this.peek()
      if (c === '#') {
        while (this.i < this.p.length && this.p[this.i] !== '\n') this.i++
      } else if (c !== '' && /\s/.test(c)) this.i++
      else return
    }
  }

  parseSeq(f: Flags): Node {
    const items: Node[] = []
    for (;;) {
      this.skipSpace(f)
      const c = this.peek()
      if (c === '' || c === '|' || c === ')') return seq(items)
      if (this.eat('\\Q')) {
        const end = this.p.indexOf('\\E', this.i)
        const text = end < 0 ? this.p.slice(this.i) : this.p.slice(this.i, end)
        this.i = end < 0 ? this.p.length : end + 2
        for (const ch of text) items.push(this.quantified(f, this.char(ch.codePointAt(0)!, f), true))
        continue
      }
      if (this.eat('\\E')) continue
      const [atom, repeatable] = this.atom(f)
      items.push(this.quantified(f, atom, repeatable))
    }
  }

  char(c: number, f: Flags): Node {
    return { t: 'char', c, i: f.i, r: f.r }
  }

  /** A quantifier at the cursor: [min, max] or null (`{` that is no quantifier is a literal). */
  quantifier(): [number, number] | null {
    const c = this.peek()
    if (c === '*') return (this.i++, [0, Infinity])
    if (c === '+') return (this.i++, [1, Infinity])
    if (c === '?') return (this.i++, [0, 1])
    if (c !== '{') return null
    const m = /^\{\s*(\d*)\s*(?:(,)\s*(\d*)\s*)?\}/.exec(this.p.slice(this.i))
    if (!m || (m[1] === '' && (m[2] === undefined || m[3] === ''))) return null
    this.i += m[0].length
    const min = m[1] === '' ? 0 : Number(m[1])
    const max = m[2] === undefined ? min : m[3] === '' ? Infinity : Number(m[3])
    if (min > MAX_REPEAT || (max !== Infinity && max > MAX_REPEAT) || min > max) this.fail()
    return [min, max]
  }

  quantified(f: Flags, atom: Node, repeatable: boolean): Node {
    this.skipSpace(f)
    const q = this.quantifier()
    if (!q) return atom
    if (!repeatable) this.fail()
    let mode: 0 | 1 | 2 = f.U ? 1 : 0
    if (this.eat('+')) mode = 2
    else if (this.eat('?')) mode = f.U ? 0 : 1
    this.skipSpace(f)
    const save = this.i
    if (this.quantifier()) this.fail()
    this.i = save
    const [min, max] = q
    // an assertion repeated is obeyed once, or never with {0}, or as optional with {0,n}
    if (atom.t === 'look') return max === 0 ? empty : min === 0 ? { t: 'rep', body: atom, min: 0, max: 1, mode: 0 } : atom
    return { t: 'rep', body: atom, min, max, mode }
  }

  /** One item and whether a quantifier may follow it. */
  atom(f: Flags): [Node, boolean] {
    const c = this.peek()
    if (c === '(') return this.group(f)
    if (c === '[') return [this.charClass(f), true]
    if (c === '.') return (this.i++, [{ t: 'any', all: f.s }, true])
    if (c === '^') return (this.i++, [{ t: 'assert', k: f.m ? 'mbol' : 'bol' }, false])
    if (c === '$') return (this.i++, [{ t: 'assert', k: f.m ? 'meol' : 'eol' }, false])
    if (c === '\\') return this.escape(f)
    if (c === '*' || c === '+' || c === '?') this.fail()
    if (c === '{') {
      const save = this.i
      if (this.quantifier()) this.fail()
      this.i = save
    }
    return [this.char(this.next(), f), true]
  }

  /** A number in decimal at the cursor, or -1. */
  num(): number {
    const m = /^\d+/.exec(this.p.slice(this.i))
    if (!m) return -1
    this.i += m[0].length
    return Number(m[0])
  }

  name(): string {
    const start = this.i
    if (!isWordStart(String.fromCodePoint(this.p.codePointAt(this.i) ?? 0))) this.fail()
    while (this.i < this.p.length && isWordPart(String.fromCodePoint(this.p.codePointAt(this.i)!))) this.i += this.p.codePointAt(this.i)! > 0xffff ? 2 : 1
    if (this.i - start > 128) this.fail()
    return this.p.slice(start, this.i)
  }

  /** A reference by number (relative ones resolved now) or by name, its closing text consumed. */
  ref(close: string, relative: boolean): Ref {
    const sign = relative && (this.peek() === '+' || this.peek() === '-') ? this.p[this.i++] : ''
    let ref: Ref
    if (isDigit(this.p.charCodeAt(this.i))) {
      const n = this.num()
      if (sign === '-' && n === 0) this.fail()
      ref = { groups: [sign === '-' ? this.groups - n + 1 : sign === '+' ? this.groups + n : n] }
      if (ref.groups[0] <= 0 && sign) this.fail()
    } else if (sign) this.fail()
    else ref = this.named(this.name())
    if (!this.eat(close)) this.fail()
    return ref!
  }

  named(name: string): Ref {
    const ref: Ref = { groups: [], name }
    this.refs.push(ref)
    return ref
  }

  /** A back reference, its closing text consumed. */
  bref(f: Flags, close: string, relative = true): [Node, boolean] {
    return [{ t: 'bref', ref: this.ref(close, relative), i: f.i, r: f.r }, true]
  }

  escape(f: Flags): [Node, boolean] {
    this.i++
    const c = this.peek()
    if (c === '') this.fail()
    const code = this.p.charCodeAt(this.i)
    if (isDigit(code) && c !== '0') {
      const save = this.i
      const n = this.num()
      if (n < 10 || c === '8' || c === '9' || n <= this.groups) return [{ t: 'bref', ref: { groups: [n] }, i: f.i, r: f.r }, true]
      this.i = save
      return [this.char(this.octal(), f), true]
    }
    this.i++
    switch (c) {
      case 'g':
        if (this.eat('{')) return this.bref(f, '}')
        if (this.eat('<')) return [{ t: 'call', ref: this.ref('>', true) }, true]
        if (this.eat("'")) return [{ t: 'call', ref: this.ref("'", true) }, true]
        return this.bref(f, '')
      case 'k':
        if (this.eat('<')) return this.bref(f, '>', false)
        if (this.eat("'")) return this.bref(f, "'", false)
        if (this.eat('{')) return this.bref(f, '}', false)
        return this.fail()
      case 'b':
        return [{ t: 'assert', k: f.aW ? 'ab' : 'b' }, false]
      case 'B':
        return [{ t: 'assert', k: f.aW ? 'aB' : 'B' }, false]
      case 'A':
      case 'z':
      case 'Z':
      case 'G':
        return [{ t: 'assert', k: c }, false]
      case 'K':
        return [empty, false]
      case 'R':
        return [{ t: 'nl' }, true]
      case 'X':
        return [{ t: 'cluster' }, true]
      case 'N':
        if (this.peek() === '{' && this.peek(1) === 'U') return [this.char(this.charEscape('N'), f), true]
        return [{ t: 'any', all: false }, true]
    }
    const js = this.typeEscape(c, f)
    if (js) return [{ t: 'set', test: classOf(js), js }, true]
    this.i--
    return [this.char(this.charEscape(this.p[this.i++]), f), true]
  }

  octal(): number {
    const m = /^[0-7]{1,3}/.exec(this.p.slice(this.i))!
    this.i += m[0].length
    return parseInt(m[0], 8)
  }

  /** `\d \D \s \S \w \W \h \H \v \V \p \P` (the letter consumed) as a JavaScript class, or null. */
  typeEscape(c: string, f: Flags): string | null {
    const lower = c.toLowerCase()
    if (lower === 'd' || lower === 's' || lower === 'w') return typeClass(c, { d: f.aD, s: f.aS, w: f.aW })
    if (lower === 'h' || lower === 'v') return c === lower ? (c === 'h' ? H : V) : `[^${c === 'H' ? H : V}]`
    if (lower !== 'p') return null
    let name: string
    let neg = c === 'P'
    if (this.eat('{')) {
      const end = this.p.indexOf('}', this.i)
      if (end < 0) this.fail()
      name = this.p.slice(this.i, end)
      this.i = end + 1
      if (name.startsWith('^')) ((neg = !neg), (name = name.slice(1)))
    } else name = this.p[this.i++] ?? ''
    return property(name, neg, f.i) ?? this.fail()
  }

  /** The code point of a character escape whose letter is consumed (`\x41`, `\n`, `\.`); letters without a meaning are errors. */
  charEscape(c: string): number {
    switch (c) {
      case 'a':
        return 7
      case 'e':
        return 27
      case 'f':
        return 12
      case 'n':
        return 10
      case 'r':
        return 13
      case 't':
        return 9
      case '0':
        this.i--
        return this.octal()
      case 'c': {
        const x = this.p.charCodeAt(this.i++)
        if (!(x >= 32 && x < 127)) this.fail()
        return (x >= 97 && x <= 122 ? x - 32 : x) ^ 0x40
      }
      case 'o':
        return this.braced(/^\{([0-7]+)\}/, 8)
      case 'x':
        if (this.peek() === '{') return this.braced(/^\{\s*([0-9a-fA-F]+)\s*\}/, 16)
        return this.digits(/^[0-9a-fA-F]{0,2}/, 16)
      case 'N':
        return this.braced(/^\{U\+([0-9a-fA-F]+)\}/, 16)
    }
    if (/^[\p{L}\p{N}]$/u.test(c)) this.fail()
    this.i--
    return this.next()
  }

  digits(re: RegExp, base: number): number {
    const m = re.exec(this.p.slice(this.i))!
    this.i += m[0].length
    return m[0] === '' ? 0 : parseInt(m[0], base)
  }

  braced(re: RegExp, base: number): number {
    const m = re.exec(this.p.slice(this.i))
    if (!m) this.fail()
    this.i += m[0].length
    const v = parseInt(m[1], base)
    if (v > 0x10ffff || (v >= 0xd800 && v <= 0xdfff)) this.fail()
    return v
  }

  charClass(f: Flags): Node {
    this.i++
    const negated = this.eat('^')
    const ranges: number[] = []
    const parts: string[] = []
    let first = true
    for (;;) {
      if (this.i >= this.p.length) this.fail()
      if (f.xx && (this.peek() === ' ' || this.peek() === '\t')) {
        this.i++
        continue
      }
      if (this.peek() === ']' && !first) {
        this.i++
        break
      }
      first = false
      if (this.eat('\\Q')) {
        const end = this.p.indexOf('\\E', this.i)
        const text = end < 0 ? this.p.slice(this.i) : this.p.slice(this.i, end)
        this.i = end < 0 ? this.p.length : end + 2
        for (const ch of text) ranges.push(ch.codePointAt(0)!, ch.codePointAt(0)!)
        continue
      }
      if (this.eat('\\E')) continue
      const lo = this.classAtom(f, parts)
      if (lo < 0) {
        // after a class such as \d a hyphen is literal only before the end
        if (this.peek() === '-' && this.peek(1) !== ']') this.fail()
        continue
      }
      if (this.peek() === '-' && this.peek(1) !== ']' && this.peek(1) !== '') {
        const save = this.i++
        if (this.peek() === '[' && /^\[:\^?[a-z]+:\]/.test(this.p.slice(this.i))) {
          this.i = save
          ranges.push(lo, lo)
          continue
        }
        const hi = this.classAtom(f, parts)
        if (hi < 0 || hi < lo) this.fail()
        ranges.push(lo, hi)
      } else ranges.push(lo, lo)
    }
    let js = negated ? '[^' : '['
    for (let k = 0; k < ranges.length; k += 2) js += ranges[k] === ranges[k + 1] ? hex(ranges[k]) : `${hex(ranges[k])}-${hex(ranges[k + 1])}`
    js += `${parts.join('')}]`
    // caseless, the ranges take the characters of their case at each test; the rest is the class JavaScript reads
    const test = f.i ? classTest(ranges, parts.map(classOf), negated, true, f.r) : classOf(js)
    return { t: 'set', test, js, ranges, parts, negated, caseless: f.i, r: f.r }
  }

  /** One character of a class (its code point), or a class within it added to `parts` in JavaScript (-1). */
  classAtom(f: Flags, parts: string[]): number {
    const posix = /^\[:(\^?)([a-z]+):\]/.exec(this.p.slice(this.i))
    if (posix) {
      const js = posixClass(posix[2], f, f.i) ?? this.fail()
      this.i += posix[0].length
      parts.push(posix[1] ? `[^${js}]` : js)
      return -1
    }
    if (/^\[[.=]/.test(this.p.slice(this.i)) && /^\[([.=]).*?\1\]/.test(this.p.slice(this.i))) this.fail()
    if (!this.eat('\\')) return this.next()
    const c = this.p[this.i++] ?? this.fail()
    if (c === 'b') return 8
    if (isDigit(c.charCodeAt(0))) {
      if (c === '8' || c === '9') this.fail()
      this.i--
      return this.octal()
    }
    if ('RXBNgkKAzZG'.includes(c) && !(c === 'N' && this.peek() === '{')) this.fail()
    const js = this.typeEscape(c, f)
    if (js) {
      parts.push(js)
      return -1
    }
    return this.charEscape(c)
  }

  /** Option letters up to `)` or `:` (`i`, `-x`, `^`, `aD`…) applied to `f`. */
  options(f: Flags): void {
    let on = true
    if (this.eat('^')) Object.assign(f, { i: false, m: false, n: false, s: false, x: false, xx: false })
    for (;;) {
      const c = this.peek()
      if (c === ')' || c === ':') return
      this.i++
      if (c === '-' && on) on = false
      else if (c === 'x') {
        const double = this.eat('x')
        f.x = on
        if (double || !on) f.xx = on && double
      } else if (c === 'a') {
        const k = 'DSWPT'.includes(this.peek()) && this.peek() !== '' ? this.p[this.i++] : ''
        for (const key of k ? [k] : ['D', 'S', 'W', 'P', 'T']) f[`a${key}` as keyof Ascii] = on
      } else if ('imnsUJr'.includes(c) && c !== '') f[c as 'i' | 'm' | 'n' | 's' | 'U' | 'J' | 'r'] = on
      else this.fail()
    }
  }

  group(f: Flags): [Node, boolean] {
    this.i++
    if (++this.depth > MAX_NEST) this.fail()
    try {
      return this.groupBody(f)
    } finally {
      this.depth--
    }
  }

  close() {
    if (!this.eat(')')) this.fail()
  }

  /** The inside of a group after `(`, its `)` consumed. */
  groupBody(f: Flags): [Node, boolean] {
    if (this.eat('*')) return this.starGroup(f)
    if (!this.eat('?')) {
      if (f.n) return [this.inner(f, -1, false), true]
      return [this.capture(f, null), true]
    }
    const c = this.peek()
    if (c === '#') {
      const end = this.p.indexOf(')', this.i)
      if (end < 0) this.fail()
      this.i = end + 1
      return [empty, false]
    }
    if (this.eat(':')) return [this.inner(f, -1, false), true]
    if (this.eat('|')) {
      const body = this.parseAlts({ ...f }, true)
      this.close()
      return [{ t: 'group', cap: -1, atomic: false, body }, true]
    }
    if (this.eat('>')) return [this.inner(f, -1, true), true]
    if (this.eat('=') || this.eat('*')) return [this.look(f, false, false), true]
    if (this.eat('!')) return [this.look(f, false, true), true]
    if (this.eat('<=') || this.eat('<*')) return [this.look(f, true, false), true]
    if (this.eat('<!')) return [this.look(f, true, true), true]
    if (this.eat('P<') || this.eat('<')) return [this.capture(f, this.nameUntil('>')), true]
    if (this.eat("'")) return [this.capture(f, this.nameUntil("'")), true]
    if (this.eat('P=')) return this.bref(f, ')', false)
    if (this.eat('P>') || this.eat('&')) return [{ t: 'call', ref: this.ref(')', false) }, true]
    if (this.eat('R)')) return [{ t: 'call', ref: { groups: [0] } }, true]
    if (c === '+' || c === '-' ? isDigit(this.p.charCodeAt(this.i + 1)) : isDigit(this.p.charCodeAt(this.i))) return [{ t: 'call', ref: this.ref(')', true) }, true]
    if (this.eat('(')) return [this.conditional(f), true]
    if (c === 'C') this.fail()
    // (?i) sets options for the rest of the enclosing group, (?i:…) inside its own
    const own = { ...f }
    this.options(own)
    if (this.eat(')')) return (Object.assign(f, own), [empty, false])
    this.i++
    return [this.inner(own, -1, false, true), true]
  }

  nameUntil(close: string): string {
    const name = this.name()
    if (!this.eat(close)) this.fail()
    return name
  }

  /** The alternatives of a group and its `)`. */
  inner(f: Flags, cap: number, atomic: boolean, scoped = false): Node {
    const body = this.parseAlts(scoped ? f : { ...f })
    this.close()
    return { t: 'group', cap, atomic, body }
  }

  capture(f: Flags, name: string | null): Node {
    const n = ++this.groups
    if (name !== null) {
      const had = this.names.get(name)
      if (had && !f.J && !had.includes(n)) this.fail()
      this.names.set(name, had ? [...new Set([...had, n])] : [n])
    }
    const g: Node & { t: 'group' } = { t: 'group', cap: n, atomic: false, body: empty }
    this.bodies[n] ??= g
    g.body = this.parseAlts({ ...f })
    this.close()
    return g
  }

  look(f: Flags, behind: boolean, neg: boolean): Node & { t: 'look' } {
    const body = this.parseAlts({ ...f })
    this.close()
    return { t: 'look', behind, neg, body }
  }

  /** `(*VERB)`, `(*VERB:NAME)` and the alphabetic groups (`(*pla:…)`, `(*atomic:…)`). */
  starGroup(f: Flags): [Node, boolean] {
    const m = /^([a-z_]+):/.exec(this.p.slice(this.i))
    if (m) {
      const kind = ALPHA_GROUPS[m[1]]
      if (kind === undefined) this.fail()
      this.i += m[0].length
      if (kind === '>') return [this.inner(f, -1, true), true]
      return [this.look(f, kind.startsWith('<'), kind.endsWith('!')), true]
    }
    const v = /^([A-Z]*)(?::([^)]*))?\)/.exec(this.p.slice(this.i))
    if (!v || !(v[1] in VERBS)) this.fail()
    const k = VERBS[v[1]]
    const name = v[2] ?? null
    if ((k === 'mark' && !name) || ((k === 'accept' || k === 'fail') && name !== null)) this.fail()
    this.i += v[0].length
    return [{ t: 'verb', k, name }, false]
  }

  conditional(f: Flags): Node {
    let cond: Cond
    const rest = this.p.slice(this.i)
    if (rest.startsWith('?=') || rest.startsWith('?!') || rest.startsWith('?<=') || rest.startsWith('?<!') || rest.startsWith('*')) {
      this.i--
      const [look] = this.group(f)
      if (look.t !== 'look') this.fail()
      cond = { k: 'look', look }
    } else {
      let m: RegExpExecArray | null
      if ((m = /^R(\d+)\)/.exec(rest))) ((this.i += m[0].length), (cond = { k: 'R', ref: { groups: [Number(m[1])] } }))
      else if (this.eat('R&')) cond = { k: 'R', ref: this.ref(')', false) }
      else if (this.eat('R)')) cond = { k: 'R', ref: null }
      else if (this.eat('DEFINE)')) cond = { k: 'define' }
      else if ((m = /^VERSION(>?=)(\d+)(?:\.(\d+))?\)/.exec(rest))) {
        this.i += m[0].length
        const v = Number(m[2]) * 100 + Number((m[3] ?? '0').padEnd(2, '0').slice(0, 2))
        cond = { k: (m[1] === '=' ? v === 1045 : 1045 >= v) ? 'true' : 'false' }
      } else if (this.eat('<')) cond = { k: 'group', ref: this.ref('>)', false) }
      else if (this.eat("'")) cond = { k: 'group', ref: this.ref("')", false) }
      else cond = { k: 'group', ref: this.ref(')', true) }
    }
    const body = this.parseAlts({ ...f })
    this.close()
    const alts = body.t === 'alt' ? body.alts : [body]
    if (alts.length > 2 || (cond.k === 'define' && alts.length > 1)) this.fail()
    return { t: 'cond', cond, yes: alts[0], no: alts[1] ?? empty }
  }
}

const NO_FLAGS: Flags = { i: false, m: false, s: false, x: false, xx: false, n: false, U: false, J: false, r: false, aD: false, aS: false, aW: false, aP: false, aT: false }

/** Group numbers a node refers to, checked after the parse (`\2` without a second group is an error). */
function checkRefs(node: Node, groups: number, fail: () => never): void {
  const check = (r: Ref, zero: boolean) => r.groups.some((g) => g > groups || (g === 0 && !zero)) && fail()
  const walk = (n: Node): void => {
    switch (n.t) {
      case 'seq':
        return n.items.forEach(walk)
      case 'alt':
        return n.alts.forEach(walk)
      case 'group':
      case 'look':
      case 'rep':
        return walk(n.body)
      case 'bref':
        return void check(n.ref, false)
      case 'call':
        return void check(n.ref, true)
      case 'cond':
        if (n.cond.k === 'group') check(n.cond.ref, false)
        if (n.cond.k === 'look') walk(n.cond.look)
        walk(n.yes)
        return walk(n.no)
    }
  }
  walk(node)
}

/** The fewest and most characters a node matches (Infinity for no bound), groups looked up for calls and references. */
export function widths(parsed: Parsed): (n: Node) => [number, number] {
  const busy = new Set<number>()
  const group = (g: number): [number, number] => {
    if (busy.has(g)) return [0, Infinity]
    busy.add(g)
    try {
      return g === 0 ? w(parsed.root) : w(parsed.bodies[g])
    } finally {
      busy.delete(g)
    }
  }
  const w = (n: Node): [number, number] => {
    switch (n.t) {
      case 'char':
      case 'set':
      case 'any':
        return [1, 1]
      case 'nl':
        return [1, 2]
      case 'cluster':
        return [1, Infinity]
      case 'seq':
        return n.items.map(w).reduce(([a, b], [c, d]) => [a + c, b + d], [0, 0])
      case 'alt': {
        const all = n.alts.map(w)
        return [Math.min(...all.map((x) => x[0])), Math.max(...all.map((x) => x[1]))]
      }
      case 'group':
        return w(n.body)
      case 'rep': {
        const [a, b] = w(n.body)
        return [a * n.min, b === 0 ? 0 : b * n.max]
      }
      case 'bref': {
        const all = n.ref.groups.map(group)
        return [0, Math.max(...all.map((x) => x[1]))]
      }
      case 'call':
        return group(n.ref.groups[0])
      case 'cond': {
        if (n.cond.k === 'define') return [0, 0]
        const [y, z] = [w(n.yes), w(n.no)]
        return [Math.min(y[0], z[0]), Math.max(y[1], z[1])]
      }
      default:
        return [0, 0]
    }
  }
  return w
}

/** Every alternative of a lookbehind is of bounded length, 255 characters at most where it varies (PCRE2's default). */
function checkLookbehinds(parsed: Parsed, fail: () => never): void {
  const w = widths(parsed)
  const walk = (n: Node): void => {
    if (n.t === 'look' && n.behind)
      for (const a of n.body.t === 'alt' ? n.body.alts : [n.body]) {
        const [min, max] = w(a)
        if (max === Infinity || max > 65535 || (min !== max && max > 255)) fail()
      }
    if (n.t === 'seq') n.items.forEach(walk)
    else if (n.t === 'alt') n.alts.forEach(walk)
    else if (n.t === 'group' || n.t === 'look' || n.t === 'rep') walk(n.body)
    else if (n.t === 'cond') [...(n.cond.k === 'look' ? [n.cond.look] : []), n.yes, n.no].forEach(walk)
  }
  walk(parsed.root)
}

/** A pattern parsed, or a `RegexSyntaxError`. */
export function parse(pattern: string): Parsed {
  const p: Parser = new Parser(pattern)
  p.parseHead()
  const root = p.parseAlts({ ...NO_FLAGS })
  if (p.i < pattern.length) p.fail()
  for (const r of p.refs) {
    const g = p.names.get(r.name!)
    if (!g) p.fail()
    r.groups = g
  }
  checkRefs(root, p.groups, () => p.fail())
  const parsed = { root, groups: p.groups, bodies: p.bodies, head: p.head }
  checkLookbehinds(parsed, () => p.fail())
  return parsed
}
