/**
 * Match mode 3's patterns without the PCRE2 engine (`xlcalc/regex`): a PCRE2 pattern (UTF and UCP, newlines CR, LF and
 * CRLF, as Excel for Mac takes it) written as a JavaScript regular expression with the `v` flag. Unicode `\d \w \s \b`,
 * `\A \z \Z \G`, the dot and anchors by PCRE2's newlines, `\Q…\E`, possessive quantifiers and atomic groups (by a
 * lookahead and a back reference), `{,m}`, option settings (`(?i)` inside a pattern by JavaScript's modifiers), named
 * groups in each spelling, `(?|…)`, alphabetic assertions and `(*FAIL)` are translated. Recursion, conditional groups,
 * the verbs but `(*FAIL)` and `(*MARK)`, callouts, `\C`, `\X`, `(*NOTEMPTY)` and a lookbehind of unbounded length are
 * not: null, as is a pattern that is none.
 */
import { boundary, braces, END, START, textEnd, typeClass } from '../functions/lookup/regex-syntax.js'
import { ASCII_POSIX, H, POSIX, property, V, W } from './classes.js'

const hex = (c: number) => `\\u{${c.toString(16)}}`
/** A literal character: letters and digits as they are, any other by its code. */
const lit = (c: number) => (/[\p{L}\p{N}]/u.test(String.fromCodePoint(c)) ? String.fromCodePoint(c) : hex(c))

const NEWLINES: Record<string, string> = { CR: '\\r', LF: '\\n', CRLF: '\\n\\r', ANYCRLF: '\\n\\r', ANY: '\\n-\\r\\x85\\u2028\\u2029', NUL: '\\0' }
const LOOKS: Record<string, string> = { pla: '(?=', positive_lookahead: '(?=', napla: '(?=', non_atomic_positive_lookahead: '(?=', nla: '(?!', negative_lookahead: '(?!', plb: '(?<=', positive_lookbehind: '(?<=', naplb: '(?<=', non_atomic_positive_lookbehind: '(?<=', nlb: '(?<!', negative_lookbehind: '(?<!', atomic: '>' }

interface Flags {
  i: boolean
  m: boolean
  s: boolean
  x: boolean
  xx: boolean
  n: boolean
  U: boolean
  J: boolean
  aD: boolean
  aS: boolean
  aW: boolean
  aP: boolean
}

interface Level {
  f: Flags
  /** What closes the group in JavaScript. */
  close: string
  /** Of `(?|…)`: the group number its alternatives start from, and the highest they reached. */
  reset: number
  most: number
  behind: boolean
}

class Untranslatable extends Error {}
const no = (): never => {
  throw new Untranslatable()
}
/** A pattern that leaves its leading `(?i)`: translated again with the caseless parts spelled out. */
class Restart extends Error {}

/** The characters of `c`'s case (lower, upper, their round trips), one code point each. */
function variants(c: number): number[] {
  const s = String.fromCodePoint(c)
  const all = [s, s.toLowerCase(), s.toUpperCase(), s.toUpperCase().toLowerCase(), s.toLowerCase().toUpperCase()].filter((t) => [...t].length === 1)
  return [...new Set(all.map((t) => t.codePointAt(0)!))]
}

/** A PCRE2 pattern as a JavaScript regular expression, or null where it has no translation or is no pattern. */
export function pcreToRegExp(pattern: string): RegExp | null {
  for (const whole of [true, false]) {
    try {
      return translate(pattern, whole)
    } catch (e) {
      if (e instanceof Restart) continue
      if (e instanceof Untranslatable || e instanceof SyntaxError) return null
      throw e
    }
  }
  return null
}

/**
 * The translation; with `whole`, a leading `(?i)` makes the regular expression caseless (`i`), else the caseless parts
 * are classes of each character's case, as JavaScript's modifiers (`(?i:…)`) are missing before Node 23.
 */
function translate(p: string, whole: boolean): RegExp {
  let i = 0
  let nl = '\\n\\r'
  let anyR = true
  for (let m; (m = /^\(\*([A-Z_]+)(?:=\d+)?\)/.exec(p.slice(i))); i += m[0].length) {
    const v = m[1]
    if (v in NEWLINES) nl = NEWLINES[v]
    else if (v.startsWith('BSR_')) anyR = v === 'BSR_UNICODE'
    else if (!/^(UTF|UCP|NO_AUTO_POSSESS|NO_DOTSTAR_ANCHOR|NO_JIT|NO_START_OPT|LIMIT_[A-Z]+)$/.test(v)) break
  }
  const lead = p.startsWith('(?i)', i)
  if (lead) i += 4
  const base = whole && lead
  const f0: Flags = { i: lead, m: false, s: false, x: false, xx: false, n: false, U: false, J: false, aD: false, aS: false, aW: false, aP: false }
  const stack: Level[] = [{ f: f0, close: '', reset: -1, most: 0, behind: false }]
  const top = () => stack[stack.length - 1]
  let out = ''
  let groups = 0
  let atoms = 0
  const names: string[] = []
  const refers = /\\[1-9gk]|\(\?P=/.test(p)
  const used = new Set<string>()
  /** Where the last item that a quantifier may follow starts in `out`, -1 when none. */
  let atom = -1
  const eol = (multi: boolean) => textEnd(nl, multi)
  const word = () => (top().f.aW ? '[0-9A-Za-z_]' : W)

  /** A class escape (`d`, `W`, `p`…) as a JavaScript class, or null for another letter. */
  const typed = (c: string): string | null => {
    const f = top().f
    const lower = c.toLowerCase()
    let t: string | null = null
    if (lower === 'd' || lower === 's' || lower === 'w') return typeClass(c, { d: f.aD, s: f.aS, w: f.aW })
    if (lower === 'h') t = H
    else if (lower === 'v') t = V
    else if (lower === 'p') {
      let name = p[i++] ?? no()
      if (name === '{') {
        const end = p.indexOf('}', i)
        if (end < 0) no()
        name = p.slice(i, end)
        i = end + 1
      }
      const neg = c === 'P' !== name.startsWith('^')
      return property(name.replace(/^\^/, ''), neg, f.i) ?? no()
    } else return null
    return c === lower ? t : `[^${t}]`
  }

  /** A character escape's code point (`\x{41}`, `\n`, `\.`), the letter at `i - 1` already read. */
  const charEscape = (c: string): number => {
    const take = (re: RegExp, base: number) => {
      const m = re.exec(p.slice(i)) ?? no()
      i += m[0].length
      const v = m[1] === '' ? 0 : parseInt(m[1], base)
      return v > 0x10ffff || (v >= 0xd800 && v <= 0xdfff) ? no() : v
    }
    const simple = 'a\x07e\x1bf\fn\nr\rt\t'.indexOf(c)
    if (simple >= 0 && simple % 2 === 0) return 'a\x07e\x1bf\fn\nr\rt\t'.charCodeAt(simple + 1)
    if (c === '0') return (i--, take(/^([0-7]{1,3})/, 8))
    if (c === 'o') return take(/^\{([0-7]+)\}/, 8)
    if (c === 'x') return p[i] === '{' ? take(/^\{\s*([0-9a-fA-F]+)\s*\}/, 16) : take(/^([0-9a-fA-F]{0,2})/, 16)
    if (c === 'N') return take(/^\{U\+([0-9a-fA-F]+)\}/, 16)
    if (c === 'c') {
      const x = p.charCodeAt(i++)
      return x >= 32 && x < 127 ? (x >= 97 && x <= 122 ? x - 32 : x) ^ 0x40 : no()
    }
    if (/[\p{L}\p{N}]/u.test(c)) no()
    return c.codePointAt(0)!
  }

  const cp = () => {
    const c = p.codePointAt(i) ?? no()
    i += c > 0xffff ? 2 : 1
    return c
  }
  const name = (close: string) => {
    const m = /^[_\p{L}][_\p{L}\p{N}]*/u.exec(p.slice(i))
    if (!m || !p.startsWith(close, i + m[0].length)) no()
    i += m![0].length + close.length
    return m![0]
  }
  /** A back reference to a group by number (its JavaScript name is known at the end) or by name. */
  const ref = (n: number | string) => {
    if ((typeof n === 'number' && n <= 0) || (top().f.i && !base)) no()
    atom = out.length
    out += `\\k<\u0001${n}\u0001>`
  }
  const relative = (close: string) => {
    const m = new RegExp(`^([+-]?)(\\d+)${close.replace(/[{}]/g, '\\$&')}`).exec(p.slice(i))
    if (!m) return ref(name(close))
    i += m[0].length
    ref(m[1] === '-' ? groups - Number(m[2]) + 1 : m[1] === '+' ? no() : Number(m[2]))
  }

  /** A literal character, caseless by its case's characters where JavaScript's `i` is not on. */
  const chr = (c: number) => (top().f.i && !base ? `[${variants(c).map(hex).join('')}]` : lit(c))
  const options = (f: Flags) => {
    let on = true
    if (p[i] === '^') ((i++), Object.assign(f, { i: false, m: false, n: false, s: false, x: false, xx: false }))
    for (let c; (c = p[i]) !== ')' && c !== ':'; i++) {
      if (c === '-' && on) on = false
      else if (c === 'x' && p[i + 1] === 'x') ((i++), (f.x = f.xx = on))
      else if (c === 'x') ((f.x = on), on || (f.xx = false))
      else if (c === 'a') {
        const k = /[DSWPT]/.test(p[i + 1] ?? '') ? p[++i] : 'DSWP'
        for (const key of k === 'T' ? 'P' : k) f[`a${key}` as 'aD'] = on
      } else if (c !== undefined && 'imnsUJr'.includes(c)) {
        if (c !== 'r') f[c as 'i'] = on
      } else no()
    }
  }
  const open = (js: string, close: string, f = { ...top().f }, behind = top().behind) => {
    atom = -1
    out += js
    stack.push({ f, close, reset: -1, most: 0, behind })
  }

  const charClass = () => {
    let s = p[i] === '^' ? ((i++), '[^') : '['
    const f = top().f
    // caseless without JavaScript's `i`: the other cases of the characters and ranges, added at the end
    const extra = f.i && !base ? new Set<number>() : null
    let lo = -1
    let first = true
    for (;;) {
      if (i >= p.length) no()
      if (f.xx && (p[i] === ' ' || p[i] === '\t')) {
        i++
        continue
      }
      if (p[i] === ']' && !first) break
      first = false
      if (p.startsWith('\\Q', i)) {
        const end = p.indexOf('\\E', i)
        for (const ch of p.slice(i + 2, end < 0 ? p.length : end)) s += hex(ch.codePointAt(0)!)
        i = end < 0 ? p.length : end + 2
        continue
      }
      if (p.startsWith('\\E', i)) {
        i += 2
        continue
      }
      const posix = /^\[:(\^?)([a-z]+):\]/.exec(p.slice(i))
      let c: number
      if (posix) {
        const t = f.i && /^(lower|upper)$/.test(posix[2]) && !f.aP ? '\\p{LC}' : (f.aP && ASCII_POSIX[posix[2]]) || POSIX[posix[2]] || no()
        i += posix[0].length
        s += posix[1] ? `[^${t}]` : t
        if (p[i] === '-' && p[i + 1] !== ']') no()
        continue
      } else if (p[i] === '\\') {
        i++
        const e = p[i++] ?? no()
        const t = e === 'b' ? null : typed(e)
        if (t) {
          s += t
          if (p[i] === '-' && p[i + 1] !== ']') no()
          continue
        }
        if (e === 'b') c = 8
        else if (/[0-7]/.test(e)) {
          const o = /^[0-7]{1,3}/.exec(p.slice(--i))![0]
          i += o.length
          c = parseInt(o, 8)
        } else if ('RXBgkKAzZGC89'.includes(e) || (e === 'N' && p[i] !== '{')) c = no()
        else c = charEscape(e)
      } else c = cp()
      s += hex(c)
      if (extra) {
        if (lo >= 0 && c - lo > 5000) no()
        for (let k = lo >= 0 ? lo : c; k <= c; k++) for (const v of variants(k)) extra.add(v)
      }
      lo = -1
      if (p[i] === '-' && p[i + 1] !== ']' && p[i + 1] !== undefined && !/^\[:/.test(p.slice(i + 1))) {
        i++
        s += '-'
        lo = c
      }
    }
    i++
    return s + [...(extra ?? [])].map(hex).join('') + ']'
  }

  const quantifier = (): string | null => {
    const c = p[i]
    if (c === '*' || c === '+' || c === '?') return (i++, c)
    const q = c === '{' ? braces(p, i) : null
    if (!q) return null
    i += q[1]
    if (q[0].split(/[{,}]/).some((n) => Number(n) > 65535)) no()
    return q[0]
  }

  while (i < p.length) {
    const l = top()
    const f = l.f
    const c = p[i]
    if (f.x && /\s/.test(c)) {
      i++
      continue
    }
    if (f.x && c === '#') {
      while (i < p.length && p[i] !== '\n') i++
      continue
    }
    const q = quantifier()
    if (q !== null) {
      if (atom < 0) no()
      // in a lookbehind a quantifier is bounded, and one of variable length takes 255 at most
      const bounds = q.slice(1, -1).split(',')
      const [min, max] = q.length === 1 ? [0, q === '?' ? 1 : Infinity] : [Number(bounds[0]), bounds.length === 1 ? Number(bounds[0]) : bounds[1] === '' ? Infinity : Number(bounds[1])]
      if (l.behind && (max === Infinity || (min !== max && max > 255))) no()
      let mode = p[i] === '+' ? ((i++), 2) : p[i] === '?' ? ((i++), 1) : 0
      if (f.U && mode < 2) mode = 1 - mode
      if (mode === 2) {
        const n = `_a${atoms++}`
        out = `${out.slice(0, atom)}(?:(?=(?<${n}>${out.slice(atom)}${q}))\\k<${n}>)`
      } else out += q + (mode ? '?' : '')
      // what follows a quantifier takes none
      atom = -1
      continue
    }
    if (c === '|') {
      i++
      out += '|'
      if (l.reset >= 0) ((l.most = Math.max(l.most, groups)), (groups = l.reset))
      atom = -1
      continue
    }
    if (c === ')') {
      if (stack.length === 1) no()
      i++
      if (l.reset >= 0) groups = Math.max(l.most, groups)
      stack.pop()
      const start = out.lastIndexOf('\u0002')
      out = out.slice(0, start) + out.slice(start + 1) + l.close
      atom = start
      continue
    }
    const at = out.length
    i++
    if (c === '\\') {
      const e = p[i++] ?? no()
      if (e === 'Q') {
        const end = p.indexOf('\\E', i)
        for (const ch of p.slice(i, end < 0 ? p.length : end)) ((atom = out.length), (out += chr(ch.codePointAt(0)!)))
        i = end < 0 ? p.length : end + 2
        continue
      }
      atom = at
      if (/[1-9]/.test(e)) {
        const n = /^\d+/.exec(p.slice(i - 1))![0]
        if (Number(n) < 10 || e === '8' || e === '9' || Number(n) <= groups) {
          i += n.length - 1
          ref(Number(n))
          continue
        }
        i--
        const o = /^[0-7]{1,3}/.exec(p.slice(i))![0]
        i += o.length
        out += hex(parseInt(o, 8))
        continue
      }
      if (e === 'g') {
        if (p[i] === '<' || p[i] === "'") no()
        if (p[i] === '{') ((i++), relative('}'))
        else relative('')
        continue
      }
      if (e === 'k') {
        const close = { '<': '>', "'": "'", '{': '}' }[p[i++]] ?? no()
        ref(name(close))
        continue
      }
      const t = typed(e)
      if (t) {
        out += t
        continue
      }
      atom = -1
      const w = word()
      if (e === 'b' || e === 'B') out += boundary(w, e === 'B')
      else if (e === 'A' || e === 'G') out += START
      else if (e === 'z') out += END
      else if (e === 'Z') out += eol(false)
      else if (e === 'K' || e === 'E') out += ''
      else if (e === 'R') ((atom = at), (out += `(?:\\r\\n|${anyR ? V : '[\\n\\r]'})`))
      else if (e === 'N' && p[i] !== '{') ((atom = at), (out += `[^${nl}]`))
      else ((atom = at), (out += chr(charEscape(e))))
      continue
    }
    if (c === '[') {
      atom = at
      out += charClass()
      continue
    }
    if (c === '.') {
      atom = at
      out += f.s ? '[\\s\\S]' : `[^${nl}]`
      continue
    }
    if (c === '^') {
      atom = -1
      // after a newline at the end too, as in Excel
      out += f.m ? `(?:(?<![\\s\\S])|(?<=[${nl}])(?<!\\r(?=\\n)))` : '(?<![\\s\\S])'
      continue
    }
    if (c === '$') {
      atom = -1
      out += eol(f.m)
      continue
    }
    if (c !== '(') {
      i--
      atom = at
      out += chr(cp())
      continue
    }
    // a group: '\u0002' marks where it starts in `out`, for its quantifier
    out += '\u0002'
    if (p.startsWith('*', i)) {
      const m = /^\*([a-z_]+):/.exec(p.slice(i))
      const verb = /^\*(FAIL|F|MARK:[^)]+|:[^)]+)\)/.exec(p.slice(i))
      if (m && LOOKS[m[1]]) {
        i += m[0].length
        const js = LOOKS[m[1]]
        if (js === '>') atomic()
        else open(js, ')', undefined, js.startsWith('(?<'))
      } else if (verb) {
        i += verb[0].length
        out = out.slice(0, -1) + (verb[1].startsWith('F') ? '(?!)' : '')
        atom = -1
      } else no()
      continue
    }
    if (p[i] !== '?') {
      if (f.n) open('(?:', ')')
      else capture(null)
      continue
    }
    i++
    const r = p.slice(i)
    if (r.startsWith('#')) {
      const end = p.indexOf(')', i)
      if (end < 0) no()
      i = end + 1
      out = out.slice(0, -1)
      atom = -1
    } else if (r.startsWith(':')) ((i++), open('(?:', ')'))
    else if (r.startsWith('|')) {
      i++
      open('(?:', ')')
      top().reset = groups
    } else if (r.startsWith('>')) ((i++), atomic())
    else if (/^[=!]/.test(r)) open(`(?${p[i++]}`, ')')
    else if (/^<[=!]/.test(r)) ((i += 2), open(`(?<${p[i - 1]}`, ')', undefined, true))
    else if (/^(P?<|')[_\p{L}]/u.test(r)) {
      i += r[0] === 'P' ? 2 : 1
      capture(name(r[0] === "'" ? "'" : '>'))
    } else if (r.startsWith('P=')) {
      i += 2
      out = out.slice(0, -1)
      ref(name(')'))
    } else {
      // option settings: (?i) for the rest of the group, (?i:…) for its own
      const g = { ...f }
      options(g)
      if (p[i++] === ')') {
        out = out.slice(0, -1)
        Object.assign(f, g)
        atom = -1
      } else open('(?:', ')', g)
      if (base && !g.i) throw new Restart()
    }
  }

  function atomic() {
    const n = `_a${atoms++}`
    open(`(?:(?=(?<${n}>`, `))\\k<${n}>)`)
  }
  function capture(n: string | null) {
    const k = ++groups
    // a number again in (?|…): one name, which only newer engines take twice, where a back reference needs it
    const js = n ?? (names[k] && !refers ? `_${k}_${atoms++}` : `_${k}`)
    if (n !== null) {
      if (used.has(n) && !top().f.J) no()
      used.add(n)
    }
    names[k] ??= js
    open(`(?<${js}>`, ')')
  }

  if (stack.length > 1) no()
  const js = out.replace(/\\k<\u0001([^\u0001]+)\u0001>/g, (_, r: string) => `\\k<${/^\d+$/.test(r) ? (names[Number(r)] ?? no()) : r}>`)
  return new RegExp(js, base ? 'iv' : 'v')
}
