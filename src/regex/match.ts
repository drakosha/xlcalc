/**
 * The matcher of a compiled pattern (`program.ts`): backtracking over the code points of a UTF-16 text with its own
 * stack, as PCRE2's interpreter does (`pcre2_match`), so that its work is counted. Each start position may take
 * Excel's match limit of backtracking points (`'limit'`, `#VALUE!` in Excel), and every 1,024 steps are spent on the
 * computation's budget, whose deadline and signal stop it (`exec/budget.ts`).
 */
import { spend } from '../exec/budget.js'
import { cpAt, cpBefore, Newline, newlineAt, newlineBefore, sameCase, wordTest } from './chars.js'
import {
  ACCEPT, ANY, ANYNL, ASSERT, ASSERTS, ATOME, ATOMS, BACK, BEHIND, BREF, CALL, CHAR, CHARI, CINC, CINIT, CLOSE, CLUSTER, CNT, COMMIT, COND, CONDG, CONDR,
  ECHK, ESTART, FAIL, GMARK, JMP, LOOKE, LOOKS, MARK, MATCH, NEG, NL, OPEN, PRUNE, PUSH, REP, SET, SKIP, SKIPN, THEN, VERB, type Program,
} from './program.js'

/**
 * Backtracking points one start position may take, PCRE2's own count, as Excel's: a pattern that needs more is
 * `#VALUE!`. Excel for Mac 16.114 (`xmatch-regex-syntax-2`): `(a+)+$` over 18 letters `a` and `b` stays under it and
 * over 19 not, `(a|aa)+$` over 25 and 26 the same, `(*LIMIT_MATCH=500000)` stops the 18 and 1,000,000 does not.
 */
export const MATCH_LIMIT = 1_000_000
/**
 * This matcher's backtracking points for one of PCRE2's: of the heavy patterns Excel stopped and did not
 * (`xmatch-regex-syntax-2`, `xmatch-regex-own-held`), the highest count it let through is 786,427 (`(a|a?)+$` over 17)
 * and the lowest it stopped 917,502 (`(a+)+$` over 19), so the limit lies at 0.85 of PCRE2's.
 */
const SCALE = 0.85

const F_CHOICE = 1, F_UNDO = 2, F_OBJ = 3, F_REPG = 4, F_REPL = 5, F_BACK = 6, F_ATOMIC = 7, F_LOOK = 8, F_CALL = 9, F_GMARK = 10, F_VERB = 11
const O_CALL = 0, O_MARKS = 1, O_SLOTS = 2

/** The outcome of one start position. */
const FAILED = 0, MATCHED = 1, NONE = 2, LIMITED = 3
/** Backtracking goes on (a verb that confines nothing more). */
const ON = -3

interface Call {
  g: number
  ret: number
  at: number
  saved: Int32Array
  prev: Call | null
}
interface Mark {
  name: number
  at: number
  prev: Mark | null
}

let segmenter: { segment(s: string): Iterable<{ segment: string }> } | null | undefined
/** The extended grapheme cluster at `i` (`\X`): by `Intl.Segmenter`, or a character and the marks after it. */
function clusterEnd(s: string, i: number): number {
  if (segmenter === undefined) segmenter = typeof Intl !== 'undefined' && 'Segmenter' in Intl ? new (Intl as unknown as { Segmenter: new (l: undefined, o: object) => typeof segmenter }).Segmenter(undefined, { granularity: 'grapheme' }) : null
  if (segmenter) for (const g of segmenter.segment(s.slice(i, i + 64))) return i + g.segment.length
  let j = i + (cpAt(s, i) > 0xffff ? 2 : 1)
  while (j < s.length && /\p{M}/u.test(String.fromCodePoint(cpAt(s, j)))) j += cpAt(s, j) > 0xffff ? 2 : 1
  return j
}

/** Runs a program over texts. */
export class Matcher {
  private s = ''
  private len = 0
  private slots: Int32Array
  private st = new Int32Array(256)
  private sp = 0
  private objs: unknown[] = []
  private call: Call | null = null
  private marks: Mark | null = null
  private points = 0
  private ticks = 0
  private start = 0
  private skipTo = -1
  private accepting = false
  private readonly limit: number
  private readonly word: (c: number) => boolean

  constructor(private readonly prog: Program) {
    this.slots = new Int32Array(prog.slots)
    this.limit = Math.min(MATCH_LIMIT, prog.head.limit) * SCALE
    this.word = wordTest({ aD: false, aS: false, aW: false, aP: false, aT: false })
  }

  /** Whether the pattern is found anywhere in `s`, or `'limit'` past the match limit at a start position. */
  test(s: string): boolean | 'limit' {
    const p = this.prog
    this.s = s
    this.len = s.length
    const { first, prefix, required, minLength } = p
    // a required character found nowhere: no match, without trying (as PCRE2 does; Excel, for 6,000 characters too)
    if (required && s.indexOf(required) < 0) return false
    for (let at = 0; at <= this.len; ) {
      if (this.len - at < minLength) return false
      if (p.anchored) {
        if (at > 0 || (prefix && !s.startsWith(prefix))) return false
      } else if (prefix) {
        at = s.indexOf(prefix, at)
        if (at < 0) return false
      } else if (first) {
        for (let c = cpAt(s, at); at < this.len && !first(c); c = cpAt(s, at)) at += c > 0xffff ? 2 : 1
        if (at >= this.len) return false
      }
      this.skipTo = -1
      const r = this.attempt(at)
      if (r === MATCHED) return true
      if (r === LIMITED) return 'limit'
      if (r === NONE || p.anchored) return false
      at = Math.max(at + (at < this.len && cpAt(s, at) > 0xffff ? 2 : 1), this.skipTo)
    }
    return false
  }

  private push(t: number, a: number, b: number, c: number) {
    if (this.sp + 4 > this.st.length) {
      const bigger = new Int32Array(this.st.length * 2)
      bigger.set(this.st)
      this.st = bigger
    }
    const st = this.st
    st[this.sp] = t
    st[this.sp + 1] = a
    st[this.sp + 2] = b
    st[this.sp + 3] = c
    this.sp += 4
  }

  private set(k: number, v: number) {
    // without a backtracking point below, nothing goes back to the old value
    if (this.sp > 0) this.push(F_UNDO, k, this.slots[k], 0)
    this.slots[k] = v
  }

  private keep(kind: number, o: unknown) {
    this.push(F_OBJ, kind, this.objs.push(o) - 1, 0)
  }

  private undo(at: number) {
    const st = this.st
    if (st[at] === F_UNDO) this.slots[st[at + 1]] = st[at + 2]
    else if (st[at] === F_OBJ) {
      const o = this.objs[st[at + 2]]
      if (st[at + 1] === O_CALL) this.call = o as Call | null
      else if (st[at + 1] === O_MARKS) this.marks = o as Mark | null
      else this.slots.set(o as Int32Array)
    }
  }

  /** Removes everything above `at`, undoing what it set. */
  private unwind(at: number) {
    while (this.sp > at) {
      this.sp -= 4
      this.undo(this.sp)
    }
  }

  /** Drops the frame at `at` and the backtracking points above it, keeping what undoes the registers. */
  private cut(at: number) {
    const st = this.st
    let w = at
    for (let r = at + 4; r < this.sp; r += 4) {
      if (st[r] !== F_UNDO && st[r] !== F_OBJ) continue
      st.copyWithin(w, r, r + 4)
      w += 4
    }
    this.sp = w
  }

  /** The nearest frame of type `t` below the top. */
  private find(t: number): number {
    for (let at = this.sp - 4; at >= 0; at -= 4) if (this.st[at] === t) return at
    return -1
  }

  /** `n` code points back from `i`, or -1. */
  private back(i: number, n: number): number {
    for (; n > 0; n--) {
      if (i <= 0) return -1
      i -= cpBefore(this.s, i) > 0xffff ? 2 : 1
    }
    return i
  }

  /** Where an item of a repeat (CHAR, CHARI, ANY, ANYNL, SET) at `i` ends, or -1. */
  private item(op: number, a: number, b: number, i: number): number {
    if (i >= this.len) return -1
    const c = cpAt(this.s, i)
    const next = i + (c > 0xffff ? 2 : 1)
    switch (op) {
      case CHAR:
        return c === a ? next : -1
      case CHARI:
        return sameCase(c, a, b === 1) ? next : -1
      case ANY:
        return next
      case ANYNL:
        return newlineAt(this.s, i, this.prog.head.newline) ? -1 : next
      default:
        return this.prog.tests[a](c) ? next : -1
    }
  }

  private assert(k: string, i: number): boolean {
    const { s, len } = this
    const nl = this.prog.head.newline
    switch (k) {
      case 'bol':
      case 'A':
        return i === 0
      case 'G':
        return i === 0
      case 'mbol':
        // after a newline at the end too, as Excel's `(?m)^$` finds one in `a` and a line feed
        return i === 0 || newlineBefore(s, i, nl)
      case 'z':
        return i === len
      case 'eol':
      case 'Z':
        return i === len || newlineAt(s, i, nl) === len - i
      case 'meol':
        return i === len || newlineAt(s, i, nl) > 0
    }
    const ascii = k.startsWith('a')
    const w = (c: number) => (ascii ? c < 128 && this.word(c) : this.word(c))
    const before = i > 0 && w(cpBefore(s, i))
    const after = i < len && w(cpAt(s, i))
    return (before !== after) === (k === 'b' || k === 'ab')
  }

  /** Returns from the innermost call: the registers as they were before it, the position as it is. */
  private ret() {
    const fr = this.call!
    this.keep(O_SLOTS, this.slots.slice())
    this.slots.set(fr.saved)
    this.keep(O_CALL, fr)
    this.call = fr.prev
    return fr.ret
  }

  /** One start position. */
  private attempt(start: number): number {
    const p = this.prog
    const code = p.code
    const s = this.s
    const len = this.len
    const slots = this.slots
    for (let k = 0; k < slots.length; k++) slots[k] = -1
    this.sp = 0
    if (this.objs.length > 0) this.objs = []
    this.call = null
    this.marks = null
    this.points = 0
    this.start = start
    let pc = 0
    let i = start
    for (;;) {
      if (++this.ticks >= 1024) {
        this.ticks = 0
        spend(1)
      }
      let ok = true
      switch (code[pc]) {
        case CHAR: {
          const c = code[pc + 1]
          if (c < 0x10000 ? i < len && s.charCodeAt(i) === c && !((c & 0xfc00) === 0xd800) : i < len && cpAt(s, i) === c) {
            i += c > 0xffff ? 2 : 1
            pc += 2
          } else ok = false
          break
        }
        case CHARI:
        case ANY:
        case ANYNL:
        case SET: {
          const op = code[pc]
          const j = this.item(op, code[pc + 1], code[pc + 2], i)
          if (j < 0) ok = false
          else {
            i = j
            pc += op === CHARI ? 3 : op === SET ? 2 : 1
          }
          break
        }
        case REP: {
          const op = code[pc + 1]
          const a = code[pc + 2]
          const b = code[pc + 3]
          const min = code[pc + 4]
          const max = code[pc + 5]
          const mode = code[pc + 6]
          const most = max < 0 ? Infinity : max
          let n = 0
          let j = i
          let minAt = i
          const upto = mode === 1 ? min : most
          while (n < upto) {
            const k = this.item(op, a, b, j)
            if (k < 0) break
            j = k
            if (++n === min) minAt = j
          }
          this.ticks += n
          if (n < min) {
            ok = false
            break
          }
          if (min === 0) minAt = i
          if (mode === 0 && j > minAt) {
            if (++this.points > this.limit) return LIMITED
            this.push(F_REPG, pc + 7, j, minAt)
          } else if (mode === 1 && n < most) {
            if (++this.points > this.limit) return LIMITED
            this.push(F_REPL, pc, j, n)
          }
          i = j
          pc += 7
          break
        }
        case NL: {
          const c = i < len ? s.charCodeAt(i) : -1
          if (c === 13 && s.charCodeAt(i + 1) === 10) i += 2
          else if (c === 10 || c === 13 || (p.head.anyR && (c === 11 || c === 12 || c === 0x85 || c === 0x2028 || c === 0x2029))) i++
          else {
            ok = false
            break
          }
          pc++
          break
        }
        case CLUSTER:
          if (i >= len) ok = false
          else ((i = clusterEnd(s, i)), pc++)
          break
        case JMP:
          pc = code[pc + 1]
          break
        case PUSH:
          if (++this.points > this.limit) return LIMITED
          this.push(F_CHOICE, code[pc + 1], i, code[pc + 2])
          pc += 3
          break
        case GMARK:
          this.push(F_GMARK, code[pc + 1], 0, 0)
          pc += 2
          break
        case OPEN:
          this.set(3 * code[pc + 1], i)
          pc += 2
          break
        case CLOSE: {
          const g = code[pc + 1]
          this.set(3 * g + 1, this.slots[3 * g])
          this.set(3 * g + 2, i)
          pc = this.call && this.call.g === g ? this.ret() : pc + 2
          break
        }
        case BREF: {
          const groups = p.refs[code[pc + 1]]
          const g = groups.find((x) => this.slots[3 * x + 2] >= 0)
          if (g === undefined) {
            ok = false
            break
          }
          const from = this.slots[3 * g + 1]
          const to = this.slots[3 * g + 2]
          let j = i
          if (code[pc + 2]) {
            for (let k = from; k < to && ok; ) {
              const c = cpAt(s, k)
              if (j >= len || !sameCase(cpAt(s, j), c, code[pc + 3] === 1)) ok = false
              else {
                j += cpAt(s, j) > 0xffff ? 2 : 1
                k += c > 0xffff ? 2 : 1
              }
            }
          } else if (s.startsWith(s.slice(from, to), i)) j = i + to - from
          else ok = false
          if (ok) ((i = j), (pc += 4))
          break
        }
        case ASSERT:
          if (this.assert(ASSERTS[code[pc + 1]], i)) pc += 2
          else ok = false
          break
        case LOOKS:
          if (++this.points > this.limit) return LIMITED
          this.push(F_LOOK, code[pc + 1], i, code[pc + 2])
          pc += 3
          break
        case LOOKE: {
          const kind = code[pc + 1]
          const at = this.find(F_LOOK)
          const from = this.st[at + 2]
          if (kind & BEHIND && !this.accepting && i !== from) {
            ok = false
            break
          }
          this.accepting = false
          if (kind & NEG) {
            this.unwind(at)
            if (!(kind & COND)) {
              ok = false
              break
            }
            pc = code[pc + 2]
          } else {
            this.cut(at)
            pc += 3
          }
          i = from
          break
        }
        case BACK: {
          const min = code[pc + 1]
          const max = code[pc + 2]
          let k = max
          let j = this.back(i, k)
          while (j < 0 && k > min) j = this.back(i, --k)
          if (j < 0) {
            ok = false
            break
          }
          if (k > min) this.push(F_BACK, pc, i, k - 1)
          i = j
          pc += 3
          break
        }
        case ATOMS:
          if (++this.points > this.limit) return LIMITED
          this.push(F_ATOMIC, 0, 0, 0)
          pc++
          break
        case ATOME:
          this.cut(this.find(F_ATOMIC))
          pc++
          break
        case CINIT:
          this.set(code[pc + 1], 0)
          pc += 2
          break
        case CNT: {
          const r = code[pc + 1]
          const min = code[pc + 2]
          const max = code[pc + 3]
          const lazy = code[pc + 4]
          const exit = code[pc + 5]
          const c = this.slots[r]
          if (c < min) pc += 6
          else if (max >= 0 && c >= max) pc = exit
          else {
            if (++this.points > this.limit) return LIMITED
            if (lazy) {
              this.push(F_CHOICE, pc + 6, i, -1)
              pc = exit
            } else {
              this.push(F_CHOICE, exit, i, -1)
              pc += 6
            }
          }
          break
        }
        case CINC:
          this.set(code[pc + 1], this.slots[code[pc + 1]] + 1)
          pc += 2
          break
        case ESTART:
          this.set(code[pc + 1], i)
          pc += 2
          break
        case ECHK: {
          // an iteration that matched nothing ends the loop (once its minimum is reached)
          const r = code[pc + 2]
          pc = i === this.slots[code[pc + 1]] && (r < 0 || this.slots[r] >= code[pc + 3]) ? code[pc + 4] : pc + 5
          break
        }
        case CALL: {
          const g = code[pc + 1]
          for (let c = this.call; c; c = c.prev) if (c.g === g && c.at === i) return LIMITED
          if (++this.points > this.limit) return LIMITED
          this.push(F_CALL, 0, 0, 0)
          this.keep(O_CALL, this.call)
          this.call = { g, ret: pc + 2, at: i, saved: this.slots.slice(), prev: this.call }
          pc = p.groupPc[g]
          break
        }
        case CONDG:
          pc = p.refs[code[pc + 1]].some((g) => this.slots[3 * g + 2] >= 0) ? pc + 3 : code[pc + 2]
          break
        case CONDR: {
          const g = code[pc + 1]
          pc = this.call && (g < 0 || this.call.g === g) ? pc + 3 : code[pc + 2]
          break
        }
        case VERB: {
          const kind = code[pc + 1]
          const arg = code[pc + 2]
          if (kind === MARK) {
            this.keep(O_MARKS, this.marks)
            this.marks = { name: arg, at: i, prev: this.marks }
          } else this.push(F_VERB, kind, i, arg)
          pc += 4
          break
        }
        case ACCEPT:
          if (code[pc + 1] >= 0) {
            this.accepting = true
            pc = code[pc + 1]
          } else if (this.call) pc = this.ret()
          else pc = this.matchEnd(i) ? -1 : -2
          if (pc === -1) return MATCHED
          if (pc === -2) ok = false
          break
        case FAIL:
          ok = false
          break
        case MATCH:
          if (this.matchEnd(i)) return MATCHED
          ok = false
          break
      }
      if (ok) continue
      const r = this.backtrack()
      if (r >= 0) return r
      pc = this.pc
      i = this.pos
    }
  }

  private matchEnd(i: number): boolean {
    const h = this.prog.head
    return !(i === this.start && (h.notEmpty || (h.notEmptyAtStart && this.start === 0)))
  }

  private pc = 0
  private pos = 0

  /** Goes back to the last backtracking point (`pc`, `pos`): -1, or the start position's outcome when none is left. */
  private backtrack(): number {
    const st = this.st
    const code = this.prog.code
    while (this.sp > 0) {
      this.sp -= 4
      const at = this.sp
      const t = st[at]
      if (t === F_UNDO || t === F_OBJ) {
        this.undo(at)
        continue
      }
      if (t === F_CHOICE) return this.resume(st[at + 1], st[at + 2])
      if (t === F_REPG) {
        const next = st[at + 1]
        const end = st[at + 2]
        const minAt = st[at + 3]
        let j = this.back(end, 1)
        // before a literal, only where it follows
        if (code[next] === CHAR && code[next + 1] < 0x10000) while (j > minAt && this.s.charCodeAt(j) !== code[next + 1]) j = this.back(j, 1)
        if (j < minAt) continue
        if (j > minAt) {
          st[at + 2] = j
          this.sp += 4
        }
        if (++this.points > this.limit) return LIMITED
        return this.resume(next, j)
      }
      if (t === F_REPL) {
        const rep = st[at + 1]
        const i = st[at + 2]
        const n = st[at + 3]
        const j = this.item(code[rep + 1], code[rep + 2], code[rep + 3], i)
        if (j < 0) continue
        const max = code[rep + 5]
        if (max < 0 || n + 1 < max) {
          st[at + 2] = j
          st[at + 3] = n + 1
          this.sp += 4
        }
        if (++this.points > this.limit) return LIMITED
        return this.resume(rep + 7, j)
      }
      if (t === F_BACK) {
        const back = st[at + 1]
        const from = st[at + 2]
        const k = st[at + 3]
        if (k > code[back + 1]) {
          st[at + 3] = k - 1
          this.sp += 4
        }
        const j = this.back(from, k)
        if (j < 0) continue
        return this.resume(back + 3, j)
      }
      if (t === F_LOOK) {
        const kind = st[at + 1]
        // a positive assertion that fails fails; a negative one holds; a condition takes its other branch
        if (!(kind & NEG) && !(kind & COND)) continue
        return this.resume(st[at + 3], st[at + 2])
      }
      if (t === F_VERB) {
        const r = this.verb(at)
        if (r !== ON) return r
      }
    }
    return FAILED
  }

  private resume(pc: number, pos: number): number {
    this.pc = pc
    this.pos = pos
    return -1
  }

  /** Backtracking into a verb at `at` (already off the stack): -1 resumed, ON to go on backtracking, else the start position's outcome. */
  private verb(at: number): number {
    const st = this.st
    const kind = st[at + 1]
    const pos = st[at + 2]
    const arg = st[at + 3]
    if (kind === THEN && arg >= 0) {
      // to the next alternative of the innermost group with alternatives around it
      while (this.sp > 0) {
        this.sp -= 4
        const t = st[this.sp]
        if (t === F_CHOICE && st[this.sp + 3] === arg) return this.resume(st[this.sp + 1], st[this.sp + 2])
        if (t === F_GMARK && st[this.sp + 1] === arg) return ON
        this.undo(this.sp)
      }
      return FAILED
    }
    let skip = pos
    if (kind === SKIPN) {
      let m = this.marks
      while (m && m.name !== arg) m = m.prev
      if (!m) return ON
      skip = m.at
    }
    // a negative assertion, a condition or a call around it confines it: that one fails
    for (let b = this.sp - 4; b >= 0; b -= 4) {
      const t = st[b]
      if (t === F_CALL || (t === F_LOOK && st[b + 1] & (NEG | COND))) {
        this.unwind(b + 4)
        return ON
      }
    }
    if (kind === COMMIT) return NONE
    if (kind === SKIP || kind === SKIPN) this.skipTo = skip > this.start ? skip : -1
    void PRUNE
    return FAILED
  }
}
