/**
 * A parsed pattern made into the program the matcher runs (`match.ts`): instructions in an Int32Array, character tests
 * and reference lists beside them, registers for captures, counters and empty iterations, and what a start
 * position needs (anchoring, the first character, a literal prefix, a required character, a minimum length), as PCRE2's
 * own start-of-match optimizations skip positions without work.
 */
import { type CharTest, remembered, sameCase } from './chars.js'
import { type Head, type Node, type Parsed, widths } from './parse.js'

/** Instructions; the operands follow each one (`OPERANDS`). */
export const CHAR = 0, CHARI = 1, ANY = 2, ANYNL = 3, SET = 4, REP = 5, NL = 6, CLUSTER = 7, JMP = 8, PUSH = 9, GMARK = 10,
  OPEN = 11, CLOSE = 12, BREF = 13, ASSERT = 14, LOOKS = 15, LOOKE = 16, BACK = 17, ATOMS = 18, ATOME = 19, CINIT = 20,
  CNT = 21, CINC = 22, ESTART = 23, ECHK = 24, CALL = 25, CONDG = 26, CONDR = 27, VERB = 28, ACCEPT = 29, FAIL = 30, MATCH = 31
/** Operands of each instruction. */
export const OPERANDS = [1, 2, 0, 0, 1, 6, 0, 0, 1, 2, 1, 1, 1, 3, 1, 2, 2, 2, 0, 0, 1, 5, 1, 1, 4, 1, 2, 2, 3, 1, 0, 0]

/** Assertions by number (ASSERT's operand). */
export const ASSERTS = ['bol', 'mbol', 'eol', 'meol', 'A', 'z', 'Z', 'G', 'b', 'B', 'ab', 'aB'] as const
/** Verbs by number (VERB's first operand). */
export const COMMIT = 0, PRUNE = 1, SKIP = 2, SKIPN = 3, THEN = 4, MARK = 5
/** LOOKS and LOOKE kinds: bits of a negative assertion, a condition and a lookbehind. */
export const NEG = 1, COND = 2, BEHIND = 4

export interface Program {
  code: Int32Array
  tests: CharTest[]
  refs: number[][]
  names: string[]
  /** Registers: three for each group (0 the whole pattern: start pending, start, end), then counters and starts of iterations. */
  slots: number
  groupPc: number[]
  head: Head
  anchored: boolean
  /** A test of the first character of every match, or null. */
  first: CharTest | null
  /** A literal every match starts with, or ''. */
  prefix: string
  /** A character every match holds, or ''. */
  required: string
  /** Characters every match spans at least. */
  minLength: number
}

const has = (n: Node, f: (n: Node) => boolean): boolean => {
  if (f(n)) return true
  switch (n.t) {
    case 'seq':
      return n.items.some((x) => has(x, f))
    case 'alt':
      return n.alts.some((x) => has(x, f))
    case 'group':
    case 'look':
    case 'rep':
      return has(n.body, f)
    case 'cond':
      return (n.cond.k === 'look' && has(n.cond.look, f)) || has(n.yes, f) || has(n.no, f)
  }
  return false
}

const zeroWidth = (n: Node) => n.t === 'assert' || n.t === 'look' || n.t === 'empty' || (n.t === 'verb' && n.k !== 'accept')
const simple = (n: Node) => n.t === 'char' || n.t === 'set' || n.t === 'any'

/** The items a node starts with, through groups and sequences. */
function leading(n: Node): Node[] {
  if (n.t === 'seq') return n.items
  if (n.t === 'group' && n.body.t !== 'alt') return leading(n.body)
  return [n]
}

/** Whether every match starts at the start of the text (`^`, `\A`, `\G` first). */
function anchoredAt(n: Node): boolean {
  if (n.t === 'alt') return n.alts.every(anchoredAt)
  for (const item of leading(n)) {
    if (item.t === 'assert') {
      if (item.k === 'bol' || item.k === 'A' || item.k === 'G') return true
      continue
    }
    if (zeroWidth(item)) continue
    return item.t === 'group' ? anchoredAt(item.body) : false
  }
  return false
}

/** A test of the first character, or null when it cannot be told. */
function firstOf(n: Node, w: (n: Node) => [number, number]): CharTest | null {
  switch (n.t) {
    case 'char': {
      return n.i ? (c) => sameCase(c, n.c, n.r) : (c) => c === n.c
    }
    case 'set':
      return n.test
    case 'group':
      return firstOf(n.body, w)
    case 'rep':
      return n.min > 0 ? firstOf(n.body, w) : null
    case 'alt': {
      const all = n.alts.map((a) => firstOf(a, w))
      if (all.some((t) => t === null)) return null
      return (c) => all.some((t) => t!(c))
    }
    case 'seq':
      for (const item of n.items) {
        if (zeroWidth(item)) continue
        return w(item)[0] > 0 ? firstOf(item, w) : null
      }
      return null
  }
  return null
}

/** The case-sensitive literal every match starts with. */
function prefixOf(n: Node): string {
  let out = ''
  for (const item of leading(n)) {
    if (zeroWidth(item) && out === '') continue
    if (item.t === 'char' && !item.i) out += String.fromCodePoint(item.c)
    else if (item.t === 'group' && !item.atomic && out === '') return prefixOf(item.body)
    else break
  }
  return out
}

/** The last case-sensitive literal that every match takes, on the top level or in groups taken at least once. */
function requiredOf(n: Node): string {
  let req = ''
  for (const item of leading(n)) {
    if (item.t === 'char') req = item.i ? '' : String.fromCodePoint(item.c)
    else if (item.t === 'group' && item.body.t !== 'alt') req = requiredOf(item.body) || req
    else if (item.t === 'rep' && item.min > 0) req = requiredOf(item.body) || req
  }
  return req
}

const assertCode = (k: string) => ASSERTS.indexOf(k as (typeof ASSERTS)[number])

/** Compiles a parsed pattern; a lookbehind without a bounded length is a `RegexSyntaxError`. */
export function compile(parsed: Parsed): Program {
  const code: number[] = []
  const tests: CharTest[] = []
  const refs: number[][] = []
  const names: string[] = []
  const groupPc: number[] = []
  let slots = 3 * (parsed.groups + 1)
  let groupIds = 0
  const w = widths(parsed)
  const thens = has(parsed.root, (n) => n.t === 'verb' && n.k === 'then')

  const emit = (...xs: number[]) => {
    code.push(...xs)
    return code.length - xs.length
  }
  const here = () => code.length
  const patch = (at: number, v: number) => (code[at] = v)
  const ref = (groups: number[]) => refs.push(groups) - 1
  const name = (s: string) => (names.includes(s) ? names.indexOf(s) : names.push(s) - 1)
  const item = (n: Node): [number, number, number] => {
    if (n.t === 'char') return n.i ? [CHARI, n.c, n.r ? 1 : 0] : [CHAR, n.c, 0]
    if (n.t === 'set') return [SET, tests.push(n.test) - 1, 0]
    return [(n as Node & { t: 'any' }).all ? ANY : ANYNL, 0, 0]
  }

  interface Ctx {
    /** The LOOKE instructions an ACCEPT jumps to: its operand's patches. */
    accept: number[] | null
    then: number
  }

  const node = (n: Node, cx: Ctx, next?: Node): void => {
    switch (n.t) {
      case 'char':
      case 'set':
      case 'any': {
        const [op, a, b] = item(n)
        if (op === CHARI) emit(op, a, b)
        else if (op === ANY || op === ANYNL) emit(op)
        else emit(op, a)
        return
      }
      case 'nl':
        return void emit(NL)
      case 'cluster':
        return void emit(CLUSTER)
      case 'empty':
        return
      case 'seq':
        return n.items.forEach((x, i) => node(x, cx, n.items[i + 1]))
      case 'alt': {
        const gid = thens ? groupIds++ : -1
        if (gid >= 0) emit(GMARK, gid)
        const ends: number[] = []
        n.alts.forEach((a, i) => {
          const last = i === n.alts.length - 1
          const push = last ? -1 : emit(PUSH, 0, gid)
          node(a, { ...cx, then: gid >= 0 ? gid : cx.then })
          if (!last) {
            ends.push(emit(JMP, 0) + 1)
            patch(push + 1, here())
          }
        })
        return ends.forEach((at) => patch(at, here()))
      }
      case 'group':
        if (n.cap >= 0) {
          groupPc[n.cap] = emit(OPEN, n.cap)
          node(n.body, cx)
          return void emit(CLOSE, n.cap)
        }
        if (!n.atomic) return node(n.body, cx)
        emit(ATOMS)
        node(n.body, cx)
        return void emit(ATOME)
      case 'look': {
        const [start] = look(n, 0)
        // a negative assertion that holds goes on after LOOKE
        return void (n.neg && patch(start + 2, here()))
      }
      case 'rep':
        return repeat(n, cx, next)
      case 'assert':
        return void emit(ASSERT, assertCode(n.k))
      case 'bref':
        return void emit(BREF, ref(n.ref.groups), n.i ? 1 : 0, n.r ? 1 : 0)
      case 'call':
        return void emit(CALL, n.ref.groups[0])
      case 'cond':
        return conditional(n, cx)
      case 'verb':
        return verb(n, cx)
    }
  }

  const verb = (n: Node & { t: 'verb' }, cx: Ctx) => {
    if (n.k === 'fail') return void emit(FAIL)
    if (n.k === 'accept') {
      const at = emit(ACCEPT, -1)
      return void cx.accept?.push(at + 1)
    }
    const arg = n.name === null ? -1 : name(n.name)
    if (n.k === 'skip' && arg >= 0) return void emit(VERB, SKIPN, arg, 0)
    if (arg >= 0) emit(VERB, MARK, arg, 0)
    if (n.k === 'commit') emit(VERB, COMMIT, 0, 0)
    else if (n.k === 'prune') emit(VERB, PRUNE, 0, 0)
    else if (n.k === 'skip') emit(VERB, SKIP, 0, 0)
    else if (n.k === 'then') emit(VERB, THEN, cx.then, 0)
  }

  /** An assertion: LOOKS, its alternatives (each moved back by its length in a lookbehind), LOOKE; their addresses. */
  const look = (n: Node & { t: 'look' }, cond: number): [number, number] => {
    const start = emit(LOOKS, (n.neg ? NEG : 0) | cond | (n.behind ? BEHIND : 0), 0)
    const accept: number[] = []
    const inner: Ctx = { accept, then: -1 }
    if (n.behind) {
      const alts = n.body.t === 'alt' ? n.body.alts : [n.body]
      const ends: number[] = []
      alts.forEach((a, i) => {
        const [min, max] = w(a)
        const last = i === alts.length - 1
        const push = last ? -1 : emit(PUSH, 0, -1)
        emit(BACK, min, max)
        node(a, inner)
        if (!last) {
          ends.push(emit(JMP, 0) + 1)
          patch(push + 1, here())
        }
      })
      ends.forEach((at) => patch(at, here()))
    } else node(n.body, inner)
    const end = emit(LOOKE, code[start + 1], 0)
    accept.forEach((at) => patch(at, end))
    return [start, end]
  }

  const conditional = (n: Node & { t: 'cond' }, cx: Ctx) => {
    const c = n.cond
    if (c.k === 'look') {
      // LOOKS's operand: where a failing assertion goes (a negative one holds: the yes branch); LOOKE's: where a holding negative one goes
      const [start, end] = look(c.look, COND)
      if (c.look.neg) patch(start + 2, here())
      node(n.yes, cx)
      const skip = emit(JMP, 0)
      patch(c.look.neg ? end + 2 : start + 2, here())
      node(n.no, cx)
      return void patch(skip + 1, here())
    }
    let toNo = -1
    if (c.k === 'group') toNo = emit(CONDG, ref(c.ref.groups), 0) + 2
    else if (c.k === 'R') toNo = emit(CONDR, c.ref ? c.ref.groups[0] : -1, 0) + 2
    else if (c.k !== 'true') toNo = emit(JMP, 0) + 1
    node(n.yes, cx)
    const skip = emit(JMP, 0)
    if (toNo >= 0) patch(toNo, here())
    node(n.no, cx)
    patch(skip + 1, here())
  }

  const repeat = (n: Node & { t: 'rep' }, cx: Ctx, next?: Node) => {
    const { body, min, max } = n
    let mode = n.mode
    if (max === 0) return
    if (simple(body)) {
      // auto-possessification, as PCRE2's: a repeat that the next literal cannot continue gives nothing back
      const can = firstOf(body, w)
      if (mode !== 2 && next?.t === 'char' && !next.i && can && !can(next.c)) mode = 2
      const [op, a, b] = item(body)
      return void emit(REP, op, a, b, min, max === Infinity ? -1 : max, mode)
    }
    if (min === 1 && max === 1) return node(body, cx)
    if (mode === 2) emit(ATOMS)
    const empty = w(body)[0] === 0 ? slots++ : -1
    const lazy = mode === 1
    if (max === 1) {
      // optional
      if (lazy) {
        const p = emit(PUSH, 0, -1)
        const j = emit(JMP, 0)
        patch(p + 1, here())
        node(body, cx)
        patch(j + 1, here())
      } else {
        const p = emit(PUSH, 0, -1)
        node(body, cx)
        patch(p + 1, here())
      }
    } else if (max === Infinity && min <= 1) {
      // L: [ESTART] body [ECHK] (greedy: PUSH exit, JMP L; lazy: PUSH L, JMP exit)
      const pre = min === 0 ? emit(lazy ? JMP : PUSH, 0, -1) : -1
      if (pre >= 0 && lazy) code.length = pre
      const enter = lazy && min === 0 ? emit(JMP, 0) : -1
      const loop = here()
      if (empty >= 0) emit(ESTART, empty)
      node(body, cx)
      const check = empty >= 0 ? emit(ECHK, empty, -1, 0, 0) : -1
      let exit: number
      if (lazy) {
        if (enter >= 0) patch(enter + 1, here())
        emit(PUSH, loop, -1)
        exit = emit(JMP, 0) + 1
      } else {
        emit(PUSH, 0, -1)
        exit = here() - 2
        emit(JMP, loop)
      }
      const out = here()
      patch(exit, out)
      if (pre >= 0 && !lazy) patch(pre + 1, out)
      if (check >= 0) patch(check + 4, out)
    } else {
      const counter = slots++
      emit(CINIT, counter)
      const loop = emit(CNT, counter, min, max === Infinity ? -1 : max, lazy ? 1 : 0, 0)
      if (empty >= 0) emit(ESTART, empty)
      node(body, cx)
      emit(CINC, counter)
      const check = empty >= 0 ? emit(ECHK, empty, counter, min, 0) : -1
      emit(JMP, loop)
      patch(loop + 5, here())
      if (check >= 0) patch(check + 4, here())
    }
    if (mode === 2) emit(ATOME)
  }

  groupPc[0] = emit(OPEN, 0)
  node(parsed.root, { accept: null, then: -1 })
  emit(CLOSE, 0, MATCH)
  const root = parsed.root
  const accepts = has(root, (x) => x.t === 'verb' && x.k === 'accept')
  return {
    code: Int32Array.from(code),
    tests,
    refs,
    names,
    slots,
    groupPc,
    head: parsed.head,
    anchored: anchoredAt(root),
    first: ((t) => t && remembered(t))(firstOf(root, w)),
    prefix: prefixOf(root),
    // (*ACCEPT) ends a match anywhere
    required: accepts ? '' : requiredOf(root),
    minLength: accepts ? 0 : w(root)[0],
  }
}
