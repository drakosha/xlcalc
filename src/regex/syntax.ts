/**
 * `pcreSyntax`: a pattern of match mode 3 parsed as the engine parses it (`parse.ts`) and written as a JavaScript
 * regular expression with the flag `v`, so that it accepts what Excel accepts. Possessive quantifiers and atomic groups
 * become a lookahead and a back reference, `(?i)` past the start the characters of each letter's case, the dot and the
 * anchors PCRE2's newlines, a condition on an assertion two alternatives. Recursion, conditions on groups, `(*ACCEPT)`
 * and the verbs that cut backtracking, `\X` and `(*NOTEMPTY)` have no translation: null, as for a pattern that is none.
 */
import { boundary, END, START, textEnd, typeClass } from '../functions/lookup/regex-syntax.js'
import { caseVariants } from './chars.js'
import { hex, V, W } from './classes.js'
import { type Node, type Parsed, parse, RegexSyntaxError } from './parse.js'

class Untranslatable extends Error {}
const no = (): never => {
  throw new Untranslatable()
}

/** The newline characters of each convention as a class's insides (`Newline`; CRLF as ANYCRLF). */
const NEWLINES = ['\\r', '\\n', '\\n\\r', '\\n\\r', '\\n-\\r\\x85\\u2028\\u2029', '\\0']

/** A literal character: letters and digits as they are, any other by its code. */
const lit = (c: number) => (/[\p{L}\p{N}]/u.test(String.fromCodePoint(c)) ? String.fromCodePoint(c) : hex(c))

const quantifier = (min: number, max: number) =>
  max === Infinity ? (min === 0 ? '*' : min === 1 ? '+' : `{${min},}`) : min === 0 && max === 1 ? '?' : min === max ? `{${min}}` : `{${min},${max}}`

/** A PCRE2 pattern as a JavaScript regular expression, or null where it has no translation or is no pattern. */
export function pcreToRegExp(pattern: string): RegExp | null {
  try {
    return translate(parse(pattern))
  } catch (e) {
    if (e instanceof RegexSyntaxError || e instanceof Untranslatable || e instanceof SyntaxError) return null
    throw e
  }
}

/** Each node of a tree, depth first. */
function* nodes(n: Node): Generator<Node> {
  yield n
  if (n.t === 'seq') for (const x of n.items) yield* nodes(x)
  else if (n.t === 'alt') for (const x of n.alts) yield* nodes(x)
  else if (n.t === 'group' || n.t === 'look' || n.t === 'rep') yield* nodes(n.body)
  else if (n.t === 'cond') {
    if (n.cond.k === 'look') yield* nodes(n.cond.look)
    yield* nodes(n.yes)
    yield* nodes(n.no)
  }
}

function translate({ root, head }: Parsed): RegExp {
  if (head.notEmpty || head.notEmptyAtStart) no()
  const nl = NEWLINES[head.newline]
  const all = [...nodes(root)]
  // the whole pattern caseless: JavaScript's flag i; else the caseless characters by their case's
  const cased = all.filter((n) => n.t === 'char' || n.t === 'bref' || (n.t === 'set' && n.ranges)) as (Node & { t: 'char' | 'bref' | 'set' })[]
  const caseless = (n: (typeof cased)[number]) => (n.t === 'set' ? n.caseless! : n.i)
  const whole = cased.length > 0 && cased.every(caseless)
  // a group a back reference names gets one name for every group of the reference (duplicate names, as JavaScript
  // takes them from Node 23); another a name of its own each time, as in (?|…)
  const shared = new Map<number, number>()
  for (const n of all) if (n.t === 'bref') for (const g of n.ref.groups) shared.set(g, shared.get(g) ?? n.ref.groups[0])
  const seen = new Map<number, number>()
  const groupName = (g: number) => {
    if (shared.has(g)) return `_${shared.get(g)}`
    const k = seen.get(g) ?? 0
    seen.set(g, k + 1)
    return k ? `_${g}_${k}` : `_${g}`
  }
  let atoms = 0
  const atomic = (js: string) => {
    const n = `_a${atoms++}`
    return `(?:(?=(?<${n}>${js}))\\k<${n}>)`
  }
  const look = (n: Node & { t: 'look' }, neg = n.neg) => `(?${n.behind ? '<' : ''}${neg ? '!' : '='}${js(n.body)})`

  const set = (n: Node & { t: 'set' }) => {
    if (!n.ranges || !n.caseless || whole) return n.js
    const extra = new Set<number>()
    for (let k = 0; k < n.ranges.length; k += 2) {
      if (n.ranges[k + 1] - n.ranges[k] > 5000) no()
      for (let c = n.ranges[k]; c <= n.ranges[k + 1]; c++) for (const v of caseVariants(c, n.r!)) extra.add(v)
    }
    return `[${n.negated ? '^' : ''}${[...extra].map(hex).join('')}${n.parts!.join('')}]`
  }

  const js = (n: Node): string => {
    switch (n.t) {
      case 'char':
        return n.i && !whole ? `[${caseVariants(n.c, n.r).map(hex).join('')}]` : lit(n.c)
      case 'set':
        return set(n)
      case 'any':
        return n.all ? '[\\s\\S]' : `[^${nl}]`
      case 'seq':
        return n.items.map(js).join('')
      case 'alt':
        return n.alts.map(js).join('|')
      case 'group':
        return n.cap >= 0 ? `(?<${groupName(n.cap)}>${js(n.body)})` : n.atomic ? atomic(js(n.body)) : `(?:${js(n.body)})`
      case 'look':
        return look(n)
      case 'rep': {
        if (n.max === 0) return ''
        const body = n.body.t === 'look' ? `(?:${js(n.body)})` : js(n.body)
        const q = quantifier(n.min, n.max)
        return n.mode === 2 ? atomic(body + q) : body + q + (n.mode === 1 ? '?' : '')
      }
      case 'assert':
        switch (n.k) {
          case 'bol':
          case 'A':
          case 'G':
            return START
          case 'mbol':
            // after a newline at the end too, as in Excel
            return `(?:(?<![\\s\\S])|(?<=[${nl}])(?<!\\r(?=\\n)))`
          case 'z':
            return END
          case 'meol':
            return textEnd(nl, true)
          case 'eol':
          case 'Z':
            return textEnd(nl)
          default:
            return boundary(n.k.startsWith('a') ? typeClass('w', { w: true })! : W, n.k.endsWith('B'))
        }
      case 'bref':
        return n.i && !whole ? no() : `\\k<${groupName(n.ref.groups[0])}>`
      case 'cond': {
        const c = n.cond
        if (c.k === 'define') return ''
        if (c.k === 'true') return js(n.yes)
        if (c.k === 'false') return js(n.no)
        // on an assertion: the yes branch where it holds, the no branch where it does not
        if (c.k === 'look') return `(?:${look(c.look)}${js(n.yes)}|${look(c.look, !c.look.neg)}${js(n.no)})`
        return no()
      }
      case 'verb':
        return n.k === 'fail' ? '(?!)' : n.k === 'mark' ? '' : no()
      case 'nl':
        return `(?:\\r\\n|${head.anyR ? V : '[\\n\\r]'})`
      case 'empty':
        return ''
      default:
        return no()
    }
  }

  return new RegExp(js(root), whole ? 'iv' : 'v')
}
