/**
 * XPath 1.0 over the nodes of `dom.ts`, this engine's own after the W3C's recommendation
 * (https://www.w3.org/TR/1999/REC-xpath-19991116/): the axes, node tests, predicates, operators and the core function
 * library. A name test compares the name as written, prefix and all, whatever namespace it is in, as Excel's FILTERXML
 * does (`//b` finds `b` under a default namespace, `//x:a` finds `x:a`). There are no variables; an expression outside
 * the grammar is an error.
 */
import { stringValue, type XNode } from './dom.js'

export type XValue = XNode[] | number | string | boolean

class XPathError extends Error {}
const fail = (what: string): never => {
  throw new XPathError(what)
}

interface Context {
  node: XNode
  pos: number
  size: number
}
type Eval = (c: Context) => XValue

// tokens: numbers, literals, names (with a prefix or `*`), two-character operators, single characters
const TOKEN = /\s*(\d+(?:\.\d*)?|\.\d+|"[^"]*"|'[^']*'|[A-Za-z_À-�][\w.\-·À-�]*(?::(?:\*|[A-Za-z_À-�][\w.\-·À-�]*))?|\/\/|::|\.\.|!=|<=|>=|\$|[/()[\]@,|+\-=<>*.])/y

function tokens(src: string): string[] {
  const out: string[] = []
  TOKEN.lastIndex = 0
  while (TOKEN.lastIndex < src.length) {
    const at = TOKEN.lastIndex
    const m = TOKEN.exec(src)
    if (!m) {
      if (/^\s*$/.test(src.slice(at))) break
      fail('token')
    }
    out.push(m![1])
  }
  return out
}

const isNodes = (v: XValue): v is XNode[] => Array.isArray(v)

/** XPath's number of a string: digits with one point, signed, between whitespace; anything else NaN. */
const toNumber = (v: XValue): number => {
  if (typeof v === 'number') return v
  if (typeof v === 'boolean') return v ? 1 : 0
  const s = (isNodes(v) ? (v.length ? stringValue(v[0]) : '') : v).trim()
  return /^-?(\d+(\.\d*)?|\.\d+)$/.test(s) ? Number(s) : NaN
}

/** XPath's string of a number: no exponent, integers without a point. */
function numberText(n: number): string {
  if (Number.isNaN(n)) return 'NaN'
  if (!Number.isFinite(n)) return n > 0 ? 'Infinity' : '-Infinity'
  if (Object.is(n, -0) || n === 0) return '0'
  const s = String(n)
  if (!s.includes('e')) return s
  const [m, e] = s.split('e')
  const digits = m.replace('-', '').replace('.', '')
  const point = (m.replace('-', '').indexOf('.') + 1 || m.replace('-', '').length + 1) - 1 + Number(e)
  const body = point <= 0 ? '0.' + '0'.repeat(-point) + digits : point >= digits.length ? digits + '0'.repeat(point - digits.length) : digits.slice(0, point) + '.' + digits.slice(point)
  return (n < 0 ? '-' : '') + body
}

const toString = (v: XValue): string =>
  isNodes(v) ? (v.length ? stringValue(v[0]) : '') : typeof v === 'number' ? numberText(v) : typeof v === 'boolean' ? String(v) : v
const toBoolean = (v: XValue): boolean => (isNodes(v) ? v.length > 0 : typeof v === 'number' ? v !== 0 && !Number.isNaN(v) : typeof v === 'string' ? v.length > 0 : v)

/** Nodes in document order, each once. */
const ordered = (ns: XNode[]): XNode[] => [...new Set(ns)].sort((a, b) => a.order - b.order)

function descendants(n: XNode, out: XNode[]) {
  for (const c of n.children) {
    out.push(c)
    descendants(c, out)
  }
}

/** The nodes along an axis from a node, in the axis's own order (reverse axes nearest first). */
function axis(name: string, n: XNode): XNode[] {
  const out: XNode[] = []
  switch (name) {
    case 'child':
      return n.children
    case 'attribute':
      return n.attrs
    case 'self':
      return [n]
    case 'parent':
      return n.parent ? [n.parent] : []
    case 'descendant':
      descendants(n, out)
      return out
    case 'descendant-or-self':
      out.push(n)
      descendants(n, out)
      return out
    case 'ancestor':
    case 'ancestor-or-self':
      for (let p = name === 'ancestor' ? n.parent : n; p; p = p.parent) out.push(p)
      return out
    case 'following-sibling':
    case 'preceding-sibling': {
      if (!n.parent || n.kind === 'attribute') return []
      const sib = n.parent.children
      const i = sib.indexOf(n)
      return name === 'following-sibling' ? sib.slice(i + 1) : sib.slice(0, i).reverse()
    }
    case 'following':
    case 'preceding': {
      let top = n
      while (top.parent) top = top.parent
      const all: XNode[] = []
      descendants(top, all)
      const anc = new Set(axis('ancestor', n))
      const at = n.kind === 'attribute' ? n.parent! : n
      const end = name === 'following' ? at.order + 1 + countBelow(at) : 0
      return name === 'following' ? all.filter((x) => x.order >= end) : all.filter((x) => x.order < at.order && !anc.has(x)).reverse()
    }
    case 'namespace':
      return namespaces(n)
  }
  return fail('axis')
}

/** The namespace nodes of an element: its prefixes in scope, `xml` first, each its URI. */
function namespaces(n: XNode): XNode[] {
  if (n.kind !== 'element') return []
  const scope = new Map<string, string>([['xml', 'http://www.w3.org/XML/1998/namespace']])
  const chain: XNode[] = []
  for (let p: XNode | null = n; p; p = p.parent) chain.unshift(p)
  for (const p of chain) for (const [k, v] of Object.entries(p.declared)) scope.set(k, v)
  return [...scope].filter(([, v]) => v).map(([k, v], i) => ({ ...n, kind: 'namespace', name: k, qname: k, ns: '', value: v, parent: n, children: [], attrs: [], declared: {}, order: n.order + (i + 1) / 1024 }))
}

/** How many nodes and attributes lie inside a node, so that its following nodes start past them. */
function countBelow(n: XNode): number {
  let k = n.attrs.length
  for (const c of n.children) k += 1 + countBelow(c)
  return k
}

const REVERSE = new Set(['ancestor', 'ancestor-or-self', 'preceding', 'preceding-sibling'])
const AXES = new Set(['child', 'attribute', 'self', 'parent', 'descendant', 'descendant-or-self', 'ancestor', 'ancestor-or-self', 'following-sibling', 'preceding-sibling', 'following', 'preceding', 'namespace'])
const NODE_TYPES = new Set(['node', 'text', 'comment', 'processing-instruction'])

/** Keeps the nodes a predicate holds for, positions counted in the axis's order. */
function filter(ns: XNode[], pred: Eval): XNode[] {
  return ns.filter((node, i) => {
    const v = pred({ node, pos: i + 1, size: ns.length })
    return typeof v === 'number' ? v === i + 1 : toBoolean(v)
  })
}

function compare(op: string, a: XValue, b: XValue): boolean {
  if (isNodes(a) || isNodes(b)) {
    if (isNodes(a) && isNodes(b)) return a.some((x) => b.some((y) => compare(op, stringValue(x), stringValue(y))))
    const [ns, other, flip] = isNodes(a) ? [a, b, false] : [b as XNode[], a, true]
    if (typeof other === 'boolean') return compare(op, toBoolean(ns), other)
    return ns.some((x) => {
      const s: XValue = typeof other === 'number' ? toNumber(stringValue(x)) : stringValue(x)
      return flip ? compare(op, other, s) : compare(op, s, other)
    })
  }
  if (op === '=' || op === '!=') {
    const eq = typeof a === 'boolean' || typeof b === 'boolean' ? toBoolean(a) === toBoolean(b) : typeof a === 'number' || typeof b === 'number' ? toNumber(a) === toNumber(b) : a === b
    return op === '=' ? eq : !eq
  }
  const x = toNumber(a)
  const y = toNumber(b)
  return op === '<' ? x < y : op === '>' ? x > y : op === '<=' ? x <= y : x >= y
}

const FUNCTIONS: Record<string, (c: Context, a: XValue[]) => XValue> = {
  last: (c) => c.size,
  position: (c) => c.pos,
  count: (_, [v]) => (isNodes(v) ? v.length : fail('count')),
  id: () => [],
  'local-name': (c, a) => nameOf(a.length ? a[0] : [c.node], (n) => n.name),
  'namespace-uri': (c, a) => nameOf(a.length ? a[0] : [c.node], (n) => n.ns),
  name: (c, a) => nameOf(a.length ? a[0] : [c.node], (n) => n.qname),
  string: (c, a) => toString(a.length ? a[0] : [c.node]),
  concat: (_, a) => (a.length < 2 ? fail('concat') : a.map(toString).join('')),
  'starts-with': (_, [s, t]) => toString(s).startsWith(toString(t)),
  contains: (_, [s, t]) => toString(s).includes(toString(t)),
  'substring-before': (_, [s, t]) => around(toString(s), toString(t), false),
  'substring-after': (_, [s, t]) => around(toString(s), toString(t), true),
  substring: (_, [s, from, len]) => {
    const chars = [...toString(s)]
    const a = Math.round(toNumber(from))
    const b = len === undefined ? Infinity : a + Math.round(toNumber(len))
    return chars.filter((_, i) => i + 1 >= a && i + 1 < b).join('')
  },
  'string-length': (c, a) => [...toString(a.length ? a[0] : [c.node])].length,
  'normalize-space': (c, a) => toString(a.length ? a[0] : [c.node]).replace(/[ \t\r\n]+/g, ' ').trim(),
  translate: (_, [s, from, to]) => {
    const f = [...toString(from)]
    const t = [...toString(to)]
    return [...toString(s)].map((ch) => (f.indexOf(ch) < 0 ? ch : (t[f.indexOf(ch)] ?? ''))).join('')
  },
  boolean: (_, [v]) => toBoolean(v),
  not: (_, [v]) => !toBoolean(v),
  true: () => true,
  false: () => false,
  lang: (c, [v]) => {
    for (let n: XNode | null = c.node; n; n = n.parent) {
      const at = n.attrs.find((x) => x.name === 'lang' && x.ns === 'http://www.w3.org/XML/1998/namespace')
      if (at) return at.value.toLowerCase().split('-')[0] === toString(v).toLowerCase() || at.value.toLowerCase() === toString(v).toLowerCase()
    }
    return false
  },
  number: (c, a) => toNumber(a.length ? a[0] : [c.node]),
  sum: (_, [v]) => (isNodes(v) ? v.reduce((s, n) => s + toNumber(stringValue(n)), 0) : fail('sum')),
  floor: (_, [v]) => Math.floor(toNumber(v)),
  ceiling: (_, [v]) => Math.ceil(toNumber(v)),
  round: (_, [v]) => {
    const x = toNumber(v)
    return Number.isFinite(x) ? Math.floor(x + 0.5) : x
  },
}

/** The argument counts of the core functions: [least, most]. */
const ARITY: Record<string, [number, number]> = {
  last: [0, 0], position: [0, 0], count: [1, 1], id: [1, 1], 'local-name': [0, 1], 'namespace-uri': [0, 1], name: [0, 1],
  string: [0, 1], concat: [2, Infinity], 'starts-with': [2, 2], contains: [2, 2], 'substring-before': [2, 2], 'substring-after': [2, 2],
  substring: [2, 3], 'string-length': [0, 1], 'normalize-space': [0, 1], translate: [3, 3], boolean: [1, 1], not: [1, 1], true: [0, 0],
  false: [0, 0], lang: [1, 1], number: [0, 1], sum: [1, 1], floor: [1, 1], ceiling: [1, 1], round: [1, 1],
}

/** The text before or after the first occurrence of a part, '' without one. */
function around(x: string, y: string, after: boolean): string {
  const i = x.indexOf(y)
  return i < 0 ? '' : after ? x.slice(i + y.length) : x.slice(0, i)
}

const nameOf = (v: XValue, f: (n: XNode) => string): string => (isNodes(v) ? (v.length && v[0].kind !== 'text' && v[0].kind !== 'comment' && v[0].kind !== 'root' ? f(v[0]) : '') : fail('name'))

/** The parser: recursive descent over the tokens, building the evaluation as closures. */
function compile(src: string): Eval {
  const t = tokens(src)
  let i = 0
  const peek = (k = 0) => t[i + k]
  const take = (s?: string) => {
    if (s !== undefined && t[i] !== s) fail(`expected ${s}`)
    return t[i++]
  }
  const isName = (s: string | undefined) => !!s && /^[A-Za-z_À-�]/.test(s)

  const binary = (next: () => Eval, ops: string[], apply: (op: string, a: XValue, b: XValue) => XValue) => (): Eval => {
    let left = next()
    while (ops.includes(peek())) {
      const op = take()
      const l = left
      const r = next()
      left = (c) => apply(op, l(c), r(c))
    }
    return left
  }
  const arith = (op: string, a: XValue, b: XValue): number => {
    const x = toNumber(a)
    const y = toNumber(b)
    return op === '+' ? x + y : op === '-' ? x - y : op === '*' ? x * y : op === 'div' ? x / y : x % y
  }

  const step = (): ((n: XNode) => XNode[]) => {
    if (peek() === '.') return take(), (n) => [n]
    if (peek() === '..') return take(), (n) => (n.parent ? [n.parent] : [])
    let ax = 'child'
    if (peek() === '@') take(), (ax = 'attribute')
    else if (isName(peek()) && peek(1) === '::') {
      ax = take()
      take('::')
      if (!AXES.has(ax)) fail('axis')
    }
    const test = take()
    if (test === undefined) fail('step')
    let match: (n: XNode) => boolean
    const principal = ax === 'attribute' ? 'attribute' : ax === 'namespace' ? 'namespace' : 'element'
    if (NODE_TYPES.has(test) && peek() === '(') {
      take('(')
      let target: string | undefined
      if (test === 'processing-instruction' && peek() !== ')') target = take().slice(1, -1)
      take(')')
      match = test === 'node' ? () => true : test === 'text' ? (n) => n.kind === 'text' : test === 'comment' ? (n) => n.kind === 'comment' : (n) => n.kind === 'pi' && (target === undefined || n.name === target)
    } else if (test === '*') match = (n) => n.kind === principal
    else if (test.endsWith(':*')) match = (n) => n.kind === principal && n.qname.startsWith(test.slice(0, -1))
    else if (isName(test)) match = (n) => n.kind === principal && n.qname === test
    else return fail('node test')
    const preds: Eval[] = []
    while (peek() === '[') {
      take('[')
      preds.push(expr())
      take(']')
    }
    const reverse = REVERSE.has(ax)
    return (n) => {
      let ns = axis(ax, n).filter(match)
      for (const p of preds) ns = filter(ns, p)
      return reverse ? ns.reverse() : ns
    }
  }

  const relative = (start: Eval): Eval => {
    let cur = start
    while (peek() === '/' || peek() === '//') {
      const deep = take() === '//'
      const s = step()
      const prev = cur
      cur = (c) => {
        const v = prev(c)
        if (!isNodes(v)) return fail('path over a value')
        const from = deep ? ordered(v.flatMap((n) => axis('descendant-or-self', n))) : v
        return ordered(from.flatMap(s))
      }
    }
    return cur
  }

  const path = (): Eval => {
    const tok = peek()
    if (tok === '/' || tok === '//') {
      const top: Eval = (c) => {
        let n = c.node
        while (n.parent) n = n.parent
        return [n]
      }
      if (tok === '/' && !(peek(1) === '.' || peek(1) === '..' || peek(1) === '@' || peek(1) === '*' || isName(peek(1)))) return take(), top
      if (tok === '/') {
        take()
        const s = step()
        return relative((c) => ordered((top(c) as XNode[]).flatMap(s)))
      }
      return relative(top)
    }
    const primary = isPrimary()
    if (primary) {
      let v = primary
      while (peek() === '[') {
        take('[')
        const p = expr()
        take(']')
        const prev = v
        v = (c) => {
          const x = prev(c)
          return isNodes(x) ? filter(x, p) : fail('predicate over a value')
        }
      }
      return relative(v)
    }
    const s = step()
    return relative((c) => s(c.node))
  }

  const isPrimary = (): Eval | null => {
    const tok = peek()
    if (tok === undefined) return fail('end')
    if (tok === '$') return fail('variable')
    if (tok === '(') {
      take()
      const e = expr()
      take(')')
      return e
    }
    if (tok[0] === '"' || tok[0] === "'") {
      take()
      const s = tok.slice(1, -1)
      return () => s
    }
    if (/^[\d.]/.test(tok) && tok !== '.' && tok !== '..') {
      take()
      const n = Number(tok)
      return () => n
    }
    if (isName(tok) && peek(1) === '(' && !NODE_TYPES.has(tok)) {
      take()
      take('(')
      const args: Eval[] = []
      while (peek() !== ')') {
        args.push(expr())
        if (peek() === ',') take()
        else if (peek() !== ')') fail('arguments')
      }
      take(')')
      const f = FUNCTIONS[tok]
      const arity = ARITY[tok]
      if (!f || args.length < arity[0] || args.length > arity[1]) fail(`function ${tok}`)
      return (c) => f(c, args.map((a) => a(c)))
    }
    return null
  }

  const union = (): Eval => {
    let left = path()
    while (peek() === '|') {
      take()
      const l = left
      const r = path()
      left = (c) => {
        const a = l(c)
        const b = r(c)
        return isNodes(a) && isNodes(b) ? ordered([...a, ...b]) : fail('union of values')
      }
    }
    return left
  }
  const unary = (): Eval => {
    if (peek() === '-') {
      take()
      const u = unary()
      return (c) => -toNumber(u(c))
    }
    return union()
  }
  // `*` and the operator names multiply after a value, else they are name tests (`*`, `div` as an element)
  const multiplicative = (): Eval => {
    let left = unary()
    while (peek() === '*' || peek() === 'div' || peek() === 'mod') {
      const op = take()
      const l = left
      const r = unary()
      left = (c) => arith(op, l(c), r(c))
    }
    return left
  }
  const additive = binary(multiplicative, ['+', '-'], arith)
  const relational = binary(additive, ['<', '>', '<=', '>='], compare)
  const equality = binary(relational, ['=', '!='], compare)
  const and = binary(equality, ['and'], (_, a, b) => toBoolean(a) && toBoolean(b))
  const expr: () => Eval = binary(and, ['or'], (_, a, b) => toBoolean(a) || toBoolean(b))

  const e = expr()
  if (i < t.length) fail('trailing tokens')
  return e
}

/** The nodes an expression selects from the document, null for an expression that is not one or gives no node-set. */
export function selectNodes(doc: XNode, src: string): XNode[] | null {
  try {
    const v = compile(src)({ node: doc, pos: 1, size: 1 })
    return isNodes(v) ? v : null
  } catch (e) {
    if (e instanceof XPathError) return null
    throw e
  }
}
