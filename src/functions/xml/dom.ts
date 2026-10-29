/**
 * The XML that FILTERXML reads, this engine's own after the W3C's XML 1.0 (Fifth Edition, https://www.w3.org/TR/xml/)
 * and Namespaces in XML 1.0 (https://www.w3.org/TR/xml-names/): a well-formed document without a DTD (MSXML 6, which
 * Excel reads it with, refuses one), its elements, attributes, text, comments and processing instructions as the XPath
 * 1.0 data model sees them. Whitespace between elements is no text node (MSXML's default); anything not well formed
 * is null, which FILTERXML answers with #VALUE!.
 */
import { decodeXml } from '../../xlsx/xml/text.js'

export type Kind = 'root' | 'element' | 'attribute' | 'text' | 'comment' | 'pi'

export interface XNode {
  kind: Kind
  /** The local name of an element or attribute, a processing instruction's target. */
  name: string
  /** The name as written, with its prefix. */
  qname: string
  /** The namespace of an element or attribute, '' for none. */
  ns: string
  /** The text of a text node, a comment, a processing instruction, an attribute's value. */
  value: string
  parent: XNode | null
  children: XNode[]
  attrs: XNode[]
  /** The place in document order. */
  order: number
}

const XML_NS = 'http://www.w3.org/XML/1998/namespace'
const NAME = /[A-Za-z_À-ÖØ-öø-˿Ͱ-ͽͿ-῿‌‍⁰-↏Ⰰ-⿯、-퟿豈-﷏ﷰ-�][\w.\-·À-ÖØ-öø-ͽͿ-῿‌‍‿⁀⁰-↏Ⰰ-⿯、-퟿豈-﷏ﷰ-�]*/y
/** A character XML 1.0 does not allow, a reference that is no entity of the five or a character. */
const BAD_CHAR = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/
const REFERENCE = /&(?!(?:amp|lt|gt|quot|apos|#[0-9]+|#x[0-9A-Fa-f]+);)/

const node = (kind: Kind, parent: XNode | null, value = ''): XNode => ({ kind, name: '', qname: '', ns: '', value, parent, children: [], attrs: [], order: 0 })

/** Character data with its references: null for a bare `&`, an unknown entity or a reference to a character XML lacks. */
function text(s: string): string | null {
  if (REFERENCE.test(s)) return null
  const t = decodeXml(s)
  return BAD_CHAR.test(t) ? null : t
}

/** The document, or null where it is not well formed. */
export function parseXml(src: string): XNode | null {
  const s = src.replace(/\r\n?/g, '\n')
  if (BAD_CHAR.test(s)) return null
  let p = s.charCodeAt(0) === 0xfeff ? 1 : 0
  const root = node('root', null)
  let cur = root
  /** In-scope namespace prefixes, innermost last. */
  const scopes: Record<string, string>[] = [{ xml: XML_NS }]
  const lookup = (prefix: string): string | undefined => {
    for (let i = scopes.length - 1; i >= 0; i--) if (prefix in scopes[i]) return scopes[i][prefix]
    return undefined
  }
  let seenRoot = false
  const name = (): string | null => {
    NAME.lastIndex = p
    const m = NAME.exec(s)
    if (!m) return null
    p = NAME.lastIndex
    return m[0]
  }
  const ws = () => {
    while (p < s.length && ' \t\n'.includes(s[p])) p++
  }
  /** A QName's prefix and local part, null for more than one colon or an empty part. */
  const split = (q: string): [string, string] | null => {
    const parts = q.split(':')
    if (parts.length > 2 || parts.some((x) => x === '')) return null
    return parts.length === 2 ? [parts[0], parts[1]] : ['', q]
  }
  const pushText = (t: string) => {
    const last = cur.children.at(-1)
    if (last?.kind === 'text') last.value += t
    else cur.children.push(node('text', cur, t))
  }

  if (s.startsWith('<?xml', p) && /[\s?]/.test(s[p + 5] ?? '')) {
    const end = s.indexOf('?>', p)
    if (end < 0 || !/^<\?xml\s+version\s*=\s*(["'])1\.\d+\1(\s+encoding\s*=\s*(["'])[A-Za-z][\w.-]*\3)?(\s+standalone\s*=\s*(["'])(yes|no)\5)?\s*$/.test(s.slice(p, end))) return null
    p = end + 2
  }
  while (p < s.length) {
    const lt = s.indexOf('<', p)
    const chunk = s.slice(p, lt < 0 ? s.length : lt)
    if (chunk) {
      if (cur === root) {
        if (/\S/.test(chunk)) return null
      } else {
        if (chunk.includes(']]>')) return null
        const t = text(chunk)
        if (t === null) return null
        pushText(t)
      }
    }
    if (lt < 0) break
    p = lt
    if (s.startsWith('<!--', p)) {
      const end = s.indexOf('-->', p + 4)
      if (end < 0 || s.slice(p + 4, end).includes('--') || s[end - 1] === '-') return null
      cur.children.push(node('comment', cur, s.slice(p + 4, end)))
      p = end + 3
    } else if (s.startsWith('<![CDATA[', p)) {
      const end = s.indexOf(']]>', p + 9)
      if (end < 0 || cur === root) return null
      pushText(s.slice(p + 9, end))
      p = end + 3
    } else if (s.startsWith('<?', p)) {
      p += 2
      const target = name()
      const end = s.indexOf('?>', p)
      if (!target || end < 0 || /^xml$/i.test(target) || target.includes(':') || (p < end && !' \t\n'.includes(s[p]))) return null
      const pi = node('pi', cur, s.slice(p, end).replace(/^[ \t\n]+/, ''))
      pi.name = pi.qname = target
      cur.children.push(pi)
      p = end + 2
    } else if (s.startsWith('</', p)) {
      p += 2
      const q = name()
      ws()
      if (!q || q !== cur.qname || s[p] !== '>') return null
      p++
      scopes.pop()
      cur = cur.parent!
    } else if (s[p + 1] === '!') {
      return null
    } else {
      p++
      const q = name()
      if (!q || (cur === root && seenRoot)) return null
      const el = node('element', cur)
      el.qname = q
      const raw: [string, string][] = []
      const scope: Record<string, string> = {}
      for (;;) {
        const before = p
        ws()
        if (s[p] === '>' || s.startsWith('/>', p)) break
        if (p === before) return null
        const a = name()
        ws()
        if (!a || s[p] !== '=') return null
        p++
        ws()
        const quote = s[p]
        if (quote !== '"' && quote !== "'") return null
        const end = s.indexOf(quote, p + 1)
        if (end < 0) return null
        const v = s.slice(p + 1, end)
        p = end + 1
        if (v.includes('<') || raw.some(([n]) => n === a)) return null
        const t = text(v.replace(/[\t\n]/g, ' '))
        if (t === null) return null
        if (a === 'xmlns') scope[''] = t
        else if (a.startsWith('xmlns:')) {
          if (!t || a === 'xmlns:xmlns') return null
          scope[a.slice(6)] = t
        } else raw.push([a, t])
      }
      scopes.push(scope)
      const qn = split(q)
      if (!qn) return null
      const ns = lookup(qn[0]) ?? (qn[0] ? undefined : '')
      if (ns === undefined) return null
      ;[el.name, el.ns] = [qn[1], ns]
      for (const [a, v] of raw) {
        const an = split(a)
        if (!an) return null
        // an attribute without a prefix is in no namespace, whatever the default
        const ans = an[0] ? lookup(an[0]) : ''
        if (ans === undefined) return null
        const at = node('attribute', el, v)
        ;[at.name, at.qname, at.ns] = [an[1], a, ans]
        if (el.attrs.some((x) => x.name === at.name && x.ns === at.ns)) return null
        el.attrs.push(at)
      }
      cur.children.push(el)
      if (cur === root) seenRoot = true
      if (s.startsWith('/>', p)) {
        p += 2
        scopes.pop()
      } else {
        p++
        cur = el
      }
    }
  }
  if (cur !== root || !seenRoot) return null
  let order = 0
  const number = (n: XNode) => {
    n.order = order++
    for (const a of n.attrs) a.order = order++
    for (const c of n.children) {
      if (c.kind === 'text' && !/\S/.test(c.value)) continue
      number(c)
    }
    n.children = n.children.filter((c) => c.kind !== 'text' || /\S/.test(c.value))
  }
  number(root)
  return root
}

/** The string-value of a node: the text of an element's or the root's descendants, a leaf's own text. */
export function stringValue(n: XNode): string {
  if (n.kind !== 'element' && n.kind !== 'root') return n.value
  let out = ''
  const walk = (x: XNode) => {
    for (const c of x.children) {
      if (c.kind === 'text') out += c.value
      else if (c.kind === 'element') walk(c)
    }
  }
  walk(n)
  return out
}
