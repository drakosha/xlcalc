/**
 * Elements and attributes of XML parts by regular expressions, without a DOM: a tag may carry a namespace
 * prefix (`<x:c>`), attributes may be quoted either way and hold `>`, and a self-closing element is matched as
 * such (it never swallows the next element). Elements of the same name are not nested in the parts read here.
 */

/**
 * Attributes of a tag up to its end, quoted values skipped whole. They never hold `<` (XML refuses it in an attribute
 * value), so a tag left open costs no more than the text up to the next tag.
 */
const ATTRS = `(?:\\s(?:"[^"<]*"|'[^'<]*'|[^>"'<])*?)?`

export interface Element {
  /** The attribute text of the opening tag. */
  attrs: string
  /** The content, null for a self-closing element. */
  body: string | null
}

/** An element and where it lies in the text: from its opening tag to the end of its closing tag. */
export interface Placed extends Element {
  start: number
  end: number
}

const cache = new Map<string, [RegExp, RegExp]>()

/** The opening tag of a local name (its last group `/` when the element closes itself) and its closing tag; `bare`, without a namespace prefix. */
function tags(local: string, bare: boolean): [RegExp, RegExp] {
  // a local name holds no colon
  const key = bare ? `:${local}` : local
  let r = cache.get(key)
  if (!r) {
    const prefix = bare ? '' : '(?:[\\w.-]+:)?'
    cache.set(key, (r = [new RegExp(`<${prefix}${local}(${ATTRS})(/?)>`, 'g'), new RegExp(`</${prefix}${local}\\s*>`, 'g')]))
  }
  return r
}

/**
 * The elements of a local name in document order (with `first`, the first alone), as `<local …(?:/>|>(…)</local>)`
 * matches them, in time linear in the text: an opening tag without a closing tag after it costs one search for the
 * closing tag, where the regex searched to the end of the text once for each such tag.
 */
export function placed(xml: string, local: string, first = false, bare = false): Placed[] {
  const [open, close] = tags(local, bare)
  const out: Placed[] = []
  // the first closing tag at or after `from` lies from `at` to `to`; `at` -1: there is none
  let from = Infinity
  let at = -1
  let to = -1
  open.lastIndex = 0
  for (let m = open.exec(xml); m; m = open.exec(xml)) {
    const body = open.lastIndex
    if (m[2]) out.push({ attrs: m[1], body: null, start: m.index, end: body })
    else {
      if (body < from || (at >= 0 && body > at)) {
        close.lastIndex = body
        const c = close.exec(xml)
        from = body
        at = c ? c.index : -1
        to = c ? close.lastIndex : -1
      }
      // without its closing tag the element is none; the text after its name is searched again
      if (at < 0) {
        open.lastIndex = m.index + 1
        continue
      }
      out.push({ attrs: m[1], body: xml.slice(body, at), start: m.index, end: to })
      open.lastIndex = to
    }
    if (first) break
  }
  open.lastIndex = 0
  return out
}

/** Every element of a local name, in document order. */
export const elements = (xml: string, local: string): Element[] => placed(xml, local)

/** The first element of a local name, or undefined. */
export const element = (xml: string, local: string): Element | undefined => placed(xml, local, true)[0]

/**
 * The element of a local name the text ends with, but for spaces and closing tags after it: the last child of the
 * element the text closes (a worksheet's own `<extLst>`, where its other children may hold an `<extLst>` of their own).
 * Elements of its name inside it are passed over by depth; undefined when the text ends with another element.
 */
export function lastChild(xml: string, local: string): Placed | undefined {
  // back over spaces and closing tags to the closing tag of the name
  let i = xml.length
  let close = -1
  while (close < 0) {
    while (i > 0 && xml.charCodeAt(i - 1) <= 32) i--
    const lt = xml.lastIndexOf('<', i - 1)
    if (lt < 0 || xml[i - 1] !== '>' || xml[lt + 1] !== '/') return undefined
    const name = xml.slice(lt + 2, i - 1).trimEnd()
    if (name === local || name.endsWith(`:${local}`)) close = lt
    else i = lt
  }
  // its opening tag: the last one before it that no closing tag of the name between them closes
  const [open, closing] = tags(local, false)
  const marks: [at: number, depth: number, attrs: string, end: number][] = []
  open.lastIndex = 0
  for (let m = open.exec(xml); m && m.index < close; m = open.exec(xml)) if (!m[2]) marks.push([m.index, 1, m[1], open.lastIndex])
  closing.lastIndex = 0
  for (let m = closing.exec(xml); m && m.index < close; m = closing.exec(xml)) marks.push([m.index, -1, '', 0])
  open.lastIndex = closing.lastIndex = 0
  marks.sort((a, b) => b[0] - a[0])
  let depth = 0
  for (const [start, d, attrs, end] of marks) {
    depth += d
    if (depth === 1) return { attrs, body: xml.slice(end, close), start, end: i }
  }
  return undefined
}

/** The text without the elements `placed` gave for it. */
export function outside(xml: string, parts: Placed[]): string {
  let out = ''
  let at = 0
  for (const p of parts) {
    out += xml.slice(at, p.start)
    at = p.end
  }
  return out + xml.slice(at)
}

/** Whether the text has an element of the local name. */
export const has = (xml: string, local: string): boolean => new RegExp(`<(?:[\\w.-]+:)?${local}[\\s/>]`).test(xml)

/** Attributes by name as written (`r:id` keeps its prefix); values still hold their entities. A scan, not a regex: it runs once per cell. */
export function attributes(text: string): Record<string, string> {
  const out: Record<string, string> = {}
  const n = text.length
  let i = 0
  for (;;) {
    const eq = text.indexOf('=', i)
    if (eq < 0) return out
    let q = eq + 1
    while (q < n && text.charCodeAt(q) <= 32) q++
    const quote = text[q]
    if (quote !== '"' && quote !== "'") return out
    const end = text.indexOf(quote, q + 1)
    if (end < 0) return out
    let b = eq
    while (b > i && text.charCodeAt(b - 1) <= 32) b--
    let a = b
    while (a > i && text.charCodeAt(a - 1) > 32) a--
    out[text.slice(a, b)] = text.slice(q + 1, end)
    i = end + 1
  }
}

/** The value of an attribute by local name, whatever its prefix (`id` finds `r:id`). */
export function attr(a: Record<string, string>, local: string): string | undefined {
  if (local in a) return a[local]
  for (const k in a) if (k.endsWith(`:${local}`)) return a[k]
  return undefined
}

/** A file's word looked up in a table by the table's own keys: `toString` or `constructor` is no word of a table. */
export const ownEntry = <T>(table: Readonly<Record<string, T>>, key: string): T | undefined => (Object.hasOwn(table, key) ? table[key] : undefined)

/** An OOXML boolean attribute (`1`, `true`, `0`, `false`); `fallback` when absent. */
export const flag = (v: string | undefined, fallback = false): boolean => (v === undefined ? fallback : v === '1' || v === 'true')
