/**
 * Elements and attributes of XML parts by regular expressions, without a DOM: a tag may carry a namespace
 * prefix (`<x:c>`), attributes may be quoted either way and hold `>`, and a self-closing element is matched as
 * such (it never swallows the next element). Elements of the same name are not nested in the parts read here.
 */

/** Attributes of a tag up to its end, quoted values skipped whole. */
const ATTRS = `(?:\\s(?:"[^"]*"|'[^']*'|[^>"'])*?)?`

export interface Element {
  /** The attribute text of the opening tag. */
  attrs: string
  /** The content, null for a self-closing element. */
  body: string | null
}

const cache = new Map<string, RegExp>()

function elementRe(local: string): RegExp {
  let re = cache.get(local)
  if (!re) cache.set(local, (re = new RegExp(`<(?:[\\w.-]+:)?${local}(${ATTRS})(?:/>|>([\\s\\S]*?)</(?:[\\w.-]+:)?${local}\\s*>)`, 'g')))
  return re
}

/** Every element of a local name, in document order. */
export function elements(xml: string, local: string): Element[] {
  const out: Element[] = []
  for (const m of xml.matchAll(elementRe(local))) out.push({ attrs: m[1], body: m[2] ?? null })
  return out
}

/** The first element of a local name, or undefined. */
export function element(xml: string, local: string): Element | undefined {
  const re = elementRe(local)
  re.lastIndex = 0
  const m = re.exec(xml)
  re.lastIndex = 0
  return m ? { attrs: m[1], body: m[2] ?? null } : undefined
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

/** An OOXML boolean attribute (`1`, `true`, `0`, `false`); `fallback` when absent. */
export const flag = (v: string | undefined, fallback = false): boolean => (v === undefined ? fallback : v === '1' || v === 'true')
