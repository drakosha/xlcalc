/**
 * The XML a file written over another works on: the children of a part's root element where they lie, an opening
 * tag's attributes changed in place, and the relationships and content types of the package edited entry by entry.
 * What is not changed is kept as the file has it.
 */
import { attributes } from '../xml/scan.js'
import { decodeXml, escapeAttr } from '../xml/text.js'

/** An element where it lies in the text: its qualified and local names, its opening tag's end, and its end. */
export interface Placed {
  qname: string
  local: string
  start: number
  /** Past the opening tag's `>`. */
  open: number
  end: number
}

const NAME = /[\w.:-]+/y

/** The end of a tag that starts at `at` (past its `>`), quoted values passed over whole; -1 when it does not end. */
export function tagEnd(xml: string, at: number): number {
  for (let i = at + 1; i < xml.length; i++) {
    const ch = xml[i]
    if (ch === '"' || ch === "'") {
      i = xml.indexOf(ch, i + 1)
      if (i < 0) return -1
    } else if (ch === '>') return i + 1
  }
  return -1
}

/** The element whose opening tag starts at `at`: its names and where it ends, its closing tag found past those of its name inside it. */
function elementAt(xml: string, at: number): Placed | undefined {
  NAME.lastIndex = at + 1
  const m = NAME.exec(xml)
  if (!m) return undefined
  const qname = m[0]
  const local = qname.slice(qname.indexOf(':') + 1)
  const open = tagEnd(xml, at)
  if (open < 0) return undefined
  if (xml[open - 2] === '/') return { qname, local, start: at, open, end: open }
  const re = new RegExp(`<(/?)${qname.replace(/\./g, '\\.')}(?=[\\s/>])`, 'g')
  re.lastIndex = open
  let depth = 1
  for (let t = re.exec(xml); t; t = re.exec(xml)) {
    const end = tagEnd(xml, t.index)
    if (end < 0) return undefined
    if (t[1]) depth--
    else if (xml[end - 2] !== '/') depth++
    re.lastIndex = end
    if (depth === 0) return { qname, local, start: at, open, end }
  }
  return undefined
}

/** The root element of a part, past the declaration, comments and processing instructions before it. */
export function rootOf(xml: string): Placed | undefined {
  for (let i = xml.indexOf('<'); i >= 0; i = xml.indexOf('<', i + 1)) {
    if (xml.startsWith('<?', i) || xml.startsWith('<!', i)) continue
    return elementAt(xml, i)
  }
  return undefined
}

/** The child elements of an element (`parent`), in order; text, comments and processing instructions between them are left out. */
export function childrenOf(xml: string, parent: Placed): Placed[] {
  const out: Placed[] = []
  if (parent.end === parent.open) return out
  const stop = xml.lastIndexOf('</', parent.end - 1)
  const past = (i: number, end: string) => {
    const at = xml.indexOf(end, i)
    return at < 0 ? -1 : at + end.length
  }
  for (let i = xml.indexOf('<', parent.open); i >= 0 && i < stop; i = xml.indexOf('<', i)) {
    if (xml.startsWith('<!--', i)) i = past(i, '-->')
    else if (xml.startsWith('<?', i)) i = past(i, '?>')
    else if (xml.startsWith('<![CDATA[', i)) i = past(i, ']]>')
    else {
      const e = elementAt(xml, i)
      if (!e) break
      out.push(e)
      i = e.end
    }
    if (i < 0) break
  }
  return out
}

/** A tag as an opening tag, and as one that closes itself. */
export const opening = (tag: string) => (tag.endsWith('/>') ? `${tag.slice(0, -2)}>` : tag)
export const closing = (tag: string) => (tag.endsWith('/>') ? tag : `${tag.slice(0, -1)}/>`)

/** An element's opening tag, qualified name and content, apart. */
export function split(text: string): { tag: string; body: string; name: string } {
  const end = tagEnd(text, 0)
  const tag = text.slice(0, end)
  const name = /^<([\w.:-]+)/.exec(tag)![1]
  return { tag, name, body: tag.endsWith('/>') ? '' : text.slice(end, text.length - name.length - 3) }
}

/** An element with items added at the end of its content and its `count` set to `count`. */
export function appended(text: string, items: string, count: number): string {
  const e = split(text)
  return `${opening(setAttrs(e.tag, { count: String(count) }))}${e.body}${items}</${e.name}>`
}

/** An opening tag (`<x a="1">` or `<x a="1"/>`) with attributes set (a value) or removed (null), the others as written. */
export function setAttrs(tag: string, changes: Record<string, string | null>): string {
  const close = tag.endsWith('/>') ? '/>' : '>'
  let body = tag.slice(0, tag.length - close.length)
  for (const name of Object.keys(changes)) body = body.replace(new RegExp(`\\s${name.replace(/[.:]/g, '\\$&')}\\s*=\\s*("[^"]*"|'[^']*')`), '')
  for (const [name, v] of Object.entries(changes)) if (v !== null) body += ` ${name}="${escapeAttr(v)}"`
  return body + close
}

/** The attributes of an element's opening tag, as `attributes` reads them. */
export const attrsOf = (xml: string, e: Placed): Record<string, string> => attributes(xml.slice(e.start + 1 + e.qname.length, e.open - (xml[e.open - 2] === '/' ? 2 : 1)))

/** The namespaces of the package's relationships and content types. */
export const PKG_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
export const DOC_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

/** A part's relationships edited: those kept as the file has them, those removed, new ones with ids the part does not use. */
export class Rels {
  private readonly items: { id: string; xml: string }[] = []
  changed = false

  constructor(xml: string | undefined) {
    if (!xml) return
    const root = rootOf(xml)
    if (root) for (const c of childrenOf(xml, root)) if (c.local === 'Relationship') this.items.push({ id: attrsOf(xml, c).Id ?? '', xml: xml.slice(c.start, c.end) })
  }

  /** A new relationship's id: `rId` and the first number past those in use. */
  add(type: string, target: string, external = false): string {
    let n = this.items.length + 1
    const used = new Set(this.items.map((r) => r.id))
    while (used.has(`rId${n}`)) n++
    const id = `rId${n}`
    this.items.push({ id, xml: `<Relationship Id="${id}" Type="${type}" Target="${escapeAttr(target)}"${external ? ' TargetMode="External"' : ''}/>` })
    this.changed = true
    return id
  }

  remove(id: string): void {
    const i = this.items.findIndex((r) => r.id === id)
    if (i < 0) return
    this.items.splice(i, 1)
    this.changed = true
  }

  get size(): number {
    return this.items.length
  }

  xml(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${PKG_REL}">${this.items.map((r) => r.xml).join('')}</Relationships>`
  }
}

/** `[Content_Types].xml` edited: overrides of parts added and removed, a default added; the others as the file has them. */
export class ContentTypes {
  private readonly head: string
  /** Overrides by part name in lower case (`/xl/workbook.xml`), defaults by extension in lower case, each with its type and XML. */
  private readonly items: { part?: string; name?: string; ext?: string; type: string; xml: string }[] = []
  changed = false

  constructor(source: string) {
    const root = rootOf(source)
    this.head = root ? source.slice(0, root.open) : ''
    if (!root) return
    for (const c of childrenOf(source, root)) {
      const a = attrsOf(source, c)
      const type = a.ContentType ?? ''
      const name = decodeXml(a.PartName ?? '').replace(/^\//, '')
      this.items.push({ ...(c.local === 'Override' ? { part: `/${name}`.toLowerCase(), name } : { ext: (a.Extension ?? '').toLowerCase() }), type, xml: source.slice(c.start, c.end) })
    }
  }

  override(part: string, type: string): void {
    this.remove(part)
    this.items.push({ part: `/${part}`.toLowerCase(), name: part, type, xml: `<Override PartName="/${escapeAttr(part)}" ContentType="${type}"/>` })
    this.changed = true
  }

  remove(part: string): void {
    const key = `/${part}`.toLowerCase()
    const i = this.items.findIndex((x) => x.part === key)
    if (i < 0) return
    this.items.splice(i, 1)
    this.changed = true
  }

  /** A default content type by extension, unless the file has one for it. */
  default(ext: string, type: string): void {
    if (this.items.some((x) => x.ext === ext)) return
    // defaults come before the overrides, as Excel writes them
    this.items.unshift({ ext, type, xml: `<Default Extension="${ext}" ContentType="${type}"/>` })
    this.changed = true
  }

  /** The type of a part by its override or its extension's default. */
  typeOf(part: string): string | undefined {
    const key = `/${part}`.toLowerCase()
    return (this.items.find((x) => x.part === key) ?? this.items.find((x) => x.ext !== undefined && key.endsWith(`.${x.ext}`)))?.type
  }

  /** The parts the file overrides the type of with one `type` takes, as it names them. */
  parts(type: (t: string) => boolean): string[] {
    return this.items.flatMap((x) => (x.name !== undefined && type(x.type) ? [x.name] : []))
  }

  xml(): string {
    return `${this.head}${this.items.map((x) => x.xml).join('')}</Types>`
  }
}

/** Where the value of an attribute of the element at `e` (from its opening tag) lies in its opening tag; undefined without it. */
export function valueAt(xml: string, e: { start: number }, name: string): [number, number] | undefined {
  const tag = xml.slice(e.start, xml.indexOf('>', e.start) + 1)
  const m = new RegExp(`\\s${name.replace(/[.:]/g, '\\$&')}\\s*=\\s*(["'])([^"']*)\\1`).exec(tag)
  return m ? [e.start + m.index + m[0].length - 1 - m[2].length, e.start + m.index + m[0].length - 1] : undefined
}
