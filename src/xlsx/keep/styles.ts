/**
 * The styles of a file written over another: styles.xml is not built again but added to. A cell format that looks as
 * one of the file's (the same style and number format, as the reader sees them) is that one, the first of them, and so
 * keeps what the model does not hold of it (theme colours, a cell style, rotation); a new one goes at the end of
 * `<cellXfs>`, with its font, fill, border and number format after the file's, its font relative to the file's default
 * one. A conditional format's differential format goes at the end of `<dxfs>`. Only what a written part uses is added.
 */
import type { CellStyle, ConditionalStyle, NumberFormat } from '../../contract.js'
import { keyOf } from '../../style-table.js'
import type { StyleTable } from '../styles/read.js'
import { attributes, element, elements } from '../xml/scan.js'
import { borderXml, DEFAULT_FONT, dxfXml, fillXml, fontXml, numFmtId, StyleWriter, xfXml, type BaseFont, type CellFormats } from '../write/styles.js'
import { decodeXml, escapeAttr } from '../xml/text.js'
import { childrenOf, rootOf, setAttrs, type Placed } from './xml.js'

/** The lists of styles.xml added to, in the order of the schema, with the elements that come before each when it is missing. */
const LISTS = ['numFmts', 'fonts', 'fills', 'borders', 'cellXfs', 'dxfs'] as const
type List = (typeof LISTS)[number]
const ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst']

/** A number format's key: its id (`numFmtId`'s), a code of the file's own by its text. */
function formatKey(f: NumberFormat | undefined): string {
  if (f === '') return 'n0'
  const id = numFmtId(f, () => -1)
  return id >= 0 ? `n${id}` : `c${f}`
}

/** A look's key: its style by its fields and its number format. */
const lookKey = (style: CellStyle | undefined, format: NumberFormat | undefined) => `${keyOf(style ?? {})}\n${formatKey(format)}`

export class KeptStyles {
  private readonly xml: string
  private readonly root: Placed
  private readonly lists: Partial<Record<List, Placed>>
  private readonly counts: Record<List, number>
  private readonly added: Record<List, string[]> = { numFmts: [], fonts: [], fills: [], borders: [], cellXfs: [], dxfs: [] }
  /** Items added by their XML: their index among the added. */
  private readonly addedIndex: Record<List, Map<string, number>> = { numFmts: new Map(), fonts: new Map(), fills: new Map(), borders: new Map(), cellXfs: new Map(), dxfs: new Map() }
  /** A look's cell format: the file's first of that look, else one added. */
  private readonly byLook = new Map<string, number>()
  /** The file's cell formats by index: the index of the first of their look. */
  private readonly first: number[] = []
  private readonly codes = new Map<string, number>()
  private nextFormat = 164
  private readonly base: BaseFont
  /** How many of the added cell formats and dxfs a written part uses: those up to the last one used are written. */
  private used = { cellXfs: 0, dxfs: 0 }
  /** Whether the file has a styles part: without one the writer adds it. */
  readonly present: boolean

  constructor(xml: string | undefined, table: StyleTable) {
    this.present = xml !== undefined
    this.xml = xml ?? new StyleWriter().xml()
    this.root = rootOf(this.xml)!
    const kids = childrenOf(this.xml, this.root)
    this.lists = Object.fromEntries(kids.filter((c) => (LISTS as readonly string[]).includes(c.local)).map((c) => [c.local, c]))
    const count = (l: List, local: string) => (this.lists[l] ? childrenOf(this.xml, this.lists[l]).filter((c) => c.local === local).length : 0)
    this.counts = { numFmts: count('numFmts', 'numFmt'), fonts: count('fonts', 'font'), fills: count('fills', 'fill'), borders: count('borders', 'border'), cellXfs: count('cellXfs', 'xf'), dxfs: count('dxfs', 'dxf') }
    for (const e of elements(this.listText('numFmts'), 'numFmt')) {
      const a = attributes(e.attrs)
      const id = Number(a.numFmtId)
      if (!Number.isInteger(id)) continue
      if (a.formatCode !== undefined && !this.codes.has(decodeXml(a.formatCode))) this.codes.set(decodeXml(a.formatCode), id)
      this.nextFormat = Math.max(this.nextFormat, id + 1)
    }
    const xfs = this.xml === xml ? table.xfs : []
    for (let i = 0; i < this.counts.cellXfs; i++) {
      const x = xfs[i]
      const key = lookKey(x?.style === undefined ? undefined : table.styles[x.style], x?.format)
      if (!this.byLook.has(key)) this.byLook.set(key, i)
      this.first.push(this.byLook.get(key)!)
    }
    // the default font a style's font is relative to: the file's first, its colour and theme scheme as written
    const font = elements(this.listText('fonts'), 'font')[0]?.body ?? ''
    const val = (local: string) => attributes(element(font, local)?.attrs ?? '').val
    this.base = {
      name: decodeXml(val('name') ?? DEFAULT_FONT.name),
      size: Number(val('sz')) || DEFAULT_FONT.size,
      color: elementText(font, 'color'),
      rest: ['family', 'charset', 'scheme'].map((l) => elementText(font, l) ?? '').join(''),
    }
  }

  private listText(l: List): string {
    const p = this.lists[l]
    return p ? this.xml.slice(p.start, p.end) : ''
  }

  private add(l: List, xml: string): number {
    let k = this.addedIndex[l].get(xml)
    if (k === undefined) this.addedIndex[l].set(xml, (k = this.added[l].push(xml) - 1))
    return this.counts[l] + k
  }

  private numFmt(code: NumberFormat | undefined): number {
    return numFmtId(code, (c) => {
      let id = this.codes.get(c)
      if (id === undefined) {
        this.codes.set(c, (id = this.nextFormat++))
        this.add('numFmts', `<numFmt numFmtId="${id}" formatCode="${escapeAttr(c)}"/>`)
      }
      return id
    })
  }

  /** The cell formats of a book's style table, the file's where they look the same. */
  of(styles: CellStyle[] | undefined): CellFormats {
    return {
      xf: (style, format) => {
        const s = style === undefined ? undefined : styles?.[style]
        const key = lookKey(s, format)
        let i = this.byLook.get(key)
        if (i !== undefined) return i
        const numFmt = format === '' ? 0 : this.numFmt(format)
        const font = s?.font ? this.add('fonts', fontXml(s.font, this.base)) : 0
        const fill = s?.fill ? this.add('fills', fillXml(s.fill)) : 0
        const border = s?.border ? this.add('borders', borderXml(s.border)) : 0
        i = this.add('cellXfs', xfXml(numFmt, font, fill, border, s))
        this.byLook.set(key, i)
        return i
      },
      dxf: (s: ConditionalStyle) => this.add('dxfs', dxfXml(s, (c) => this.numFmt(c))),
      dxfBody: (i) => this.added.dxfs[i - this.counts.dxfs].slice(5, -6),
    }
  }

  /** The cell format a written cell of the file's cell format `own` takes for the look `xf` stands for: `own` when it looks so. */
  keep(own: number | undefined, xf: number): number {
    return own !== undefined && this.first[own] === xf ? own : xf
  }

  /** A cell format or dxf a written part uses. */
  use(kind: 'cellXfs' | 'dxfs', i: number): void {
    const k = i - this.counts[kind] + 1
    if (k > this.used[kind]) this.used[kind] = k
  }

  /** styles.xml with what written parts use added; undefined when they use nothing new. */
  text(): string | undefined {
    if (!this.used.cellXfs && !this.used.dxfs) return this.present ? undefined : this.xml
    const take: Record<List, string[]> = { ...this.added, cellXfs: this.added.cellXfs.slice(0, this.used.cellXfs), dxfs: this.added.dxfs.slice(0, this.used.dxfs) }
    const edits: { at: number; end: number; text: string }[] = []
    for (const l of LISTS) {
      if (!take[l].length) continue
      const p = this.lists[l]
      const count = this.counts[l] + take[l].length
      if (p) {
        const tag = setAttrs(this.xml.slice(p.start, p.open), { count: String(count) })
        const body = p.end === p.open ? '' : this.xml.slice(p.open, p.end - p.qname.length - 3)
        edits.push({ at: p.start, end: p.end, text: `${tag.replace(/\/>$/, '>')}${body}${take[l].join('')}</${p.qname}>` })
      } else {
        // a list the file does not have goes after the last element the schema puts before it
        const rank = ORDER.indexOf(l)
        const before = childrenOf(this.xml, this.root).filter((c) => ORDER.indexOf(c.local) >= 0 && ORDER.indexOf(c.local) < rank).at(-1)
        const at = before ? before.end : this.root.open
        edits.push({ at, end: at, text: `<${l} count="${count}">${take[l].join('')}</${l}>` })
      }
    }
    edits.sort((a, b) => a.at - b.at)
    let out = ''
    let from = 0
    for (const e of edits) {
      out += this.xml.slice(from, e.at) + e.text
      from = e.end
    }
    return out + this.xml.slice(from)
  }
}

/** An element of a local name in a text, as written; undefined without one. */
function elementText(xml: string, local: string): string | undefined {
  const m = new RegExp(`<(?:[\\w.-]+:)?${local}(?:\\s(?:"[^"]*"|'[^']*'|[^>"'])*?)?(?:/>|>[\\s\\S]*?</(?:[\\w.-]+:)?${local}>)`).exec(xml)
  return m?.[0]
}

