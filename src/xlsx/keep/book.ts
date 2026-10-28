/**
 * `writeKept`: a book written over the file it was read from. The parts the book did not change are the file's, byte
 * for byte; a sheet's part keeps what the model does not have (`sheet.ts`); styles are added to (`styles.ts`) and new
 * text goes inline, so the shared strings stay; the calculation chain goes, as Excel builds it again. A change that
 * would move references the model does not hold refuses the write (`refusals.ts`). A book unchanged gives the file as
 * it is.
 */
import type { BookInput, SheetInput, TableInput } from '../../contract.js'
import { caseKey } from '../../collate.js'
import { SHEET_STATES } from '../../input.js'
import { object, onlyOptions } from '../../input/basic.js'
import { sheetKey } from '../../syntax/print.js'
import { Issues } from '../issues.js'
import { checkLimits, fileBytes, readStyled } from '../read/book.js'
import { partText, relsOf, type Rel } from '../read/package.js'
import type { XlsxSnapshot, XlsxWritten } from '../types.js'
import { bookWriting, filterDatabase, calcAttrs, nameElements, RICH_REL, type BookWriting } from '../write/book.js'
import { nextBlock } from '../write/parts.js'
import { sheetPieces, sheetXml, type SheetContext, type SheetPieces } from '../write/sheet.js'
import { writeArgs, type ResultCell } from '../write/snapshot.js'
import { tableXml } from '../write/tables.js'
import { attr, attributes, element } from '../xml/scan.js'
import { escapeAttr, readText } from '../xml/text.js'
import { parseRange } from '../../structure/area.js'
import { areasOf } from './refs.js'
import { utf8 } from '../zip/platform.js'
import { zip } from '../zip/write.js'
import { KeptIssues } from './issues.js'
import { keyOf } from '../../style-table.js'
import { externalLinkXml } from '../write/links.js'
import { records } from './zip.js'
import { carrierNames, carriersOf, type Carriers, type Edit } from './carriers.js'
import { keptRules } from './rules.js'
import { conditionalXml } from '../write/conditional.js'
import { edited, KeptRefusal, refusalIn, refusals, type Identity } from './refusals.js'
import { spliceSheet } from './sheet.js'
import { contentTypes, openSource, relsPath, type Source, type SourceSheet } from './source.js'
import { addMeta, planMeta, type MetaPlan } from './metadata.js'
import { Formats } from './formats.js'
import { KeptStyles } from './styles.js'
import type { WriteKeptOptions } from './types.js'
import { attrsOf, childrenOf, DOC_REL, Rels, rootOf, setAttrs, type ContentTypes, type Placed } from './xml.js'

const TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml'
const VML = 'application/vnd.openxmlformats-officedocument.vmlDrawing'
const STRICT = 'http://purl.oclc.org/ooxml/spreadsheetml/main'
const REL = (kind: string) => `${DOC_REL}/${kind}`
/** The children of a workbook the schema puts before its names, and before its calculation settings. */
const BEFORE_NAMES = ['fileVersion', 'fileSharing', 'workbookPr', 'workbookProtection', 'bookViews', 'sheets', 'functionGroups', 'externalReferences']
const BEFORE_CALC = [...BEFORE_NAMES, 'definedNames']

/** The parts written: new text by part name in lower case, those removed, those added in order. */
class Parts {
  readonly text = new Map<string, string>()
  readonly removed = new Set<string>()
  readonly added: string[] = []
  private readonly names: Set<string>
  constructor(src: Source) {
    this.names = new Set(src.archive.names.map((n) => n.toLowerCase()))
  }
  has(part: string): boolean {
    const k = part.toLowerCase()
    return this.text.has(k) || (this.names.has(k) && !this.removed.has(k))
  }
  set(part: string, text: string): void {
    const k = part.toLowerCase()
    if (!this.names.has(k) && !this.text.has(k)) this.added.push(part)
    this.removed.delete(k)
    this.text.set(k, text)
  }
  remove(part: string): void {
    const k = part.toLowerCase()
    this.text.delete(k)
    const i = this.added.findIndex((p) => p.toLowerCase() === k)
    if (i >= 0) this.added.splice(i, 1)
    else if (this.names.has(k)) this.removed.add(k)
  }
  /** A new part's name: `prefix`, the first number past those taken, `ext`. */
  fresh(prefix: string, ext: string): string {
    for (let n = 1; ; n++) if (!this.has(`${prefix}${n}${ext}`)) return `${prefix}${n}${ext}`
  }
}

/** A part's target relative to the part that points at it (`xl/worksheets/sheet1.xml` → `../tables/table1.xml`). */
function relative(from: string, to: string): string {
  const a = from.split('/').slice(0, -1)
  const b = to.split('/')
  let i = 0
  while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++
  return [...a.slice(i).map(() => '..'), ...b.slice(i)].join('/')
}

/** What the file gives every sheet: its styles, the marks of its cell metadata, the first free block of VML shape ids, the text and largest id of its tables. */
interface Setup {
  /** The parts each part points at, by its name in lower case: the file's relationships, internal ones (less those `writeKept` removes). */
  rels: Map<string, string[]>
  /** The parts the moved references are written back to, by name (the sheets and the workbook apart): their text and relationships. */
  texts: Map<string, { text: string; rels: Rel[]; relsXml: string | undefined }>
  styles: KeptStyles
  meta: MetaPlan
  /** The file's metadata and rich value parts, by relationship kind: their text. */
  metaParts: Record<string, string | undefined>
  block: number
  tables: Map<string, string>
  tableId: number
}

/** The relationship kinds of the cell metadata and the rich values of errors. */
const META = ['sheetMetadata', 'rdRichValue', 'rdRichValueStructure'] as const

async function setup(src: Source, edits: Edit[], styles: KeptStyles): Promise<Setup> {
  const part = (kind: string) => partText(src.archive, src.wbRels, kind)
  const [metaXml, richValues, richStructures] = await Promise.all(META.map(part))
  let block = 1
  // the notes' VML shape ids go in blocks past those the file's drawings take
  for (const name of src.archive.names)
    if (/\.vml$/i.test(name)) for (const m of ((await src.archive.text(name)) ?? '').matchAll(/data="([\d,\s]+)"/g)) for (const b of m[1].split(',')) block = Math.max(block, Number(b) + 1)
  const tables = new Map<string, string>()
  let tableId = 0
  for (const s of src.sheets)
    for (const r of s.rels)
      if (r.kind === 'table' && !r.external) {
        const xml = await src.archive.text(r.target)
        if (xml === undefined) continue
        tables.set(r.target, xml)
        tableId = Math.max(tableId, Number(attributes(element(xml, 'table')?.attrs ?? '').id) || 0)
      }
  const rels = new Map<string, string[]>()
  for (const name of src.archive.names) {
    const m = /^(.*?)_rels\/([^/]+)\.rels$/i.exec(name)
    if (m) rels.set(`${m[1]}${m[2]}`.toLowerCase(), (await relsOf(src.archive, `${m[1]}${m[2]}`)).filter((r) => !r.external).map((r) => r.target.toLowerCase()))
  }
  const own = new Set([src.wbPath, ...src.sheets.map((s) => s.part)].map((p) => p.toLowerCase()))
  const texts: Setup['texts'] = new Map()
  for (const part of new Set(edits.map((e) => e.part)))
    if (!own.has(part.toLowerCase())) texts.set(part, { text: (await src.archive.text(part))!, rels: await relsOf(src.archive, part), relsXml: await src.archive.text(relsPath(part)) })
  return {
    rels,
    texts,
    styles,
    meta: planMeta(metaXml, richValues, richStructures),
    metaParts: { sheetMetadata: metaXml, rdRichValue: richValues, rdRichValueStructure: richStructures },
    block,
    tables,
    tableId,
  }
}

/**
 * Writes a book over the xlsx file it was read from (`openKept`): what the book changed is written, the rest is the
 * file's as it was, the parts the model does not have among it. Refuses (`KeptRefusal`, its message in the language of
 * `opts.locale`) a change that would leave references the model does not hold pointing at other cells.
 */
export async function writeKept(original: Uint8Array | ArrayBuffer, snapshot: XlsxSnapshot, opts?: WriteKeptOptions): Promise<XlsxWritten> {
  try {
    return await writeOver(original, snapshot, opts)
  } catch (e) {
    // a refusal comes once the options are checked: in the language of their locale
    throw e instanceof KeptRefusal ? refusalIn(e, opts?.locale) : e
  }
}

/** `writeKept`'s writing, its refusals in English. */
async function writeOver(original: Uint8Array | ArrayBuffer, snapshot: XlsxSnapshot, opts?: WriteKeptOptions): Promise<XlsxWritten> {
  const bytes = fileBytes(original)
  const o = opts === undefined ? {} : object(opts, 'opts', "options such as { values: 'preserve' }")
  onlyOptions(o, 'opts', ['values', 'locale', 'limits'])
  const limits = o.limits === undefined ? undefined : checkLimits(o.limits)
  const { book: given, cells, locale, unasked } = writeArgs(snapshot, { ...(o.values !== undefined ? { values: o.values } : {}), ...(o.locale !== undefined ? { locale: o.locale } : {}) })
  const src = await openSource(bytes, limits)
  const { read, stylesXml, table } = await readStyled(bytes, limits ? { limits } : undefined)
  const styles = new KeptStyles(stylesXml, table)
  const carriers = await carriersOf(src, read.book, styles)
  const identity = refusals(src, read.book, given, carriers)
  // the book as written: its names without the carriers
  const carried = new Set(carrierNames(carriers).map((n) => caseKey(n.name)))
  const book: BookInput = given.names ? { ...given, names: given.names.filter((n) => n.scope !== undefined || !carried.has(caseKey(n.name))) } : given
  // the sheets whose rules the reader could not all take, which writing them again from the workbook would lose
  const unread = (code: string) => new Set(read.issues.flatMap((x) => (x.code === code && x.sheet !== undefined ? [x.sheet] : [])))
  const lost = { conditional: unread('conditionalFormat'), validation: unread('dataValidation') }
  const k = new Kept(src, read.book, book, cells, cells && read.cached.sheets.map((s) => s.cells), identity, await setup(src, identity.edits, styles), lost, carriers.rules)
  const root = rootOf(src.wb)!
  const strict = src.wb.slice(root.start, root.open).includes(STRICT)
  try {
    k.write()
  } catch (e) {
    // a Strict file's sheets are in other namespaces than the model's elements
    throw strict ? new KeptRefusal({ key: 'keptStrict' }, 'file') : e
  }
  if (!k.changed) return { data: bytes, issues: [] }
  if (strict) throw new KeptRefusal({ key: 'keptStrict' }, 'file')
  // a sheet the reader read under another name than the file's: the file's formulas name it as the file does
  const repaired = src.sheets.findIndex((s, i) => src.decls[s.decl].name !== read.book.sheets[i].name)
  if (repaired >= 0) throw new KeptRefusal({ key: 'keptSheetName', name: src.decls[src.sheets[repaired].decl].name, sheet: read.book.sheets[repaired].name }, 'file')
  if (k.formulas && unasked) k.issues.add('noCachedValues', { key: 'noResult' })
  // a sheet with more cells that hold what only the file holds than have carriers, whose rows or columns moved
  for (const [i, count] of carriers.uncarried.entries()) {
    const j = identity.target.get(i)
    if (count && j !== undefined && identity.lines[i]) k.issues.add('richText', { key: 'keptCells', count }, { sheet: book.sheets[j].name })
  }
  // a sheet with more rectangles of cell formats that look as earlier ones than have carriers, whose rows or columns moved
  for (const [i, count] of carriers.unformatted.entries()) {
    const j = identity.target.get(i)
    if (count && j !== undefined && identity.lines[i]) k.issues.add('style', { key: 'keptFormats', count }, { sheet: book.sheets[j].name })
  }
  return { data: await k.zip(), issues: k.issues.list(locale) }
}

/** The writing of one book over one file. */
class Kept {
  readonly issues = new Issues()
  readonly parts: Parts
  readonly types: ContentTypes
  readonly wbRels: Rels
  /** Whether the book changed what the file holds (more than the values of formulas). */
  changed = false
  /** Whether a part written holds a formula of the book. */
  formulas = false
  private readonly w0: BookWriting
  private readonly w1: BookWriting
  /** The workbook's relationship id of each sheet of the book written. */
  private readonly sheetRels: string[] = []
  private block: number
  private tableId: number
  /** The file's parts removed, by name in lower case. */
  private readonly gone = new Set<string>()

  constructor(
    readonly src: Source,
    readonly book0: BookInput,
    readonly book: BookInput,
    readonly cells: Record<string, ResultCell>[] | undefined,
    readonly cached: Record<string, ResultCell>[] | undefined,
    readonly id: Identity,
    readonly setup: Setup,
    /** By name, the sheets of the file with conditional formats, and with data validation, the reader could not all take. */
    readonly lost: { conditional: Set<string>; validation: Set<string> },
    /** By sheet of the file: its conditional formats as the reader met them. */
    readonly rules0: Carriers['rules'],
  ) {
    this.parts = new Parts(src)
    this.types = contentTypes(src)
    this.wbRels = new Rels(src.wbRelsXml)
    this.w0 = bookWriting(book0)
    this.w1 = bookWriting(book)
    this.block = setup.block
    this.tableId = setup.tableId
  }

  get styles(): KeptStyles {
    return this.setup.styles
  }

  private context(book: BookInput, w: BookWriting, i: number, result: Record<string, ResultCell> | undefined, issues: Issues): SheetContext {
    return {
      styles: this.styles.of(book.styles),
      string: () => {
        throw new Error('xlsx/keep: text goes inline')
      },
      inline: true,
      dynamicMark: this.setup.meta.dynamic.mark,
      issues,
      ...w.names,
      mode: 'formulas',
      ...(result ? { result } : {}),
      richError: (code) => this.setup.meta.rich[code].vm,
      tableAt: w.tableAt(i),
      tableParts: [],
      epoch: w.epoch,
      ...(w.filterTables[i] ? { filterTable: w.filterTables[i] } : {}),
      noteBlock: this.block,
      active: w.active === i,
    }
  }

  write(): void {
    const { src, book0, book } = this
    // the file's first font is the default its styles are written against: another one would change every cell's look
    if (JSON.stringify(book.font ?? null) !== JSON.stringify(book0.font ?? null)) throw new KeptRefusal({ key: 'keptFont' }, 'font')
    const emptied = this.others()
    for (const [j, s1] of book.sheets.entries()) {
      const k = this.id.source[j]
      const issues = new KeptIssues()
      if (k === undefined) {
        this.newSheet(j, s1, issues)
        continue
      }
      const [s0, ss] = [book0.sheets[k], src.sheets[k]]
      const d = src.decls[ss.decl]
      this.sheetRels[j] = attr(d.attrs, 'id') ?? ''
      // the sheet's index in both writes is the file's: what a writer numbers by it (rules' ids) compares
      const p0 = sheetPieces(s0, k, this.context(book0, this.w0, k, this.cached?.[k], new Issues()))
      const p1 = sheetPieces(s1, k, this.context(book, this.w1, j, this.cells?.[j], issues))
      // rules the workbook does not hold go with a sheet's rules written again (`rules.ts`); validation is refused, never lost
      const name = book0.sheets[k].name
      const conditional = this.lost.conditional.has(name) ? this.rules(k, j, p0, p1) : undefined
      if (this.lost.validation.has(name) && p0.validations !== p1.validations)
        throw new KeptRefusal({ key: 'keptValidation', sheet: name }, 'rules')
      const rels = new Rels(ss.relsXml)
      const bound = { hyperlinks: this.hyperlinks(ss, p0, p1, rels), legacyDrawing: this.notes(ss, p0, p1, rels), tableParts: this.tables(ss, k, j, rels) }
      // the references of the sheet's own parts (an AutoFilter, page breaks, sparklines) written back first
      let xml = this.edited(ss.part, ss.xml)
      const drawing = ss.rels.find((r) => r.kind === 'drawing' && !r.external && emptied.has(r.target.toLowerCase()))
      if (drawing) {
        // a drawing left without anchors goes, with the sheet's `<drawing>`
        const el = childrenOf(xml, rootOf(xml)!).find((c) => c.local === 'drawing')
        if (el) xml = xml.slice(0, el.start) + xml.slice(el.end)
        rels.remove(drawing.id)
        this.unlink(ss.part, drawing.target)
      }
      const moved = this.id.cells.get(k)
      const formats = this.id.formats.get(k)
      const out = spliceSheet({ xml, p0, p1, s0, s1, styles: this.styles, bound, lines: this.id.lines[k], ...(moved ? { moved } : {}), ...(formats ? { formats: new Formats(formats, this.styles) } : {}), ...(conditional ? { conditional } : {}) })
      this.changed ||= out.changed || xml !== ss.xml
      if (out.xml !== undefined || xml !== ss.xml) this.parts.set(ss.part, out.xml ?? xml)
      if (out.xml !== undefined) this.formulas ||= p1.formulas
      if (rels.changed) this.parts.set(relsPath(ss.part), rels.xml())
      adopt(issues, this.issues, out.cells, out.written)
    }
    // the sheets the book no longer has, with the parts of their own
    for (const [k, ss] of src.sheets.entries()) if (!this.id.target.has(k)) this.deleteSheet(ss)
    this.workbook()
  }

  /**
   * The conditional formats of the file's sheet `k` (the book's `j`) where they changed and the file has rules the model
   * does not have, or read without a part of their look: the workbook's with the file's (`keptRules`); undefined where
   * neither the workbook's rules nor the carried ones changed.
   */
  private rules(k: number, j: number, p0: SheetPieces, p1: SheetPieces): { main: string; ext: string } | undefined {
    const ss = this.src.sheets[k]
    const sources = this.rules0[k] ?? []
    const part = ss.part.toLowerCase()
    const edits = this.id.edits.filter((e) => e.part.toLowerCase() === part && sources.some((s) => e.start >= s.block.start && e.end <= s.block.end))
    if (p0.conditional + p0.ext === p1.conditional + p1.ext && !edits.length) return undefined
    const [s0, s1] = [this.book0.sheets[k], this.book.sheets[j]]
    const ctx = this.context(this.book, this.w1, j, this.cells?.[j], new Issues())
    return keptRules({ xml: ss.xml, sheet: s0.name, sources, edits, deleted: this.id.deleted[k], r0: s0.conditionalFormats ?? [], r1: s1.conditionalFormats ?? [], write: (priority) => conditionalXml(s1, k, { ...ctx, priority }) })
  }

  /**
   * The parts the model does not have with their moved references written back: charts, drawings, pivot tables and
   * caches. A drawing's relationships its anchors no longer use go, with the charts and pictures only they point at, as
   * Excel deletes a drawing with its cells; gives the drawings left without anchors, which go with their sheets'
   * `<drawing>`.
   */
  private others(): Set<string> {
    const emptied = new Set<string>()
    const drawings: [string, string][] = []
    for (const [part, { text }] of this.setup.texts) {
      const next = this.edited(part, text)
      if (next === text) continue
      this.changed = true
      this.parts.set(part, next)
      if (rootOf(next)?.local === 'wsDr') drawings.push([part, next])
    }
    for (const [part, next] of drawings) {
      const { rels: list, relsXml } = this.setup.texts.get(part)!
      if (!childrenOf(next, rootOf(next)!).length) {
        emptied.add(part.toLowerCase())
        continue
      }
      const rels = new Rels(relsXml)
      for (const r of list)
        if (!r.external && !next.includes(`"${escapeAttr(r.id)}"`)) {
          rels.remove(r.id)
          this.unlink(part, r.target)
        }
      if (rels.changed) this.parts.set(relsPath(part), rels.xml())
    }
    return emptied
  }

  /** A part removed with its relationships, and the parts only it pointed at, theirs too. */
  private drop(part: string): void {
    const key = part.toLowerCase()
    if (this.gone.has(key)) return
    this.gone.add(key)
    this.parts.remove(part)
    this.types.remove(part)
    this.parts.remove(relsPath(part))
    for (const t of this.setup.rels.get(key) ?? []) this.orphan(t)
  }

  /** A part (its name in lower case) removed where no part kept points at it. */
  private orphan(target: string): void {
    const kept = [...this.setup.rels].some(([from, to]) => !this.gone.has(from) && to.includes(target))
    if (!kept) this.drop(this.src.archive.names.find((n) => n.toLowerCase() === target) ?? target)
  }

  /** A relationship of a part removed, and its target where no part kept points at it any more. */
  private unlink(part: string, target: string): void {
    const [from, to] = [part.toLowerCase(), target.toLowerCase()]
    const list = this.setup.rels.get(from) ?? []
    const i = list.indexOf(to)
    if (i >= 0) list.splice(i, 1)
    this.orphan(to)
  }

  /** A part's text with the moved references written back. */
  private edited(part: string, text: string): string {
    const key = part.toLowerCase()
    const mine = this.id.edits.filter((e) => e.part.toLowerCase() === key)
    return mine.length ? edited(text, mine) : text
  }

  /** The sheet's hyperlinks where they changed, their addresses relationships of the sheet in place of the file's. */
  private hyperlinks(ss: SourceSheet, p0: SheetPieces, p1: SheetPieces, rels: Rels): string | undefined {
    const shown = (p: SheetPieces) => p.hyperlinks.replace(/r:id="rId(\d+)"/g, (_, n) => `url="${escapeAttr(p.rels[Number(n) - 1]?.target ?? '')}"`)
    if (shown(p0) === shown(p1)) return undefined
    this.changed = true
    for (const r of ss.rels) if (r.kind === 'hyperlink') rels.remove(r.id)
    const ids = new Map<string, string>()
    return p1.hyperlinks.replace(/r:id="rId(\d+)"/g, (_, n) => {
      const url = p1.rels[Number(n) - 1].target
      let id = ids.get(url)
      if (id === undefined) ids.set(url, (id = rels.add(REL('hyperlink'), url, true)))
      return `r:id="${id}"`
    })
  }

  /** The sheet's notes where they changed: its comments part and VML drawing written, added or removed, and its `<legacyDrawing>` when it changes. */
  private notes(ss: SourceSheet, p0: SheetPieces, p1: SheetPieces, rels: Rels): string | undefined {
    if (p0.notes?.comments === p1.notes?.comments && p0.notes?.shown === p1.notes?.shown) return undefined
    const name = this.book0.sheets[this.src.sheets.indexOf(ss)].name
    if (ss.rels.some((r) => r.kind === 'threadedComment')) throw new KeptRefusal({ key: 'keptThreads', sheet: name }, 'threads')
    this.changed = true
    const comments = ss.rels.find((r) => r.kind === 'comments' && !r.external)
    const legacy = /<(?:[\w.-]+:)?legacyDrawing\b[^>]*?\bid="([^"]+)"/.exec(ss.xml)?.[1]
    const vml = legacy === undefined ? undefined : ss.rels.find((r) => r.id === legacy && !r.external)
    if (!p1.notes) {
      if (comments) {
        this.parts.remove(comments.target)
        this.types.remove(comments.target)
        rels.remove(comments.id)
      }
      if (vml) {
        this.parts.remove(vml.target)
        rels.remove(vml.id)
      }
      return legacy === undefined ? undefined : ''
    }
    this.block = nextBlock(this.block, p1.notes.count)
    const cPart = comments?.target ?? this.parts.fresh('xl/comments', '.xml')
    this.parts.set(cPart, p1.notes.comments)
    if (!comments) {
      this.types.override(cPart, `${TYPE}.comments+xml`)
      rels.add(REL('comments'), relative(ss.part, cPart))
    }
    const vPart = vml?.target ?? this.parts.fresh('xl/drawings/vmlDrawing', '.vml')
    this.parts.set(vPart, p1.notes.vml)
    if (vml) return undefined
    this.types.default('vml', VML)
    return `<legacyDrawing r:id="${rels.add(REL('vmlDrawing'), relative(ss.part, vPart))}"/>`
  }

  /** The sheet's tables: a part written again where its table changed, added for a new one, removed with its table; the `<tableParts>` when their list changed. */
  private tables(ss: SourceSheet, k: number, j: number, rels: Rels): string | undefined {
    const [s0, s1] = [this.book0.sheets[k], this.book.sheets[j]]
    const left = [...(s1.tables ?? [])]
    const before = s0.tables ?? []
    // a table by its name, else one renamed: a new name over the same range
    const take = (t0: TableInput): TableInput | undefined => {
      let i = left.findIndex((t) => caseKey(t.name) === caseKey(t0.name))
      if (i < 0) i = left.findIndex((t) => t.range === t0.range && !before.some((x) => caseKey(x.name) === caseKey(t.name)))
      return i < 0 ? undefined : left.splice(i, 1)[0]
    }
    const list: string[] = []
    let listChanged = false
    for (const r of ss.rels) {
      if (r.kind !== 'table' || r.external) continue
      const xml = this.setup.tables.get(r.target)
      const a = attributes(element(xml ?? '', 'table')?.attrs ?? '')
      const t0 = xml === undefined ? undefined : before.find((t) => caseKey(t.name) === caseKey(readText(a.displayName ?? a.name ?? '')))
      // a table the model left out stays as the file has it
      const t1 = t0 && take(t0)
      if (!t0) list.push(r.id)
      else if (!t1) {
        if (this.parts.has(relsPath(r.target))) throw new KeptRefusal({ key: 'keptTableParts', table: t0.name, path: relsPath(r.target) }, 'deleted')
        this.changed = listChanged = true
        this.parts.remove(r.target)
        this.types.remove(r.target)
        rels.remove(r.id)
      } else {
        list.push(r.id)
        const id = Number(a.id)
        const after = tableXml(t1, s1, id, this.w1.names, this.issues, t1.name === this.w1.filterTables[j])
        if (tableXml(t0, s0, id, this.w0.names, new Issues(), t0.name === this.w0.filterTables[k]) === after) continue
        this.changed = true
        this.parts.set(r.target, after)
      }
    }
    for (const t of left) {
      this.changed = listChanged = true
      const part = this.parts.fresh('xl/tables/table', '.xml')
      this.parts.set(part, tableXml(t, s1, ++this.tableId, this.w1.names, this.issues, t.name === this.w1.filterTables[j]))
      this.types.override(part, `${TYPE}.table+xml`)
      list.push(rels.add(REL('table'), relative(ss.part, part)))
    }
    if (!listChanged) return undefined
    return list.length ? `<tableParts count="${list.length}">${list.map((id) => `<tablePart r:id="${id}"/>`).join('')}</tableParts>` : ''
  }

  /** A sheet the file does not have, written as `writeXlsx` writes one, with its tables, notes and hyperlinks. */
  private newSheet(j: number, s: SheetInput, issues: KeptIssues): void {
    this.changed = true
    const part = this.parts.fresh('xl/worksheets/sheet', '.xml')
    const rels = new Rels(undefined)
    const tableParts: string[] = []
    for (const t of s.tables ?? []) {
      const tp = this.parts.fresh('xl/tables/table', '.xml')
      this.parts.set(tp, tableXml(t, s, ++this.tableId, this.w1.names, issues, t.name === this.w1.filterTables[j]))
      this.types.override(tp, `${TYPE}.table+xml`)
      tableParts.push(rels.add(REL('table'), relative(part, tp)))
    }
    const p = sheetPieces(s, j, { ...this.context(this.book, this.w1, j, this.cells?.[j], issues), tableParts })
    for (const r of p.rows) for (const m of `${r.attrs}${r.cells}`.matchAll(/ s="(\d+)"/g)) this.styles.use('cellXfs', Number(m[1]))
    for (const m of p.cols.matchAll(/ style="(\d+)"/g)) this.styles.use('cellXfs', Number(m[1]))
    for (const m of p.conditional.matchAll(/ dxfId="(\d+)"/g)) this.styles.use('dxfs', Number(m[1]))
    // the sheet's relationships after its tables', in the order its elements number them
    for (const r of p.rels) {
      if (r.type === 'hyperlink') rels.add(REL('hyperlink'), r.target, true)
      else if (r.type === 'comments') {
        const cp = this.parts.fresh('xl/comments', '.xml')
        this.parts.set(cp, p.notes!.comments)
        this.types.override(cp, `${TYPE}.comments+xml`)
        rels.add(REL('comments'), relative(part, cp))
      } else {
        const vp = this.parts.fresh('xl/drawings/vmlDrawing', '.vml')
        this.parts.set(vp, p.notes!.vml)
        this.types.default('vml', VML)
        rels.add(REL('vmlDrawing'), relative(part, vp))
      }
    }
    if (p.notes) this.block = nextBlock(this.block, p.notes.count)
    this.parts.set(part, sheetXml(p))
    if (rels.size) this.parts.set(relsPath(part), rels.xml())
    this.types.override(part, `${TYPE}.worksheet+xml`)
    this.sheetRels[j] = this.wbRels.add(REL('worksheet'), relative(this.src.wbPath, part))
    this.formulas ||= p.formulas
    adopt(issues, this.issues)
  }

  /**
   * A sheet deleted, with its relationships and the parts only it points at, theirs too: its tables, notes, printer
   * settings, and its drawings with their charts and pictures.
   */
  private deleteSheet(ss: SourceSheet): void {
    this.changed = true
    this.drop(ss.part)
    const id = attr(this.src.decls[ss.decl].attrs, 'id')
    if (id) this.wbRels.remove(id)
  }

  /** The workbook part: its sheets, names, active sheet and calculation settings, where the book changed them. */
  private workbook(): void {
    const { src, book0, book } = this
    const { decls } = src
    // the moved references of the names the model does not take written back: after the sheets, which keep their places
    const wb = this.edited(src.wbPath, src.wb)
    const root = rootOf(wb)!
    const kids = childrenOf(wb, root)
    const edits: Edit[] = []
    const insertAt = (before: string[]) => [...kids].reverse().find((c) => before.includes(c.local))?.end ?? root.open
    // the workbook's sheets in the book's order; the file's others (chart sheets) after the sheet they followed
    const others = new Map<number, number[]>()
    let after = -1
    for (const [i, d] of decls.entries()) {
      if (d.sheet !== undefined) after = this.id.target.get(d.sheet) ?? after
      else (others.get(after) ?? others.set(after, []).get(after)!).push(i)
    }
    const order: string[] = []
    /** By `<sheet>` of the file: its place now, none for one deleted. */
    const placeOf = new Map<number, number>()
    /** By sheet of the book written: its place. */
    const place = new Map<string, number>()
    let sheetId = Math.max(0, ...decls.map((d) => Number(d.attrs.sheetId) || 0))
    const others_ = (j: number) => {
      for (const i of others.get(j) ?? []) {
        placeOf.set(i, order.length)
        order.push(wb.slice(decls[i].at.start, decls[i].at.end))
      }
    }
    others_(-1)
    for (const [j, s] of book.sheets.entries()) {
      const k = this.id.source[j]
      const state = s.hidden ? SHEET_STATES.get(s.hidden)! : null
      place.set(sheetKey(s.name), order.length)
      if (k === undefined) order.push(`<sheet name="${escapeAttr(s.name)}" sheetId="${++sheetId}"${state ? ` state="${state}"` : ''} r:id="${this.sheetRels[j]}"/>`)
      else {
        const di = src.sheets[k].decl
        placeOf.set(di, order.length)
        const was = wb.slice(decls[di].at.start, decls[di].at.end)
        const s0 = book0.sheets[k]
        order.push(s0.name === s.name && s0.hidden === s.hidden ? was : setAttrs(was, { name: s.name, state }))
      }
      others_(j)
    }
    const moved = order.length !== decls.length || decls.some((_, i) => placeOf.get(i) !== i)
    const list = kids.find((c) => c.local === 'sheets')!
    const sheets = `${wb.slice(list.start, list.open)}${order.join('')}</${list.qname}>`
    if (sheets !== wb.slice(list.start, list.end)) {
      this.changed = true
      edits.push({ part: src.wbPath, start: list.start, end: list.end, text: sheets })
    }
    this.names(wb, kids, place, placeOf, edits, insertAt)
    this.links(wb, kids, edits, insertAt)
    // the active sheet, by its place now
    const views = kids.find((c) => c.local === 'bookViews')
    const view = views && childrenOf(wb, views).find((c) => c.local === 'workbookView')
    if (view && (moved || this.id.source[this.w1.active] !== this.w0.active || book.activeSheet !== book0.activeSheet)) {
      const tag = wb.slice(view.start, view.open)
      const active = this.w1.active < 0 ? 0 : place.get(sheetKey(book.sheets[this.w1.active].name))!
      const first = Number(attributes(tag.slice(tag.indexOf(' '))).firstSheet ?? 0)
      // the active sheet written as writeXlsx writes it: named when the book names one or it is not the first
      const next = setAttrs(tag, { activeTab: active || book.activeSheet !== undefined ? String(active) : null, ...(first >= order.length ? { firstSheet: null } : {}) })
      if (next !== tag) {
        this.changed = true
        edits.push({ part: src.wbPath, start: view.start, end: view.open, text: next })
      }
    }
    // the calculation option and iterative calculation as the book has them, and a full calculation on loading: Excel computes the values again
    const calc = kids.find((c) => c.local === 'calcPr')
    const it1 = calcAttrs(book)
    if (calcAttrs(book0) !== it1) this.changed = true
    if (!this.changed) return
    const a = attributes(it1)
    const full = this.formulas || book.sheets.some((s) => Object.values(s.cells).some((c) => c.formula !== undefined))
    const set: Record<string, string | null> = { calcMode: a.calcMode ?? null, iterate: a.iterate ?? null, iterateCount: a.iterateCount ?? null, iterateDelta: a.iterateDelta ?? null, ...(full ? { fullCalcOnLoad: '1' } : {}) }
    if (calc) edits.push({ part: src.wbPath, start: calc.start, end: calc.open, text: setAttrs(wb.slice(calc.start, calc.open), set) })
    else {
      const at = insertAt(BEFORE_CALC)
      edits.push({ part: src.wbPath, start: at, end: at, text: `<calcPr${full ? ' fullCalcOnLoad="1"' : ''}${it1}/>` })
    }
    this.parts.set(src.wbPath, edited(wb, edits))
    this.finish()
  }

  /**
   * The links to other workbooks (the refusals keep the file's at their numbers): a link of the file whose sheets, names
   * or stored values the book changed is written again as `writeXlsx` writes it, through the relationship it has; a new
   * one after them gets a part of its own, its relationship to the other file, and a `<externalReference>` of the workbook.
   * A link of the file that is no workbook's (DDE, OLE) is refused when the book changed it.
   */
  private links(wb: string, kids: Placed[], edits: Edit[], insertAt: (before: string[]) => number): void {
    const [l0, l1] = [this.book0.externalBooks ?? [], this.book.externalBooks ?? []]
    if (keyOf(l0) === keyOf(l1)) return
    const list = kids.find((c) => c.local === 'externalReferences')
    for (const [i, b] of l0.entries()) {
      if (keyOf(b) === keyOf(l1[i])) continue
      const was = this.src.links[i]
      // the relationship by which the file's part points at the other workbook
      const id = was && /<externalBook\b[^>]*?\sr:id="([^"]*)"/.exec(was.xml)?.[1]
      if (!was || !id) throw new KeptRefusal({ key: 'keptLinkKind', link: i + 1 }, 'links')
      this.parts.set(was.part, externalLinkXml(l1[i], i + 1, this.issues).replace(' r:id="rId1"', ` r:id="${id}"`))
      this.changed = true
    }
    const added: string[] = []
    for (let i = l0.length; i < l1.length; i++) {
      const part = this.parts.fresh('xl/externalLinks/externalLink', '.xml')
      this.parts.set(part, externalLinkXml(l1[i], i + 1, this.issues))
      const rels = new Rels(undefined)
      rels.add(REL('externalLinkPath'), l1[i].target, true)
      this.parts.set(relsPath(part), rels.xml())
      this.types.override(part, `${TYPE}.externalLink+xml`)
      added.push(`<externalReference r:id="${this.wbRels.add(REL('externalLink'), relative(this.src.wbPath, part))}"/>`)
    }
    if (!added.length) return
    this.changed = true
    if (list) {
      const close = wb.lastIndexOf('</', list.end)
      edits.push({ part: this.src.wbPath, start: close, end: close, text: added.join('') })
    } else {
      const at = insertAt(['sheets', 'functionGroups'])
      edits.push({ part: this.src.wbPath, start: at, end: at, text: `<externalReferences>${added.join('')}</externalReferences>` })
    }
  }

  /** The defined names: the book's, each as the file wrote it where it did not change, and the file's the model does not take, by their sheets' places now. */
  private names(wb: string, kids: ReturnType<typeof childrenOf>, place: Map<string, number>, placeOf: Map<number, number>, edits: Edit[], insertAt: (before: string[]) => number): void {
    const { src, book0, book } = this
    const { decls } = src
    const scopeOf = (d: number | undefined) => (d === undefined ? undefined : decls[d]?.sheet !== undefined ? book0.sheets[decls[d].sheet!].name : decls[d]?.name)
    const el = kids.find((c) => c.local === 'definedNames')
    const filePlace = new Map(decls.flatMap((d, i) => (d.sheet !== undefined ? [[sheetKey(book0.sheets[d.sheet].name), i] as const] : [])))
    const x0 = nameElements(book0, new Issues(), this.w0.names, filePlace)
    const issues = new KeptIssues()
    const x1 = nameElements(book, issues, this.w1.names, place)
    const key = (name: string, scope: string | undefined) => `${scope === undefined ? '' : sheetKey(scope)}\n${caseKey(name)}`
    const was = new Map((book0.names ?? []).map((n, i) => [key(n.name, n.scope), x0[i]]))
    const held = new Set(was.keys())
    const fileText = new Map<string, string>()
    const foreign: string[] = []
    let foreignChanged = false
    // the worksheets of the file with a filter database, by `<sheet>`
    const filtered = new Set<number>()
    for (const e of el ? childrenOf(wb, el).filter((c) => c.local === 'definedName') : []) {
      const a = attrsOf(wb, e)
      const local = a.localSheetId === undefined ? undefined : Number(a.localSheetId)
      const text = wb.slice(e.start, e.end)
      const k = key(readText(a.name ?? ''), scopeOf(local))
      if (held.has(k)) {
        fileText.set(k, text)
        continue
      }
      if (local === undefined) foreign.push(text)
      else if (/^_xlnm\._FilterDatabase$/i.test(a.name ?? '') && this.filterHeld(local)) {
        // an AutoFilter's range, as the workbook has the AutoFilter: gone with it, written again where a host moved it
        filtered.add(local)
        const at = placeOf.get(local)
        const next = at === undefined ? undefined : this.filterName(text, local, at)
        if (next !== text) foreignChanged = true
        if (next !== undefined) foreign.push(next)
      } else {
        // a name the model does not take keeps its sheet by the sheet's place now; it goes with its sheet
        const at = placeOf.get(local)
        if (at !== local) foreignChanged = true
        if (at !== undefined) foreign.push(at === local ? text : setAttrs(text.slice(0, text.indexOf('>') + 1), { localSheetId: String(at) }) + text.slice(text.indexOf('>') + 1))
      }
    }
    // a sheet's AutoFilter where the file has none gets its filter database, as writeXlsx writes it
    for (const [j, s] of book.sheets.entries()) {
      const k = this.id.source[j]
      if (!s.autoFilter || (k !== undefined && (filtered.has(src.sheets[k].decl) || book0.sheets[k].autoFilter))) continue
      foreign.push(filterDatabase(s.name, s.autoFilter.range, place.get(sheetKey(s.name))!))
      foreignChanged = true
    }
    const ours = (book.names ?? []).flatMap((n, i) => {
      const x = x1[i]
      if (x === undefined) return []
      const k = key(n.name, n.scope)
      return [was.get(k) === x && fileText.has(k) ? fileText.get(k)! : x]
    })
    // the names the model does not take with their references written back count as changed
    if (x0.filter(Boolean).join('') === x1.filter(Boolean).join('') && !foreignChanged && wb === src.wb) return
    this.changed = true
    adopt(issues, this.issues)
    const all = [...ours, ...foreign]
    const text = all.length ? `<definedNames>${all.join('')}</definedNames>` : ''
    if (el) edits.push({ part: src.wbPath, start: el.start, end: el.end, text })
    else if (text) {
      const at = insertAt(BEFORE_NAMES)
      edits.push({ part: src.wbPath, start: at, end: at, text })
    }
  }

  /** Whether the file's `<sheet>` `d` is a worksheet whose AutoFilter the workbook holds, before or after the change. */
  private filterHeld(d: number): boolean {
    const k = this.src.decls[d]?.sheet
    const j = k === undefined ? undefined : this.id.target.get(k)
    return k !== undefined && (this.book0.sheets[k].autoFilter !== undefined || (j !== undefined && this.book.sheets[j].autoFilter !== undefined))
  }

  /**
   * A `_xlnm._FilterDatabase` of the file's `<sheet>` `d` (`text`, its reference moved) at place `at`: none without the
   * workbook's AutoFilter, the file's over its range or where the AutoFilter did not change (Excel leaves it `#REF!`
   * under cells moved over the AutoFilter), else as writeXlsx writes it.
   */
  private filterName(text: string, d: number, at: number): string | undefined {
    const k = this.src.decls[d].sheet!
    const s = this.book.sheets[this.id.target.get(k)!]
    if (!s.autoFilter) return undefined
    const head = text.slice(0, text.indexOf('>') + 1)
    const area = areasOf(readText(text.slice(head.length, text.lastIndexOf('<'))))?.[0]
    const x = parseRange(s.autoFilter.range)!
    const same = area && 'rect' in area && sheetKey(area.sheet) === sheetKey(s.name) && area.rect.r1 === x.r1 && area.rect.r2 === x.r2 && area.rect.c1 === x.c1 && area.rect.c2 === x.c2
    return (same || this.book0.sheets[k].autoFilter?.range === s.autoFilter.range) && at === d ? text : filterDatabase(s.name, s.autoFilter.range, at)
  }

  /** The parts of the book as a whole: styles, metadata, the calculation chain, relationships and content types. */
  private finish(): void {
    const { src } = this
    const styles = this.styles.text()
    if (styles !== undefined) {
      const r = src.wbRels.find((x) => x.kind === 'styles' && !x.external)
      const part = r?.target ?? 'xl/styles.xml'
      this.parts.set(part, styles)
      if (!r) {
        this.wbRels.add(REL('styles'), relative(src.wbPath, part))
        this.types.override(part, `${TYPE}.styles+xml`)
      }
    }
    const chain = src.wbRels.find((x) => x.kind === 'calcChain' && !x.external)
    if (chain) {
      this.parts.remove(chain.target)
      this.types.remove(chain.target)
      this.wbRels.remove(chain.id)
    }
    this.metadata()
    if (this.wbRels.changed) this.parts.set(relsPath(src.wbPath), this.wbRels.xml())
    if (this.types.changed) this.parts.set('[Content_Types].xml', this.types.xml())
  }

  /** The marks of dynamic arrays and of errors that written cells take and the file lacks, added to its metadata or as parts of their own. */
  private metadata(): void {
    const { meta, metaParts } = this.setup
    const written = [...this.parts.text.values()]
    const used = (mark: string) => written.some((t) => t.includes(mark))
    const need = {
      dynamic: !meta.dynamic.has && used(` cm="${meta.dynamic.mark}"`),
      rich: Object.values(meta.rich).some((r) => !r.has && used(` vm="${r.vm}"`)),
    }
    if (!need.dynamic && !need.rich) return
    const added = addMeta(meta, need, metaParts.sheetMetadata, metaParts.rdRichValue, metaParts.rdRichValueStructure)
    const put = (kind: (typeof META)[number], text: string | undefined, name: string, type: string, rel: string) => {
      if (text === undefined) return
      const r = this.src.wbRels.find((x) => x.kind === kind && !x.external)
      const part = r?.target ?? this.parts.fresh(name, '.xml')
      this.parts.set(part, text)
      if (r) return
      this.wbRels.add(rel, relative(this.src.wbPath, part))
      this.types.override(part, type)
    }
    put('sheetMetadata', added.metadata, 'xl/metadata', `${TYPE}.sheetMetadata+xml`, REL('sheetMetadata'))
    put('rdRichValue', added.values, 'xl/richData/rdrichvalue', 'application/vnd.ms-excel.rdrichvalue+xml', `${RICH_REL}/rdRichValue`)
    put('rdRichValueStructure', added.structures, 'xl/richData/rdrichvaluestructure', 'application/vnd.ms-excel.rdrichvaluestructure+xml', `${RICH_REL}/rdRichValueStructure`)
  }

  async zip(): Promise<Uint8Array> {
    const { archive } = this.src
    const kept = records(this.src.bytes)
    const files: Parameters<typeof zip>[0] = []
    for (const name of archive.names) {
      const k = name.toLowerCase()
      if (this.parts.removed.has(k)) continue
      const text = this.parts.text.get(k)
      files.push(text !== undefined ? { name, data: utf8(text) } : kept.get(k)!)
    }
    for (const name of this.parts.added) files.push({ name, data: utf8(this.parts.text.get(name.toLowerCase())!) })
    return zip(files)
  }
}

/** The issues of writing a sheet that concern what was written from the model: the cells written, and the elements written. */
function adopt(from: KeptIssues, to: Issues, cells?: Set<string>, written?: Set<string>): void {
  from.copyTo(to, (_, message, where) => {
    if (!cells || !written) return true
    switch (message.key) {
      case 'condFormulaOut':
        return written.has('conditionalFormatting')
      case 'ruleFormulaOut':
      case 'ruleTextCut':
        return written.has('dataValidations')
      case 'tooltipCut':
        return written.has('hyperlinks')
      case 'filterModeOut':
        return written.has('autoFilter')
      case 'noteCut':
        return true
      default:
        return where.cell === undefined || cells.has(where.cell)
    }
  })
}
