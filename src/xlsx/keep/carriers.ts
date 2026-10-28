/**
 * Carriers: names of the book that hold the references of the file's parts the model does not have (`openKept` adds
 * them, `writeKept` reads them back). Each sheet has one on its last cell, which a row or column inserted or deleted
 * anywhere on it moves and a rename or a deletion of the sheet changes, and which tells which sheet of the book the
 * file's sheet became. Each part with references the model does not have is an item with carriers of its own, which an
 * operation of the structure moves as Excel moves a name's reference, and the edits that write them back (`fix`): the
 * series and labels of charts, the anchors of drawings by their `editAs`, the place and the source of pivot tables,
 * sparklines, a sheet's AutoFilter and sort, the names the model does not take (the print area and titles), page
 * breaks, protected ranges and the conditional formats the model does not read (`rules.ts`). A part an item cannot write back as Excel would refuses (comment threads, a pivot
 * table cut). Parts whose references are not read (`unparsed`) are known only by their sheet's carrier: rows and
 * columns inserted or deleted there are seen, a range moved over them is not. A cell with what only the file holds
 * (formatted runs of text, a rich value the writer does not write, phonetic text) has a carrier of its own, which
 * tells where the cell went; past `KEPT_CELLS` of them on a sheet they have none. A cell format that looks as an earlier
 * one has carriers of its rectangles on each sheet (`formats.ts`); past `KEPT_FORMATS` rectangles on a sheet, none.
 */
import type { BookInput, NameInput } from '../../contract.js'
import { caseKey } from '../../collate.js'
import { ADDR, MAX_COLS, MAX_ROWS, parseAddress } from '../../model.js'
import type { KeptPart } from '../../locale/messages.js'
import type { KeptMessage } from './messages.js'
import type { KeptRefusalCode } from './refusals.js'
import type { Rect } from '../../structure/area.js'
import type { Ast } from '../../syntax/ast.js'
import { FormulaError } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'
import { canonical, mapAst, quoteSheet, sheetKey, walk } from '../../syntax/print.js'
import { attr, attributes, elements, placed } from '../xml/scan.js'
import { escapeAttr, escapeText, readText } from '../xml/text.js'
import { parseFormula } from '../read/formulas.js'
import { CARRIER_PREFIX } from '../write/book.js'
import { cellMetadata } from '../read/metadata.js'
import { partText } from '../read/package.js'
import { ABSOLUTE, absText, areasOf, areaText, dollarsOf, lastCell, plainArea, rectsNow, sheetPrefix, sqrefCarrier, sqrefRects, type Area, type Dollars } from './refs.js'
import type { Source } from './source.js'
import type { RuleSource } from '../read/conditional.js'
import { palette } from '../styles/color.js'
import { readDxfs } from '../styles/dxf.js'
import { formatAreas, type FormatKind } from './formats.js'
import { fileRules, ruleItems } from './rules.js'
import type { KeptStyles } from './styles.js'
import { attrsOf, childrenOf, ContentTypes, rootOf, valueAt, type Placed } from './xml.js'

/** The prefix of the carriers' names, which `writeXlsx` leaves out. */
export const PREFIX = CARRIER_PREFIX

/** A reference of a part the model does not have, as a name of the book. */
export interface Carrier {
  name: string
  ref: string
  /** `ref`: the reference counts as written, its sheet too; `place`: the part lies on the sheet, only its cells count. */
  kind: 'ref' | 'place'
}

/** A change of a part's text: from `start` to `end` of the part as the file has it, `text` in their place. */
export interface Edit {
  part: string
  start: number
  end: number
  text: string
}

/** What writing an item's carriers back gives: the edits, or why the file cannot take the change. */
export type Fix = { edits: Edit[] } | { refuse: KeptMessage; code: KeptRefusalCode }

/** The size of a sheet's rows and columns as the book written has them, in EMUs (a drawing's offsets count in them). */
export interface Geometry {
  row(r: number): number
  col(c: number): number
}

/** A part's references: its carriers, and the edits that write what they hold now (by carrier) back to the file. */
export interface Item {
  what: KeptPart
  /** The sheet of the source the part lies on: deleted with it, its references go too. */
  on?: number
  carriers: Carrier[]
  /** The edits; `at` gives the lines' sizes of the sheet the part lies on. */
  fix(now: string[], at: Geometry): Fix
}

/** A cell of the file with what only the file holds, and its carrier. */
export interface KeptCell {
  addr: string
  carrier: Carrier
}

/** The cells of a sheet with what only the file holds past which they have no carriers: `writeKept` reports them where rows or columns moved. */
export const KEPT_CELLS = 5_000

/** A cell format's rectangle of cells, rows or columns, as a carrier (of several areas). */
export interface KeptFormat {
  kind: FormatKind
  s: number
  carrier: Carrier
}

/** The rectangles of cell formats on a sheet past which they have no carriers: `writeKept` reports their cells where rows or columns moved. */
export const KEPT_FORMATS = 5_000

/** The length of a carrier of a format's rectangles, past which they go to another. */
const FORMAT_TEXT = 2_000

export interface Carriers {
  /** Each of the source's sheets' carrier, in the book's order of them. */
  sheets: Carrier[]
  /** By sheet of the source: its cells with what only the file holds; none past `KEPT_CELLS`, their count in `uncarried`. */
  cells: KeptCell[][]
  uncarried: number[]
  /** By sheet of the source: the carriers of its cell formats that look as earlier ones (`formats.ts`); none past `KEPT_FORMATS`, the count of their cells in `unformatted`. */
  formats: (KeptFormat[] | undefined)[]
  unformatted: number[]
  items: Item[]
  /** By sheet of the source: its parts whose references no carrier holds. */
  unparsed: KeptPart[][]
  /** Parts whose references no carrier holds and that may refer to any sheet. */
  anywhere: KeptPart[]
  /** The tables and names parts take their data from by name (a pivot cache's source): where the name lies in the part. */
  byName: { name: string; what: KeptPart; part: string; at: [number, number] }[]
  /** By sheet of the source: parts that name the sheet (a pivot chart's pivot table), where the name lies. */
  named: { what: KeptPart; part: string; at: [number, number]; prefix: string; rest: string }[][]
  /** By sheet of the source: its conditional formats as the reader met them (`rules.ts`). */
  rules: RuleSource[][]
}

/** Elements of a worksheet whose references are not read. */
const UNPARSED = new Set(['oleObjects', 'controls', 'dataConsolidate', 'scenarios', 'customSheetViews', 'cellWatches', 'webPublishItems', 'smartTags', 'AlternateContent'])

const SPARKLINES = '{05C60535-1F16-4FD2-B633-F4F36F0B64E0}'
const PROTECTED = '{FC87AEE6-9EDD-4A0A-B7FB-166176984837}'

/** Extensions of a worksheet without references of cells, or with those the model or a carrier holds. */
const EXT_KNOWN = new Set(
  [
    '{78C0D931-6437-407d-A8EE-F0AAD7539E65}', // conditional formats
    '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}', // data validation
    SPARKLINES,
    PROTECTED,
    '{A8765BA9-456A-4dab-B4F3-ACF838C121DE}', // slicers of pivot tables
    '{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}', // slicers of tables
    '{7E03D99C-DC04-49d9-9315-930204A7B6E9}', // timelines
    '{01252117-D84E-4E92-8308-4BE1C098FCBB}', // ignored errors
  ].map((u) => u.toUpperCase()),
)

/** Formatted runs of text, or phonetic ones, in the XML of a string. */
const RUNS = /<(?:[\w.-]+:)?(?:r|rPh)[\s>]/
/** A cell of a sheet's XML: its attributes and content. */
const CELL_XML = /<(?:[\w.-]+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?c>)/g
/** The rich values of errors the writer writes itself (`write/book.ts`). */
const OWN_ERRORS = new Set(['#SPILL!', '#CALC!'])

/** Rows taken below an anchor to tell where it lands when its rows are deleted, and columns beside it. */
const BELOW = 100_000
const BESIDE = 1_000

/** Where an element lies in a part: from its opening tag to its end. */
interface Span {
  start: number
  end: number
}

/** The elements of a local name in a part from `from` to `to`: each where it lies, where its content lies, and its attributes and content. */
function within(xml: string, from: number, to: number, local: string, first = false): (Span & { open: number; close: number; attrs: string; body: string | null })[] {
  const sub = xml.slice(from, to)
  return placed(sub, local, first).map((p) => {
    const open = from + sub.indexOf('>', p.start) + 1
    return { start: from + p.start, end: from + p.end, open, close: p.body === null ? open : open + p.body.length, attrs: p.attrs, body: p.body }
  })
}

/** A reference written back as a file's formula: each area with its sheet and the `$`s it had (`dollars`), `#REF!` where lost; the whole `#REF!` for a sheet deleted. */
function formulaText(areas: Area[], dollars: Dollars[], parens: boolean): string {
  if (areas.every((a) => 'lost' in a && a.sheet === undefined)) return '#REF!'
  const text = areas.map((a, i) => ('rect' in a ? `${sheetPrefix(a.sheet)}!${areaText(a.rect, dollars[i] ?? ABSOLUTE)}` : a.sheet === undefined ? '#REF!' : `${sheetPrefix(a.sheet)}!#REF!`)).join(',')
  return parens ? `(${text})` : text
}

export async function carriersOf(src: Source, book: BookInput, styles: KeptStyles): Promise<Carriers> {
  const names = book.sheets.map((s) => s.name)
  const known = new Set(names.map(sheetKey))
  const taken = new Set((book.names ?? []).map((n) => caseKey(n.name)))
  /** A carrier's name: the stem and the first number from `at.n` the book does not take. */
  const fresh = (stem: string, at: { n: number }) => {
    let name = `${PREFIX}${stem}${at.n++}`
    while (taken.has(caseKey(name))) name = `${PREFIX}${stem}${at.n++}`
    return name
  }
  const parts = { n: 0 }
  const carrier = (ref: string, kind: Carrier['kind']): Carrier => ({ name: fresh('', parts), ref, kind })
  const out: Carriers = { sheets: [], cells: names.map(() => []), uncarried: names.map(() => 0), formats: names.map(() => []), unformatted: names.map(() => 0), items: [], unparsed: names.map(() => []), anywhere: [], byName: [], named: names.map(() => []), rules: [] }
  for (const [i, sheet] of names.entries()) {
    let name = `${PREFIX}s${i}`
    for (let k = 2; taken.has(caseKey(name)); k++) name = `${PREFIX}s${i}.${k}`
    out.sheets.push({ name, ref: lastCell(sheet), kind: 'place' })
  }
  const place = (sheet: number, x: Rect) => carrier(`${quoteSheet(names[sheet])}!${absText(x)}`, 'place')

  /**
   * A file's formula of references (a chart's series, a sparkline's data, a name's) at `span` of a part: an item with its
   * carrier, the references absolute (the engine moves no relative reference of a name), written back with the `$`s
   * the file gave them; none when it holds no reference of the book's sheets.
   */
  const formula = (part: string, span: [number, number], text: string, what: KeptPart, escape: (s: string) => string, on?: number) => {
    const f = parseFormula(text, { union: true })
    if (!f.ast) return void out.anywhere.push(what)
    if (!refersToBook(f.ast, known)) return
    const dollars: Dollars[] = []
    walk(f.ast, (x) => {
      if (x.k === 'ref') dollars.push(dollarsOf(x.ref))
    })
    const absolute = mapAst(f.ast, (x) => (x.k === 'ref' || x.k === 'ref3d' ? { ...x, ref: { ...x.ref, text: absText(x.ref) } } : undefined))
    const c = carrier(canonical(absolute), 'ref')
    const parens = text.trim().startsWith('(')
    out.items.push({
      what,
      ...(on === undefined ? {} : { on }),
      carriers: [c],
      fix([now]) {
        const areas = areasOf(now)
        const written = areas ? formulaText(areas, dollars, parens) : now
        return { edits: [{ part, start: span[0], end: span[1], text: escape(written) }] }
      },
    })
  }

  // the shared strings with formatted runs, and the rich values of the cells
  const part = (kind: string) => partText(src.archive, src.wbRels, kind)
  const [sst, metaXml, richValues, richStructures, stylesXml, themeXml] = await Promise.all(['sharedStrings', 'sheetMetadata', 'rdRichValue', 'rdRichValueStructure', 'styles', 'theme'].map(part))
  const colors = palette(themeXml, stylesXml)
  const dxfs = readDxfs(stylesXml, colors)
  const richStrings = new Set<number>()
  if (sst && RUNS.test(sst)) elements(sst, 'si').forEach((e, k) => RUNS.test(e.body ?? '') && richStrings.add(k))
  const meta = cellMetadata(metaXml, richValues, richStructures)

  const types = new ContentTypes(src.types)
  for (const [i, s] of src.sheets.entries()) {
    const sheet = names[i]
    const xml = s.xml
    const root = rootOf(xml)
    // the conditional formats the model does not have, with their carriers
    out.rules[i] = fileRules(xml, sheet, dxfs, colors)
    const rules = ruleItems(xml, out.rules[i], sheet, i, s.part, carrier)
    out.items.push(...rules.items)
    out.unparsed[i].push(...rules.unparsed)
    const kids = root ? childrenOf(xml, root) : []
    if (styles.twins) formats(xml, kids, i)
    for (const c of kids) {
      if (c.local === 'sheetData') keptCells(xml, c, i)
      else if (UNPARSED.has(c.local)) out.unparsed[i].push({ part: 'element', sheet, element: c.local })
      else if (c.local === 'autoFilter' || c.local === 'sortState') filter(xml, c, i, s.part)
      else if (c.local === 'rowBreaks' || c.local === 'colBreaks') breaks(xml, c, i, s.part)
      else if (c.local === 'protectedRanges') protectedRanges(xml, c, i, s.part)
      else if (c.local === 'extLst')
        for (const e of childrenOf(xml, c)) {
          const uri = (attrsOf(xml, e).uri ?? '').toUpperCase()
          if (!EXT_KNOWN.has(uri)) out.unparsed[i].push({ part: 'extension', sheet, uri })
          else if (uri === SPARKLINES) sparklines(xml, e, i, s.part)
          else if (uri === PROTECTED) protectedRanges(xml, e, i, s.part)
        }
    }
    for (const r of s.rels) {
      if (r.external) continue
      if (r.kind === 'drawing') anchors((await src.archive.text(r.target)) ?? '', i, r.target)
      else if (r.kind === 'pivotTable') pivotTable((await src.archive.text(r.target)) ?? '', i, r.target)
      else if (r.kind === 'threadedComment') threads((await src.archive.text(r.target)) ?? '', i, r.target)
      else if (r.kind === 'vmlDrawing') {
        // a legacy drawing holds the notes, which the model has, and form controls, which it does not
        const vml = (await src.archive.text(r.target)) ?? ''
        if (placed(vml, 'ClientData').some((d) => attributes(d.attrs).ObjectType !== 'Note')) out.unparsed[i].push({ part: 'formControls', sheet })
      }
    }
  }
  // the names the model does not take (built-in ones: the print area and titles, a sheet's filter database)
  const held = new Set((book.names ?? []).map((x) => `${x.scope === undefined ? '' : sheetKey(x.scope)}\n${caseKey(x.name)}`))
  const wbRoot = rootOf(src.wb)
  const defined = wbRoot && childrenOf(src.wb, wbRoot).find((c) => c.local === 'definedNames')
  for (const d of defined ? childrenOf(src.wb, defined) : []) {
    const a = attrsOf(src.wb, d)
    const name = readText(a.name ?? '')
    if (/^_xlfn\./i.test(name) || d.end === d.open) continue
    const decl = a.localSheetId === undefined ? undefined : src.decls[Number(a.localSheetId)]
    // a worksheet's name as the book has it (the reader repairs names Excel would not take)
    const scope = a.localSheetId === undefined ? '' : decl?.sheet !== undefined ? names[decl.sheet] : decl?.name
    if (held.has(`${scope === undefined ? '' : sheetKey(scope)}\n${caseKey(name)}`)) continue
    const span: [number, number] = [d.open, d.end - d.qname.length - 3]
    formula(src.wbPath, span, readText(src.wb.slice(...span)), { part: 'name', name }, escapeText)
  }
  for (const part of types.parts((t) => /drawingml\.chart\+xml$|chartex\+xml$/i.test(t))) {
    const xml = (await src.archive.text(part)) ?? ''
    for (const f of within(xml, 0, xml.length, 'f')) formula(part, [f.open, f.close], readText(f.body ?? ''), { part: 'chart', path: part }, escapeText)
    // a pivot chart's table, by its sheet's name (`[Book1.xlsx]Pivot!PivotTable1`)
    const source = within(xml, 0, xml.length, 'pivotSource', true)[0]
    const name = source && within(xml, source.start, source.end, 'name', true)[0]
    const m = name && /^('?)(\[[^\]]*\])((?:[^'!]|'')+)\1!(.*)$/.exec(readText(name.body ?? ''))
    const k = m ? names.findIndex((x) => sheetKey(x) === sheetKey(m[3].replace(/''/g, "'"))) : -1
    if (m && k >= 0) out.named[k].push({ what: { part: 'pivotChart', path: part }, part, at: [name!.open, name!.close], prefix: m[2], rest: m[4] })
  }
  for (const part of types.parts((t) => /pivotCacheDefinition\+xml$/i.test(t))) pivotCache((await src.archive.text(part)) ?? '', part)
  // macro and dialog sheets hold formulas over the worksheets
  for (const d of src.decls) if (d.kind !== undefined && d.kind !== 'worksheet' && d.kind !== 'chartsheet') out.anywhere.push({ part: 'sheetKind', kind: d.kind, name: d.name })
  return out

  /**
   * A sheet's cells with what only the file holds: a shared string with formatted runs or phonetic text, an inline one
   * with them, a rich value the writer does not write (a picture in the cell, a data type), phonetic text shown; each
   * with a carrier of its own, of a name apart from the parts' (`c<sheet>.<n>`). Past `KEPT_CELLS` of them, none.
   */
  function keptCells(xml: string, data: Placed, i: number) {
    if (!richStrings.size && !/ (?:vm|ph)="|<(?:[\w.-]+:)?is>/.test(xml)) return
    const found: string[] = []
    for (const m of xml.slice(data.open, data.end).matchAll(CELL_XML)) {
      const a = attributes(m[1])
      const body = m[2] ?? ''
      const kept =
        (a.t === 's' && richStrings.has(Number(/<(?:[\w.-]+:)?v>(\d+)</.exec(body)?.[1]))) ||
        (a.t === 'inlineStr' && RUNS.test(body)) ||
        (a.vm !== undefined && !OWN_ERRORS.has(meta.error(Number(a.vm)) ?? '')) ||
        (a.ph !== undefined && a.ph !== '0' && a.ph !== 'false')
      // a cell without its address keeps what it holds only where it stays
      if (kept && a.r !== undefined && ADDR.test(a.r)) found.push(a.r)
    }
    if (found.length > KEPT_CELLS) return void (out.uncarried[i] = found.length)
    const at = { n: 0 }
    out.cells[i] = found.map((addr) => {
      const name = fresh(`c${i}.`, at)
      const { row, col } = parseAddress(addr)
      return { addr, carrier: { name, ref: `${quoteSheet(names[i])}!${absText({ r1: row, r2: row, c1: col, c2: col })}`, kind: 'place' } }
    })
  }

  /** A sheet's cell formats that look as earlier ones: carriers of their rectangles, several to a carrier (`x<sheet>.<n>`); past `KEPT_FORMATS` rectangles, none. */
  function formats(xml: string, kids: Placed[], i: number) {
    const f = formatAreas(xml, kids.find((c) => c.local === 'sheetData'), kids.find((c) => c.local === 'cols'), (s) => styles.twin(s))
    if (f.areas.length > KEPT_FORMATS) {
      out.formats[i] = undefined
      out.unformatted[i] = f.cells
      return
    }
    // the rectangles of one format and kind in a carrier, as long as it stays short
    const by = new Map<string, KeptFormat[]>()
    const at = { n: 0 }
    for (const a of f.areas) {
      const text = `${quoteSheet(names[i])}!${absText(a.rect)}`
      const key = `${a.kind}:${a.s}`
      const last = by.get(key)?.at(-1)
      if (last && last.carrier.ref.length + text.length < FORMAT_TEXT) last.carrier.ref += `,${text}`
      else (by.get(key) ?? by.set(key, []).get(key)!).push({ kind: a.kind, s: a.s, carrier: { name: fresh(`x${i}.`, at), ref: text, kind: 'place' } })
    }
    out.formats[i] = [...by.values()].flat()
  }

  /**
   * A drawing's anchors (`twoCellAnchor`, `oneCellAnchor`; `absoluteAnchor` stays where it is) and how each follows
   * the cells: one that moves and sizes with them (`editAs="twoCell"`, the default) by both its corners, one that only
   * moves with them (`oneCell`, and a one-cell anchor) by its top left corner, its size kept, one fixed (`absolute`)
   * not at all. A corner follows its row and its column (carriers of the whole lines); where its row or column is
   * deleted it lands where the deleted lines were, its offset 0: at the rectangle's new edge, or, with all of it
   * deleted, at the probe's, the lines below and beside it. The rectangle tells a range moved over the drawing, refused.
   */
  function anchors(xml: string, i: number, part: string) {
    const root = rootOf(xml)
    /** The anchors of the drawing itself, not in alternate content (`mc:AlternateContent`): those it can lose. */
    const children = new Set(root ? childrenOf(xml, root).map((c) => c.start) : [])
    for (const local of ['twoCellAnchor', 'oneCellAnchor'] as const)
      for (const e of within(xml, 0, xml.length, local)) {
        const editAs = local === 'oneCellAnchor' ? 'oneCell' : (attributes(e.attrs).editAs ?? 'twoCell')
        if (editAs === 'absolute') continue
        const corner = (tag: string) => {
          const p = within(xml, e.open, e.close, tag, true)[0]
          if (!p) return undefined
          const at = (k: string) => {
            const v = within(xml, p.open, p.close, k, true)[0]
            return v && v.body !== null ? { value: Number(v.body), start: v.open, end: v.close } : undefined
          }
          const [col, colOff, row, rowOff] = ['col', 'colOff', 'row', 'rowOff'].map(at)
          const ok = col && row && colOff && rowOff && Number.isInteger(col.value) && Number.isInteger(row.value) && col.value >= 0 && row.value >= 0 && col.value < MAX_COLS && row.value < MAX_ROWS
          return ok ? { col: col!, colOff: colOff!, row: row!, rowOff: rowOff! } : undefined
        }
        const from = corner('from')
        const to = local === 'twoCellAnchor' ? corner('to') : undefined
        const what: KeptPart = { part: 'drawing', sheet: names[i], path: part }
        if (!from || (local === 'twoCellAnchor' && (!to || to.row.value < from.row.value || to.col.value < from.col.value))) {
          out.unparsed[i].push(what)
          continue
        }
        const sized = editAs === 'twoCell' && to !== undefined
        const end = to ?? from
        const rect = { r1: from.row.value, r2: end.row.value, c1: from.col.value, c2: end.col.value }
        const probe = { r1: rect.r1, r2: Math.min(rect.r2 + BELOW, MAX_ROWS - 1), c1: rect.c1, c2: Math.min(rect.c2 + BESIDE, MAX_COLS - 1) }
        const row = (r: number) => place(i, { r1: r, r2: r, c1: 0, c2: MAX_COLS - 1 })
        const col = (c: number) => place(i, { r1: 0, r2: MAX_ROWS - 1, c1: c, c2: c })
        out.items.push({
          what,
          on: i,
          carriers: [place(i, rect), place(i, probe), row(rect.r1), col(rect.c1), ...(sized ? [row(rect.r2), col(rect.c2)] : [])],
          fix([rectNow, probeNow, ...linesNow], at) {
            const [r, p] = [rectsNow(rectNow)?.[0], rectsNow(probeNow)?.[0]]
            const line = (k: number, axis: 'r' | 'c') => {
              const x = linesNow[k] === undefined ? undefined : rectsNow(linesNow[k])?.[0]
              return x ? (axis === 'r' ? x.r1 : x.c1) : undefined
            }
            const [r1, c1, r2, c2] = [line(0, 'r'), line(1, 'c'), line(2, 'r'), line(3, 'c')]
            // a drawing that sizes with its cells goes with all its rows or columns deleted, as Excel deletes it (the set
            // keep/struct); the parts only it points at go with it (`writeKept`). Its cells lost under a range moved over
            // them, its lines stay: so does it.
            if (sized && !r && ((r1 === undefined && r2 === undefined) || (c1 === undefined && c2 === undefined)))
              return children.has(e.start) ? { edits: [{ part, start: e.start, end: e.end, text: '' }] } : { refuse: { key: 'keptAlternate', what }, code: 'deleted' }
            // a range moved over the drawing moves its rectangle and not its lines: the drawing would stay behind
            if (r && r1 !== undefined && c1 !== undefined && (r.r1 !== r1 || r.c1 !== c1)) return { refuse: { key: 'keptDrawingMoved', what }, code: 'moved' }
            if (!p) return { refuse: { key: 'keptDrawingDeleted', what }, code: 'deleted' }
            // where a corner of a deleted line lands: the first line after the deleted ones
            const top = r1 ?? (r ? r.r1 : p.r1)
            const left = c1 ?? (r ? r.c1 : p.c1)
            const edits: Edit[] = []
            const set = (v: { start: number; end: number }, n: number) => edits.push({ part, start: v.start, end: v.end, text: String(n) })
            set(from.row, top)
            set(from.col, left)
            if (r1 === undefined) set(from.rowOff, 0)
            if (c1 === undefined) set(from.colOff, 0)
            if (to && sized) {
              set(to.row, r2 ?? (r ? r.r2 + 1 : p.r1))
              set(to.col, c2 ?? (r ? r.c2 + 1 : p.c1))
              if (r2 === undefined) set(to.rowOff, 0)
              if (c2 === undefined) set(to.colOff, 0)
            } else if (to) {
              // its size kept: a corner that lost its offset takes it off the other corner too
              const keep = (lo: number, line: number, off: number, was: number, size: (i: number) => number): [number, number] => {
                let [n, o] = [lo + line, off - was]
                while (o < 0 && n > lo) o += size(--n)
                return [n, Math.max(0, o)]
              }
              const [row, rowOff] = r1 === undefined ? keep(top, to.row.value - from.row.value, to.rowOff.value, from.rowOff.value, at.row) : [top + (to.row.value - from.row.value), to.rowOff.value]
              const [col, colOff] = c1 === undefined ? keep(left, to.col.value - from.col.value, to.colOff.value, from.colOff.value, at.col) : [left + (to.col.value - from.col.value), to.colOff.value]
              set(to.row, row)
              set(to.col, col)
              if (rowOff !== to.rowOff.value) set(to.rowOff, rowOff)
              if (colOff !== to.colOff.value) set(to.colOff, colOff)
            }
            return { edits }
          },
        })
      }
  }

  /** A pivot table's place on its sheet: it moves with its cells; a change that would cut it, which Excel refuses, is refused. */
  function pivotTable(xml: string, i: number, part: string) {
    const loc = within(xml, 0, xml.length, 'location', true)[0]
    const span = loc && valueAt(xml, loc, 'ref')
    const rects = span ? sqrefRects(xml.slice(...span)) : null
    const what: KeptPart = { part: 'pivotTable', sheet: names[i], path: part }
    if (!span || rects?.length !== 1) return void out.unparsed[i].push(what)
    const was = rects[0]
    out.items.push({
      what,
      on: i,
      carriers: [place(i, was)],
      fix([now]) {
        const x = rectsNow(now)?.[0]
        if (!x || x.r2 - x.r1 !== was.r2 - was.r1 || x.c2 - x.c1 !== was.c2 - was.c1) return { refuse: { key: 'keptCut', what }, code: 'moved' }
        return { edits: [{ part, start: span[0], end: span[1], text: plainArea(x) }] }
      },
    })
  }

  /** A pivot cache's source: cells of a sheet (`ref` and `sheet`, a consolidation's range sets too), or a table or a name by its name. */
  function pivotCache(xml: string, part: string) {
    const what: KeptPart = { part: 'pivotSource', path: part }
    for (const local of ['worksheetSource', 'rangeSet'])
      for (const e of within(xml, 0, xml.length, local)) {
        const a = attributes(e.attrs)
        if (attr(a, 'id') !== undefined) continue
        if (local === 'worksheetSource' && a.name !== undefined) {
          out.byName.push({ name: readText(a.name), what, part, at: valueAt(xml, e, 'name')! })
          continue
        }
        const [ref, sheet] = [valueAt(xml, e, 'ref'), valueAt(xml, e, 'sheet')]
        const rects = ref ? sqrefRects(xml.slice(...ref)) : null
        if (!ref || !sheet || rects?.length !== 1 || !known.has(sheetKey(readText(xml.slice(...sheet))))) continue
        out.items.push({
          what,
          carriers: [carrier(`${quoteSheet(readText(xml.slice(...sheet)))}!${absText(rects[0])}`, 'ref')],
          fix([now]) {
            const area = areasOf(now)?.[0]
            if (!area) return { refuse: { key: 'keptNoRef', what }, code: 'deleted' }
            // its sheet deleted, Excel makes the source a link to the file by its path (`r:id` of an external link path,
            // the set keep/struct), which writeKept does not know, and does not open the source left as it was (keep
            // struct-open); its cells deleted is not measured
            if ('lost' in area) return { refuse: { key: area.sheet === undefined ? 'keptCacheSheet' : 'keptCellsDeleted', what }, code: 'deleted' }
            return { edits: [{ part, start: ref[0], end: ref[1], text: plainArea(area.rect) }, { part, start: sheet[0], end: sheet[1], text: escapeAttr(area.sheet) }] }
          },
        })
      }
  }

  /**
   * A sheet's AutoFilter or sort (`e`): its range, the column of each criterion (`colId` counts from the range's first
   * column; a criterion whose column was deleted goes) and the range of each sort condition (one deleted goes); its
   * range deleted, it goes, as the workbook's AutoFilter goes (the keep sets `filter-*`); the sheet's writing compares
   * the AutoFilter with the workbook's (`spliceSheet`).
   */
  function filter(xml: string, e: Span & { local: string }, i: number, part: string) {
    const what: KeptPart = { part: e.local === 'autoFilter' ? 'autoFilter' : 'sort', sheet: names[i] }
    const span = valueAt(xml, e, 'ref')
    const was = span && sqrefRects(xml.slice(...span))
    if (!span || was?.length !== 1) return void out.unparsed[i].push(what)
    const range = was[0]
    const columns = within(xml, e.start, e.end, 'filterColumn').map((c) => ({ c, colId: Number(attributes(c.attrs).colId), at: valueAt(xml, c, 'colId')! }))
    const sorts = within(xml, e.start, e.end, 'sortCondition').map((c) => ({ c, at: valueAt(xml, c, 'ref'), rects: sqrefRects(attributes(c.attrs).ref ?? '') }))
    // an AutoFilter's own sort
    const sortState = e.local === 'autoFilter' ? within(xml, e.start, e.end, 'sortState', true)[0] : undefined
    const sortAt = sortState && valueAt(xml, sortState, 'ref')
    const sortWas = sortAt && sqrefRects(xml.slice(...sortAt))
    if (columns.some((c) => !(c.colId >= 0) || !c.at) || sorts.some((s) => !s.at || s.rects?.length !== 1) || (sortState && sortWas?.length !== 1)) return void out.unparsed[i].push(what)
    out.items.push({
      what,
      on: i,
      carriers: [place(i, range), ...columns.map(({ colId }) => place(i, { ...range, c1: range.c1 + colId, c2: range.c1 + colId })), ...sorts.map((s) => place(i, s.rects![0])), ...(sortWas ? [place(i, sortWas[0])] : [])],
      fix([rangeNow, ...rest]) {
        const now = rectsNow(rangeNow)?.[0]
        if (!now) return { edits: [{ part, start: e.start, end: e.end, text: '' }] }
        const edits: Edit[] = [{ part, start: span[0], end: span[1], text: plainArea(now) }]
        columns.forEach(({ c, at }, k) => {
          const x = rectsNow(rest[k])?.[0]
          edits.push(x ? { part, start: at[0], end: at[1], text: String(x.c1 - now.c1) } : { part, start: c.start, end: c.end, text: '' })
        })
        sorts.forEach((s, k) => {
          const x = rectsNow(rest[columns.length + k])?.[0]
          edits.push(x ? { part, start: s.at![0], end: s.at![1], text: plainArea(x) } : { part, start: s.c.start, end: s.c.end, text: '' })
        })
        if (sortState && sortAt) {
          const x = rectsNow(rest[columns.length + sorts.length])?.[0]
          edits.push(x ? { part, start: sortAt[0], end: sortAt[1], text: plainArea(x) } : { part, start: sortState.start, end: sortState.end, text: '' })
        }
        return { edits: nested(edits) }
      },
    })
  }

  /** A sheet's page breaks, each by the line it comes before (`id`, from 0): it moves with the line, and goes with it. */
  function breaks(xml: string, e: Span & { local: string; qname: string }, i: number, part: string) {
    const rows = e.local === 'rowBreaks'
    const what: KeptPart = { part: 'pageBreaks', sheet: names[i] }
    const list = within(xml, e.start, e.end, 'brk').map((b) => ({ b, id: Number(attributes(b.attrs).id), at: valueAt(xml, b, 'id') }))
    if (list.some((x) => !x.at || !(x.id >= 0 && x.id < (rows ? MAX_ROWS : MAX_COLS)))) return void out.unparsed[i].push(what)
    out.items.push({
      what,
      on: i,
      carriers: list.map(({ id }) => place(i, rows ? { r1: id, r2: id, c1: 0, c2: MAX_COLS - 1 } : { r1: 0, r2: MAX_ROWS - 1, c1: id, c2: id })),
      fix(now) {
        const kept = list.flatMap(({ b, at }, k) => {
          const x = rectsNow(now[k])?.[0]
          return x ? [xml.slice(b.start, at![0]) + String(rows ? x.r1 : x.c1) + xml.slice(at![1], b.end)] : []
        })
        if (!kept.length) return { edits: [{ part, start: e.start, end: e.end, text: '' }] }
        const manual = kept.filter((t) => /\sman\s*=\s*["'](1|true)["']/.test(t)).length
        const tag = xml.slice(e.start, xml.indexOf('>', e.start) + 1).replace(/\scount\s*=\s*(["'])\d*\1/, ` count="${kept.length}"`).replace(/\smanualBreakCount\s*=\s*(["'])\d*\1/, ` manualBreakCount="${manual}"`)
        return { edits: [{ part, start: e.start, end: e.end, text: `${tag}${kept.join('')}</${e.qname}>` }] }
      },
    })
  }

  /** Protected ranges (`<protectedRange sqref>`, and Excel 2010's `<xm:sqref>` in the sheet's extension): each area moves, a deleted one goes, a range with none left goes. */
  function protectedRanges(xml: string, e: Span, i: number, part: string) {
    const what: KeptPart = { part: 'protectedRanges', sheet: names[i] }
    for (const p of within(xml, e.start, e.end, 'protectedRange')) {
      const x14 = within(xml, p.open, p.close, 'sqref', true)[0]
      const span = valueAt(xml, p, 'sqref') ?? (x14 && x14.body !== null ? ([x14.open, x14.close] as [number, number]) : undefined)
      const ref = span ? sqrefCarrier(names[i], sqrefRects(readText(xml.slice(...span)))) : null
      if (!span || !ref) {
        out.unparsed[i].push(what)
        continue
      }
      out.items.push({
        what,
        on: i,
        carriers: [carrier(ref, 'place')],
        fix([now]) {
          const left = (rectsNow(now) ?? []).filter((x): x is Rect => x !== null)
          if (!left.length) return { edits: [{ part, start: p.start, end: p.end, text: '' }] }
          return { edits: [{ part, start: span[0], end: span[1], text: left.map(plainArea).join(' ') }] }
        },
      })
    }
  }

  /** Sparklines (the sheet's extension `e`): each one's data (`xm:f`, a formula) and its cell (`xm:sqref`); a cell deleted is refused. */
  function sparklines(xml: string, e: Span, i: number, part: string) {
    const what: KeptPart = { part: 'sparklines', sheet: names[i] }
    for (const f of within(xml, e.start, e.end, 'f')) formula(part, [f.open, f.close], readText(f.body ?? ''), what, escapeText, i)
    for (const q of within(xml, e.start, e.end, 'sqref')) {
      const ref = sqrefCarrier(names[i], sqrefRects(readText(q.body ?? '')))
      if (!ref) {
        out.unparsed[i].push(what)
        continue
      }
      out.items.push({
        what,
        on: i,
        carriers: [carrier(ref, 'place')],
        fix([now]) {
          const xs = rectsNow(now)
          if (!xs || xs.some((x) => x === null)) return { refuse: { key: 'keptSparkline', sheet: names[i] }, code: 'deleted' }
          return { edits: [{ part, start: q.open, end: q.close, text: (xs as Rect[]).map(plainArea).join(' ') }] }
        },
      })
    }
  }

  /** Comment threads by their cells: the model writes the notes it reads them as, not the threads, so a move of one is refused. */
  function threads(xml: string, i: number, part: string) {
    const what: KeptPart = { part: 'threads', sheet: names[i] }
    for (const t of within(xml, 0, xml.length, 'threadedComment')) {
      const ref = sqrefCarrier(names[i], sqrefRects(attributes(t.attrs).ref ?? ''))
      if (!ref) {
        out.unparsed[i].push(what)
        continue
      }
      out.items.push({ what, on: i, carriers: [carrier(ref, 'place')], fix: () => ({ refuse: { key: 'keptWouldMove', what, path: part }, code: 'moved' }) })
    }
  }
}

/** Edits of one element where some lie inside others (a criterion inside its filter): the inner ones dropped where an outer one removes them. */
function nested(edits: Edit[]): Edit[] {
  return edits.filter((x) => !edits.some((y) => y !== x && y.text === '' && y.start <= x.start && x.end <= y.end && (y.start !== x.start || y.end !== x.end)))
}

/** Whether a formula refers to a cell of the book's sheets (`known`, by key): references to other workbooks and names are not moved by its structure. */
function refersToBook(ast: Ast, known: Set<string>): boolean {
  let found = false
  walk(ast, (x) => {
    if ((x.k === 'ref' && x.sheet !== null && known.has(sheetKey(x.sheet))) || (x.k === 'ref3d' && known.has(sheetKey(x.first)))) found = true
  })
  return found
}

/** The carriers as the book's names, hidden ones. */
export function carrierNames(c: Carriers): NameInput[] {
  return [...c.sheets, ...c.cells.flat().map((x) => x.carrier), ...c.formats.flatMap((x) => x ?? []).map((x) => x.carrier), ...c.items.flatMap((x) => x.carriers)].map((x) => ({ name: x.name, ref: x.ref, hidden: true }))
}

/** A carrier's reference as it compares: its canonical text; of a part on a sheet, without the sheet. Null when it does not parse. */
export function compared(text: string, kind: Carrier['kind']): string | null {
  let ast: Ast
  try {
    ast = parse(text, { union: true })
  } catch (e) {
    if (e instanceof FormulaError) return null
    throw e
  }
  return kind === 'ref' ? canonical(ast) : canonical(ast, () => '')
}

/**
 * The sheet a sheet's carrier names now, whether its cell is still the last one, and the lines deleted where none were
 * inserted (its cell came off the last row or column but is on the sheet, which an inserted line pushes it off): rows
 * (`r`), or columns alone (`c`); undefined when the sheet is gone.
 */
export function sheetOfCarrier(text: string): { sheet: string; moved: boolean; deleted?: 'r' | 'c' } | undefined {
  let ast: Ast
  try {
    ast = parse(text)
  } catch (e) {
    if (e instanceof FormulaError) return undefined
    throw e
  }
  if (ast.k === 'ref' && ast.sheet !== null) {
    const [rows, cols] = [ast.ref.r1 !== MAX_ROWS - 1, ast.ref.c1 !== MAX_COLS - 1]
    return { sheet: ast.sheet, moved: rows || cols, ...(rows || cols ? { deleted: rows ? 'r' : 'c' } : {}) }
  }
  if (ast.k === 'err' && ast.sheet !== undefined && ast.book === undefined) return { sheet: ast.sheet, moved: true }
  return undefined
}
