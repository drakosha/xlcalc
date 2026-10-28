/**
 * A worksheet written over the original's part. The writer writes the sheet twice, as the book read from the file has
 * it (`p0`) and as the book written has it (`p1`), and where the two agree the file's XML stays: an element of the
 * sheet the model has, a row, a cell. Where they differ the model's XML goes in, except what the file has and the model
 * does not: a cell whose formula did not change keeps its `<f>` (a shared formula, a data table, a formula the engine
 * does not parse) and takes the new value; a cell, a row or a column keeps a cell format of the file that looks the
 * same, where rows or columns moved the format of the block it lies in now (`formats.ts`), and a row the attributes the
 * model does not have. A cell with what only the file holds that moved (rows or columns
 * inserted or deleted, a range moved) takes the file's XML at its new address where the model has it as it was. The
 * elements the model does not have stay in their place, in the order of the schema; an element the model adds goes
 * where the schema puts it.
 */
import type { CellInput, SheetInput, SheetProtection } from '../../contract.js'
import { colToIndex, formatAddress, parseAddress } from '../../model.js'
import { keyOf } from '../../style-table.js'
import type { RowWritten, SheetPieces } from '../write/sheet.js'
import { MAIN } from '../write/xml.js'
import { attributes } from '../xml/scan.js'
import { parseRange } from '../../structure/area.js'
import { protectionAttrs } from '../protection.js'
import { KeptRefusal, type Moved } from './refusals.js'
import type { Formats } from './formats.js'
import type { KeptStyles } from './styles.js'
import { attrsOf, childrenOf, closing, DOC_REL, opening, rootOf, setAttrs, split, tagEnd, type Placed } from './xml.js'

/** The children of a worksheet in the schema's order. */
const ORDER = [
  'sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter',
  'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions',
  'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing',
  'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst',
]

/** The extensions of a worksheet that hold the model's conditional formats and data validation. */
export const EXT_CF = '{78C0D931-6437-407D-A8EE-F0AAD7539E65}'
export const EXT_DV = '{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}'

/** Attributes of a row that the model does not have and that a row written again keeps. */
const ROW_KEPT = ['x14ac:dyDescent', 'thickTop', 'thickBot', 'ph']

export interface SpliceInput {
  xml: string
  p0: SheetPieces
  p1: SheetPieces
  s0: SheetInput
  s1: SheetInput
  styles: KeptStyles
  /** The elements bound to other parts, as the book's writer decided them: the new text (empty for none), undefined to keep the file's. */
  bound: { hyperlinks?: string; legacyDrawing?: string; tableParts?: string }
  /** Rows or columns were inserted or deleted on the sheet: the ignored errors, whose cells no carrier holds, are left out. */
  lines: boolean
  /** The file's cells with what only it holds that moved (`refusals.ts`). */
  moved?: Moved
  /** The sheet's conditional formats written over the file's (`rules.ts`), in place of the workbook's. */
  conditional?: { main: string; ext: string }
  /** The rectangles of the file's cell formats that look as earlier ones where they lie now (`formats.ts`); none past their limit. */
  formats?: Formats
}

export interface Spliced {
  /** The new part; undefined when it stays as the file has it. */
  xml?: string
  /** Something the model has changed, not only values of formulas. */
  changed: boolean
  /** The cells written from the model (anew, or with a new value), for the issues of writing them. */
  cells: Set<string>
  /** The elements written from the model. */
  written: Set<string>
}

/** A cell of the model's XML: its parts as the writer writes them. */
interface Gen {
  text: string
  s: number
  /** `t` and `vm` as written. */
  t: string
  vm?: string
  /** `cm` and the formula element: what counts as the cell's formula. */
  f: string
  value: string
}

/** The address of a cell the writer wrote, and of a file's cell. */
const GEN_ADDR = /^<c r="[A-Z]+\d+"/
const SRC_ADDR = /^(<(?:[\w.-]+:)?c\b[^>]*?\sr=")[^"]*"/

const GEN = /<c r="([A-Z]+)(\d+)"(?: s="(\d+)")?(?: t="(\w+)")?(?: vm="(\d+)")?( cm="\d+")?(?:\/>|>(<f[^>]*?(?:\/>|>[^<]*<\/f>))?(<v>[^<]*<\/v>|<is>[\s\S]*?<\/is>)?<\/c>)/g

function genCells(cells: string): Map<number, Gen> {
  const out = new Map<number, Gen>()
  GEN.lastIndex = 0
  for (let m = GEN.exec(cells); m; m = GEN.exec(cells)) out.set(colToIndex(m[1]), { text: m[0], s: Number(m[3] ?? 0), t: m[4] ?? '', ...(m[5] ? { vm: m[5] } : {}), f: (m[6] ?? '') + (m[7] ?? ''), value: m[8] ?? '' })
  // the writer's cells all read: one missed would keep the file's cell in its place
  if (out.size !== cells.split('<c r="').length - 1) throw new Error('xlsx/keep: a cell the writer wrote is not of the form read back')
  return out
}

/** A cell of the file: its text, cell format and formula element. */
interface Src {
  addr: string
  col: number
  text: string
  s?: number
  /** Its formula element as written, `<f …>…</f>`. */
  f?: string
  /** A shared formula's group, and whether the cell is its master (the one with the text). */
  si?: string
  master?: boolean
}

interface SrcRow {
  r: number
  start: number
  end: number
  /** The opening tag. */
  tag: string
  bodyStart: number
  bodyEnd: number
}

const OPEN = (local: string) => new RegExp(`<(?:[\\w.-]+:)?${local}(?=[\\s/>])`, 'g')
const ROW = OPEN('row')
const CELL = OPEN('c')
const F = /<((?:[\w.-]+:)?)f(?=[\s/>])/

/** The rows of `<sheetData>` in order, or null when they are out of order (a row given twice or after a later one). */
function srcRows(xml: string, from: number, to: number): SrcRow[] | null {
  const out: SrcRow[] = []
  ROW.lastIndex = from
  let r = -1
  for (let m = ROW.exec(xml); m && m.index < to; m = ROW.exec(xml)) {
    const open = tagEnd(xml, m.index)
    if (open < 0) return null
    const tag = xml.slice(m.index, open)
    const a = attributes(tag.slice(m[0].length))
    const next = a.r !== undefined && /^\s*\d+\s*$/.test(a.r) ? Number(a.r) - 1 : r + 1
    if (next <= r) return null
    r = next
    if (tag.endsWith('/>')) {
      out.push({ r, start: m.index, end: open, tag, bodyStart: open, bodyEnd: open })
      ROW.lastIndex = open
      continue
    }
    const close = `</${m[0].slice(1, -3)}row>`
    const end = xml.indexOf(close, open)
    if (end < 0 || end > to) return null
    out.push({ r, start: m.index, end: end + close.length, tag, bodyStart: open, bodyEnd: end })
    ROW.lastIndex = end + close.length
  }
  return out
}

/** A row's cells by column, or null when they are out of order. */
function srcCells(xml: string, row: SrcRow): Map<number, Src> | null {
  const out = new Map<number, Src>()
  CELL.lastIndex = row.bodyStart
  let col = -1
  for (let m = CELL.exec(xml); m && m.index < row.bodyEnd; m = CELL.exec(xml)) {
    const open = tagEnd(xml, m.index)
    if (open < 0) return null
    const tag = xml.slice(m.index, open)
    let end = open
    if (!tag.endsWith('/>')) {
      const close = `</${m[0].slice(1, -1)}c>`
      const at = xml.indexOf(close, open)
      if (at < 0 || at > row.bodyEnd) return null
      end = at + close.length
    }
    const a = attributes(tag.slice(m[0].length))
    const ref = a.r && /^([A-Za-z]{1,3})(\d+)$/.exec(a.r)
    const next = ref ? colToIndex(ref[1].toUpperCase()) : col + 1
    if (next <= col || (ref && Number(ref[2]) - 1 !== row.r)) return null
    col = next
    const text = xml.slice(m.index, end)
    const cell: Src = { addr: '', col, text, ...(a.s !== undefined ? { s: Number(a.s) } : {}) }
    const fm = F.exec(text.slice(tag.length))
    if (fm) {
      const at = tag.length + fm.index
      const fOpen = tagEnd(text, at)
      const fTag = text.slice(at, fOpen)
      const fEnd = fTag.endsWith('/>') ? fOpen : text.indexOf(`</${fm[1]}f>`, fOpen) + fm[1].length + 4
      cell.f = text.slice(at, fEnd)
      const fa = attributes(fTag.slice(fm[0].length))
      if (fa.t === 'shared' && fa.si !== undefined) {
        cell.si = fa.si
        cell.master = !fTag.endsWith('/>') && fEnd - fOpen > fm[1].length + 4
      }
    }
    out.set(col, cell)
    CELL.lastIndex = end
  }
  return out
}

/** A file's cell with a new value (`g`'s) and cell format `s`, its formula element and other attributes as the file has them. */
function patched(src: Src, g: Gen, s: number): string {
  const tag = src.text.slice(0, tagEnd(src.text, 0))
  const open = opening(setAttrs(tag, { s: s ? String(s) : null, t: g.t || null, vm: g.vm ?? null }))
  const name = /^<([\w.:-]+)/.exec(tag)![1]
  return `${open}${src.f ?? ''}${g.value}</${name}>`
}

/** A file's cell with cell format `s` in place of its own. */
function withFormat(src: Src, s: number): string {
  const end = tagEnd(src.text, 0)
  return setAttrs(src.text.slice(0, end), { s: s ? String(s) : null }) + src.text.slice(end)
}

/** A row's cell format as the file gives it (`s` counts with `customFormat`), 0 for none. */
const rowFormat = (tag: string) => {
  const a = attributes(tag.slice(tag.search(/\s/)))
  return a.customFormat === '1' || a.customFormat === 'true' ? Number(a.s) || 0 : 0
}

/** A cell input as it counts for whether the model changed it: its value, formula and look, not its note, link or meta. */
const inputKey = (c: CellInput | undefined) => (c ? keyOf({ value: c.value, formula: c.formula, array: c.array, style: c.style, numberFormat: c.numberFormat }) : '')

interface DataOut {
  text: string
  changed: boolean
}

/**
 * `<sheetData>` over the file's: rows where the model's two versions agree as the file has them, the others cell by
 * cell. A shared formula whose master is written anew is written anew in its copies too. Null when the file's rows or
 * cells are out of order: the model's rows are written then.
 */
function sheetData(xml: string, data: Placed, x: SpliceInput, cells: Set<string>): DataOut | null {
  const rows = data.end === data.open ? [] : srcRows(xml, data.open, data.end)
  if (!rows) return null
  const g0 = new Map(x.p0.rows.map((r) => [r.row, r]))
  const g1 = new Map(x.p1.rows.map((r) => [r.row, r]))
  const src = new Map(rows.map((r) => [r.r, r]))
  const same = (r: number) => {
    const [a, b] = [g0.get(r), g1.get(r)]
    return a === b || (a !== undefined && b !== undefined && a.attrs === b.attrs && a.cells === b.cells)
  }
  const all = [...new Set([...src.keys(), ...g0.keys(), ...g1.keys()])].sort((a, b) => a - b)
  let changed = false
  // the groups of shared formulas whose master was written anew, and the rows that hold their copies
  const broken = new Set<string>()
  const parsed = new Map<number, Map<number, Src>>()
  const cellsOf = (r: SrcRow) => {
    let m = parsed.get(r.r)
    if (!m) {
      const got = srcCells(xml, r)
      if (!got) throw OUT_OF_ORDER
      parsed.set(r.r, (m = got))
    }
    return m
  }
  const rowOut = new Map<number, string>()
  const gen0 = new Map<number, Map<number, Gen>>()
  /** The file's cell that moved to `at` from `from`, at its new address, where the model's cell there (`g`) is as it was. */
  const carried = (from: string, at: string, g: Gen): string | undefined => {
    const { row, col } = parseAddress(from)
    const [s, a] = [src.get(row), g0.get(row)]
    if (!s || !a) return undefined
    let cells = gen0.get(row)
    if (!cells) gen0.set(row, (cells = genCells(a.cells)))
    const [own, was] = [cellsOf(s).get(col), cells.get(col)]
    // a shared formula's cell means another formula elsewhere
    if (!own || !was || own.si !== undefined || was.text.replace(GEN_ADDR, '') !== g.text.replace(GEN_ADDR, '')) return undefined
    return own.text.replace(SRC_ADDR, `$1${at}"`)
  }
  // with rows or columns inserted or deleted, a cell or row takes the format of the rectangle it lies in now, not the file's at its address
  const placed = x.lines && x.formats ? x.formats.index(all) : undefined
  // the rows a cell with what only the file holds left or came to
  const touched = new Set([...(x.moved?.left ?? []), ...(x.moved?.to.keys() ?? [])].map((a) => parseAddress(a).row))
  const merge = (r: number): string => {
    const s = src.get(r)
    const [a, b] = [g0.get(r), g1.get(r)]
    const sc = s ? cellsOf(s) : new Map<number, Src>()
    const [c0, c1] = [a ? genCells(a.cells) : new Map<number, Gen>(), b ? genCells(b.cells) : new Map<number, Gen>()]
    const cols = [...new Set([...sc.keys(), ...c0.keys(), ...c1.keys()])].sort((p, q) => p - q)
    const xfOf = (own: Src | undefined, col: number, xf: number) => (placed ? placed.cell(r, col, xf) : x.styles.keep(own?.s, xf))
    let body = ''
    let rowChanged = false
    for (const col of cols) {
      const [x0, x1] = [c0.get(col), c1.get(col)]
      let own = sc.get(col)
      const at = x.moved && formatAddress(r, col)
      // the file's cell went elsewhere, with what only it holds: the model's goes here
      const gone = !!at && !!own && x.moved!.left.has(at)
      if (gone) {
        if (own!.master) broken.add(own!.si!)
        own = undefined
      }
      const from = at && x.moved!.to.get(at)
      const moved = from && x1 ? carried(from, at, x1) : undefined
      if (moved !== undefined) {
        body += moved
        rowChanged = true
        changed = true
        continue
      }
      const forced = gone || (own?.si !== undefined && !own.master && broken.has(own.si))
      if (x0?.text === x1?.text && !forced) {
        // a cell of the file the model has not (one that looks as its row or column) means nothing where lines moved
        const s1 = placed && own && x1 ? placed.cell(r, col, x1.s) : (own?.s ?? 0)
        if (own && (x0 || !x.lines) && s1 === (own.s ?? 0)) body += own.text
        else if (own && x0) {
          // the file's cell where another moved to it, with the format of the one that moved
          x.styles.use('cellXfs', s1)
          body += withFormat(own, s1)
          rowChanged = changed = true
        } else if (own) rowChanged = true
        continue
      }
      rowChanged = true
      const addr = (x1 ?? x0)?.text.slice(6, (x1 ?? x0)!.text.indexOf('"', 6))
      const inputs = addr === undefined || inputKey(x.s0.cells[addr]) !== inputKey(x.s1.cells[addr])
      if (!x1) {
        // a cell an array no longer spills into changes no input; a cell of a look alone does
        changed ||= inputs || (x0 !== undefined && !x0.f && !x0.value)
        if (own?.master) broken.add(own.si!)
        if (addr) cells.add(addr)
        continue
      }
      cells.add(addr!)
      // the same formula keeps the file's formula element, a value or a look of its own changes alone
      if (own && x0 && x0.f === x1.f && !forced) {
        const s1 = xfOf(own, col, x1.s)
        if (s1 !== own.s) x.styles.use('cellXfs', s1)
        changed ||= inputs || s1 !== (own.s ?? 0)
        body += patched(own, x1, s1)
        continue
      }
      changed = true
      if (own?.master) broken.add(own.si!)
      const s1 = xfOf(own, col, x1.s)
      x.styles.use('cellXfs', s1)
      body += s1 === x1.s ? x1.text : x1.text.replace(/^(<c r="\w+")(?: s="\d+")?/, s1 ? `$1 s="${s1}"` : '$1')
    }
    const attrsChanged = (a?.attrs ?? '') !== (b?.attrs ?? '')
    // the row's format where rows or columns moved: that of the rows it lies in now
    const format = placed && b ? placed.row(r, Number(/ s="(\d+)"/.exec(b.attrs)?.[1] ?? 0)) : undefined
    const reformat = format !== undefined && format !== (s ? rowFormat(s.tag) : 0)
    if (!rowChanged && !attrsChanged && !reformat && s) return xml.slice(s.start, s.end)
    if (attrsChanged || reformat) changed = true
    // a row the model no longer has, its cells gone, goes
    return !b && !body ? '' : rowXml(s, a, b, body, x, reformat ? format : undefined)
  }
  try {
    // with rows or columns inserted or deleted, every row is looked at for the file's cells the model has not
    for (const r of all) if (x.lines || !same(r) || touched.has(r)) rowOut.set(r, merge(r))
    // the copies of shared formulas whose master was written anew, in the rows kept as they were
    if (broken.size)
      for (const s of rows) {
        if (rowOut.has(s.r)) continue
        const body = xml.slice(s.bodyStart, s.bodyEnd)
        if (![...broken].some((si) => body.includes(`si="${si}"`))) continue
        if ([...cellsOf(s).values()].some((c) => c.si !== undefined && !c.master && broken.has(c.si))) rowOut.set(s.r, merge(s.r))
      }
  } catch (e) {
    if (e === OUT_OF_ORDER) return null
    throw e
  }
  if (!rowOut.size) return { text: xml.slice(data.start, data.end), changed: false }
  let body = ''
  for (const r of all) {
    const out = rowOut.get(r)
    if (out !== undefined) body += out
    else {
      const s = src.get(r)
      if (s) body += xml.slice(s.start, s.end)
    }
  }
  return { text: `${opening(xml.slice(data.start, data.open))}${body}</${data.qname}>`, changed }
}

const OUT_OF_ORDER = Symbol('out of order')

/**
 * A row written from the model's cells (`body`): the file's tag where its attributes did not change, else the model's
 * with what the file's has that the model does not; `format`, the cell format it takes in place of the one it has.
 */
function rowXml(s: SrcRow | undefined, a: RowWritten | undefined, b: RowWritten | undefined, body: string, x: SpliceInput, format?: number): string {
  let tag: string
  if (format !== undefined) x.styles.use('cellXfs', format)
  const formatted = (s: number) => (s ? { s: String(s), customFormat: '1' } : { s: null, customFormat: null })
  if (s && (a?.attrs ?? '') === (b?.attrs ?? '')) tag = setAttrs(s.tag, { spans: null, ...(format === undefined ? {} : formatted(format)) })
  else {
    const own = s ? attributes(s.tag.slice(s.tag.search(/\s/))) : {}
    let attrs = b?.attrs ?? ''
    const sm = / s="(\d+)"/.exec(attrs)
    if (sm) {
      const kept = format ?? x.styles.keep(own.s === undefined ? undefined : Number(own.s), Number(sm[1]))
      x.styles.use('cellXfs', kept)
      attrs = attrs.replace(sm[0], ` s="${kept}"`)
    } else if (format) attrs += ` s="${format}" customFormat="1"`
    const rest = ROW_KEPT.filter((k) => own[k] !== undefined).map((k) => ` ${k}="${own[k]}"`).join('')
    const name = s ? /^<([\w.:-]+)/.exec(s.tag)![1] : 'row'
    const r = (s?.r ?? a?.row ?? b?.row)! + 1
    tag = `<${name} r="${r}"${attrs}${rest}>`
  }
  const name = /^<([\w.:-]+)/.exec(tag)![1]
  return body ? `${opening(tag)}${body}</${name}>` : closing(tag)
}

/** The model's rows alone, cell formats marked used: a file whose rows or cells are out of order. */
function modelData(x: SpliceInput): string {
  for (const r of x.p1.rows) for (const m of `${r.attrs}${r.cells}`.matchAll(/ s="(\d+)"/g)) x.styles.use('cellXfs', Number(m[1]))
  return `<sheetData>${x.p1.rows.map((r) => (r.cells ? `<row r="${r.row + 1}"${r.attrs}>${r.cells}</row>` : `<row r="${r.row + 1}"${r.attrs}/>`)).join('')}</sheetData>`
}

/** An opening tag's attributes set from another tag's: those of `names` the other has, the rest removed. */
function attrsFrom(tag: string, other: string, names: string[]): string {
  const a = attributes(other.slice(other.search(/[\s/>]/)))
  return setAttrs(tag, Object.fromEntries(names.map((n) => [n, a[n] ?? null])))
}

/** The first element of a local name in a text, where it lies. */
function find(xml: string, local: string, from = 0, to = xml.length): Placed | undefined {
  const re = OPEN(local)
  re.lastIndex = from
  const m = re.exec(xml)
  if (!m || m.index >= to) return undefined
  const root = rootOf(xml.slice(m.index, to))
  return root && { ...root, start: root.start + m.index, open: root.open + m.index, end: root.end + m.index }
}

/** An attribute of an element's text, undefined without the element. */
const attrOf = (text: string | undefined, name: string) => (text === undefined ? undefined : attributes(text.slice(text.indexOf(' ') + 1, tagEnd(text, 0)))[name])

/** Whether two ranges are the same cells (`A:C` and `A1:C1048576`), or both are none. */
const sameRange = (a: string | undefined, b: string | undefined) => {
  const [x, y] = [a, b].map((t) => (t === undefined ? null : parseRange(t)))
  return x === null || y === null ? x === y && a === b : x.r1 === y.r1 && x.r2 === y.r2 && x.c1 === y.c1 && x.c2 === y.c2
}

export function spliceSheet(x: SpliceInput): Spliced {
  const { xml, p0, p1 } = x
  const root = rootOf(xml)!
  const kids = childrenOf(xml, root)
  const items = kids.map((c) => ({ local: c.local, text: xml.slice(c.start, c.end) }))
  const cells = new Set<string>()
  const written = new Set<string>()
  let changed = false
  let touched = false
  /** All the elements of a local name replaced by one text (none when empty), at the first's place or where the schema puts it. */
  const put = (local: string, text: string) => {
    touched = true
    written.add(local)
    let at = items.findIndex((e) => e.local === local)
    if (at < 0) {
      const rank = ORDER.indexOf(local)
      at = 0
      items.forEach((e, k) => {
        const r = ORDER.indexOf(e.local)
        if (r >= 0 && r <= rank) at = k + 1
      })
    }
    for (let k = items.length - 1; k >= 0; k--)
      if (items[k].local === local) {
        items.splice(k, 1)
        if (k < at) at--
      }
    if (text) items.splice(at, 0, { local, text })
  }
  const own = (local: string) => items.find((e) => e.local === local)
  if (p0.sheetPr !== p1.sheetPr) {
    changed = true
    const e = own('sheetPr')
    put('sheetPr', e ? sheetPr(e.text, p1.sheetPr) : p1.sheetPr)
  }
  if (p0.views !== p1.views) {
    changed = true
    const e = own('sheetViews')
    put('sheetViews', e ? views(e.text, p0.views, p1.views) : p1.views)
  }
  if (p0.formatPr !== p1.formatPr) {
    changed = true
    const e = own('sheetFormatPr')
    put('sheetFormatPr', e ? attrsFrom(e.text, p1.formatPr, ['outlineLevelRow', 'outlineLevelCol']) : p1.formatPr)
  }
  if (p0.cols !== p1.cols) {
    changed = true
    const cols = x.formats ? x.formats.cols(p1.cols) : p1.cols
    for (const m of cols.matchAll(/ style="(\d+)"/g)) x.styles.use('cellXfs', Number(m[1]))
    put('cols', cols)
  }
  const data = kids.find((c) => c.local === 'sheetData')
  const rowsSame = p0.rows.length === p1.rows.length && p0.rows.every((r, i) => r.row === p1.rows[i].row && r.attrs === p1.rows[i].attrs && r.cells === p1.rows[i].cells)
  if (!rowsSame) {
    const out = data ? sheetData(xml, data, x, cells) : null
    if (!out) {
      for (const r of p1.rows) for (const m of r.cells.matchAll(/<c r="(\w+)"/g)) cells.add(m[1])
      changed = true
      put('sheetData', modelData(x))
    } else if (out.text !== xml.slice(data!.start, data!.end)) {
      changed ||= out.changed
      put('sheetData', out.text)
    }
    if (written.has('sheetData')) {
      const d = own('dimension')
      put('dimension', d ? setAttrs(d.text, { ref: p1.dimension }) : `<dimension ref="${p1.dimension}"/>`)
    }
  }
  if (keyOf(x.s0.protection ?? null) !== keyOf(x.s1.protection ?? null)) {
    changed = true
    const e = own('sheetProtection')
    put('sheetProtection', x.s0.protection && x.s1.protection && e ? protectionOver(e.text, x.s1.protection) : p1.protection)
  }
  // the file's AutoFilter, its criteria and sort, as its carrier moved it where the workbook has it there; written from the
  // workbook where the filter mode changed or a host put the AutoFilter elsewhere or removed it
  if (!!x.s0.filtered !== !!x.s1.filtered || ((x.s0.autoFilter || x.s1.autoFilter) && !sameRange(attrOf(own('autoFilter')?.text, 'ref'), x.s1.autoFilter?.range))) {
    changed = true
    put('autoFilter', p1.autoFilter)
  }
  if (p0.merges !== p1.merges) {
    changed = true
    put('mergeCells', p1.merges)
  }
  const ext = { cf: undefined as string | undefined, dv: false }
  const cf = x.conditional ?? (p0.conditional + p0.ext !== p1.conditional + p1.ext ? { main: p1.conditional, ext: p1.ext } : undefined)
  if (cf) {
    changed = true
    for (const m of cf.main.matchAll(/ dxfId="(\d+)"/g)) x.styles.use('dxfs', Number(m[1]))
    put('conditionalFormatting', cf.main)
    ext.cf = cf.ext
  }
  if (p0.validations !== p1.validations) {
    changed = true
    put('dataValidations', p1.validations)
    ext.dv = true
  }
  for (const [local, text] of Object.entries(x.bound))
    if (text !== undefined) {
      changed = true
      put(local, text)
    }
  if (x.lines && own('ignoredErrors')) put('ignoredErrors', '')
  if (ext.cf !== undefined || ext.dv) {
    const e = own('extLst')
    const list = e ? extList(e.text, ext) : ext.cf ? `<extLst>${ext.cf}</extLst>` : ''
    if (list !== (e?.text ?? '')) put('extLst', list)
  }
  if (!touched) return { changed, cells, written }
  const head = xml.slice(0, root.start) + opening(rootTag(xml.slice(root.start, root.open), root.qname))
  const close = root.end === root.open ? `</${root.qname}>` : xml.slice(kids.at(-1)?.end ?? root.open, root.end)
  return { xml: head + items.map((e) => e.text).join('') + close, changed, cells, written }
}

/** The root's tag declaring the namespaces the model's elements are written in: SpreadsheetML by default, `r` for relationships. */
function rootTag(tag: string, qname: string): string {
  const a = attributes(tag.slice(qname.length + 1))
  const add: Record<string, string> = {}
  if (a.xmlns === undefined) add.xmlns = MAIN
  else if (a.xmlns !== MAIN) throw new KeptRefusal({ key: 'keptNamespace', namespace: a.xmlns }, 'file')
  if (a['xmlns:r'] === undefined) add['xmlns:r'] = DOC_REL
  else if (a['xmlns:r'] !== DOC_REL) throw new KeptRefusal({ key: 'keptRelPrefix', namespace: a['xmlns:r'] }, 'file')
  return Object.keys(add).length ? setAttrs(tag, add) : tag
}

/** The file's `<sheetProtection>` with the model's protection: its attributes of what users may do set, its password kept. */
function protectionOver(text: string, p: SheetProtection): string {
  const e = split(text)
  const tag = setAttrs(e.tag, protectionAttrs(p))
  return e.body ? `${opening(tag)}${e.body}</${e.name}>` : closing(tag)
}

/** The file's `<sheetPr>` with the model's filter mode and outline settings. */
function sheetPr(text: string, model: string): string {
  const e = split(text)
  const tag = attrsFrom(e.tag, model, ['filterMode'])
  const mine = /<outlinePr[^>]*\/>/.exec(model)?.[0] ?? '<outlinePr/>'
  const theirs = find(e.body, 'outlinePr')
  let body = e.body
  if (theirs) body = body.slice(0, theirs.start) + attrsFrom(body.slice(theirs.start, theirs.open), mine, ['summaryBelow', 'summaryRight']) + body.slice(theirs.open)
  else if (mine !== '<outlinePr/>') {
    // after the tab colour, the first of the children
    const tab = find(body, 'tabColor')
    body = tab ? body.slice(0, tab.end) + mine + body.slice(tab.end) : mine + body
  }
  return body ? `${opening(tag)}${body}</${e.name}>` : closing(tag)
}

/**
 * The file's `<sheetViews>` with the model's first view: its tab, scroll and zoom, and, where the model's panes or
 * selection changed (`was` and `model`, the writer's views of the book read and written), its panes and selection in
 * place of the file's.
 */
function views(text: string, was: string, model: string): string {
  const first = find(text, 'sheetView', tagEnd(text, 0))
  const mine = find(model, 'sheetView', tagEnd(model, 0))
  const before = find(was, 'sheetView', tagEnd(was, 0))
  if (!first || !mine || !before) return model
  const tag = attrsFrom(text.slice(first.start, first.open), model.slice(mine.start, mine.open), ['tabSelected', 'topLeftCell', 'zoomScale'])
  const body = split(model.slice(mine.start, mine.end)).body
  const view =
    body === split(was.slice(before.start, before.end)).body
      ? (first.end === first.open ? closing(tag) : opening(tag) + text.slice(first.open, first.end))
      : (() => {
          const kept = childrenOf(text, first).filter((c) => c.local !== 'pane' && c.local !== 'selection').map((c) => text.slice(c.start, c.end)).join('')
          return body + kept ? `${opening(tag)}${body}${kept}</${first.qname}>` : closing(tag)
        })()
  return text.slice(0, first.start) + view + text.slice(first.end)
}

/** The file's `<extLst>` with the model's conditional formats of Excel 2010 (`cf`) in place of its, and without its data validation (`dv`). */
function extList(text: string, ext: { cf?: string; dv: boolean }): string {
  const root = rootOf(text)!
  const kids = childrenOf(text, root).map((c) => ({ uri: (attrsOf(text, c).uri ?? '').toUpperCase(), text: text.slice(c.start, c.end) }))
  let out = kids.filter((k) => !(ext.cf !== undefined && k.uri === EXT_CF) && !(ext.dv && k.uri === EXT_DV)).map((k) => k.text)
  if (ext.cf) {
    // Excel 2010's conditional formats first, as Excel writes them
    const at = kids.findIndex((k) => k.uri === EXT_CF)
    out.splice(at < 0 ? 0 : Math.min(at, out.length), 0, ext.cf)
  }
  out = out.filter(Boolean)
  return out.length ? `${text.slice(0, root.open)}${out.join('')}</${root.qname}>` : ''
}
