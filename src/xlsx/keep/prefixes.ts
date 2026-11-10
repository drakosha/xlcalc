/**
 * `writeKept`'s `prefixes: 'repair'`: the formulas of a file that call newer functions without their prefix (the reader's
 * `noPrefix`) written again with it, in the text the file holds them in (`formulas/repair.ts`), everything else of the
 * part as it is. Where the reader finds such calls, these find them: cells' formulas, a shared formula's master (its
 * copies take the master's text, so they stay shared, as Excel saves a shared formula with its prefixes), an array
 * formula of the old kind, conditional formats and their thresholds, data validation, Excel 2010's of both in the
 * sheet's extensions, a table's calculated and totals formulas, and defined names. Each call repaired is reported where
 * the reader reports it, under the code `noPrefix`.
 */
import { formatAddress, colToIndex } from '../../model.js'
import { areaText } from '../../structure/area.js'
import { repairPrefixes } from '../formulas/repair.js'
import type { Issues, Where, XlsxMessage } from '../issues.js'
import { fileRange, sqrefRanges } from '../read/refs.js'
import { attributes } from '../xml/scan.js'
import { escapeAttr, escapeText, readText } from '../xml/text.js'

const REPAIRED: XlsxMessage = { key: 'noPrefixRepaired' }

interface Splice {
  start: number
  end: number
  text: string
}

const spliced = (xml: string, edits: Splice[]): string => {
  if (!edits.length) return xml
  let out = ''
  let at = 0
  for (const e of edits.sort((a, b) => a.start - b.start)) {
    out += xml.slice(at, e.start) + e.text
    at = e.end
  }
  return out + xml.slice(at)
}

/** The formula elements whose text is a formula: a cell's, a rule's, a validation's, Excel 2010's (`xm:f`), a table column's. */
const FORMULA = /<((?:[\w.-]+:)?)(f|formula|formula1|formula2|calculatedColumnFormula|totalsRowFormula)(?=[\s>])([^>]*)>([^<]*)<\/\1\2>/g

/** Each formula element of `xml` from `from` to `to` repaired: its edit, and the functions given their prefix. */
function formulas(xml: string, from: number, to: number, edits: Splice[], each: (fns: string[], local: string) => void): void {
  FORMULA.lastIndex = from
  for (let m = FORMULA.exec(xml); m && m.index < to; m = FORMULA.exec(xml)) {
    const r = repairPrefixes(readText(m[4]))
    if (!r) continue
    const at = m.index + m[0].length - m[4].length - `</${m[1]}${m[2]}>`.length
    edits.push({ start: at, end: at + m[4].length, text: escapeText(r.formula) })
    each(r.fns, m[2])
  }
}

const report = (issues: Issues, fns: string[], where: Where) => {
  for (const fn of fns) issues.add('noPrefix', REPAIRED, { ...where, fn })
}

/** An element where it lies: its opening tag, where its body begins and ends (the same for an empty one). */
interface Block {
  tag: string
  open: number
  end: number
}

/** The elements of a local name in `xml` from `from` to `to`, not nested in one another, empty ones among them. */
function blocks(xml: string, local: string, from = 0, to = xml.length): Block[] {
  const out: Block[] = []
  const re = new RegExp(`<((?:[\\w.-]+:)?)${local}(?=[\\s/>])[^>]*>`, 'g')
  re.lastIndex = from
  for (let m = re.exec(xml); m && m.index < to; m = re.exec(xml)) {
    const open = m.index + m[0].length
    if (m[0].endsWith('/>')) {
      out.push({ tag: m[0], open, end: open })
      continue
    }
    const end = xml.indexOf(`</${m[1]}${local}>`, open)
    if (end < 0) break
    out.push({ tag: m[0], open, end })
    re.lastIndex = end
  }
  return out
}

const attrsOfTag = (tag: string) => attributes(tag.slice(tag.search(/[\s/>]/), -1))

/** A rule's ranges: its `sqref`, or Excel 2010's `<xm:sqref>`. */
function sqrefOf(xml: string, b: Block): string {
  const own = attrsOfTag(b.tag).sqref
  if (own !== undefined) return own
  const m = /<(?:[\w.-]+:)?sqref>([^<]*)</.exec(xml.slice(b.open, b.end))
  return m ? readText(m[1]) : ''
}

const ROW_OR_CELL = /<((?:[\w.-]+:)?)(row|c)(?=[\s/>])([^>]*?)(\/?)>/g
const CELL_F = /<((?:[\w.-]+:)?)f(?=[\s/>])([^>]*?)(?:\/>|>([^<]*)<\/\1f>)/

/** The formulas of a sheet's cells (`<sheetData>`): their edits, and the cells repaired, a shared formula's copies with their master. */
function cells(xml: string, sheet: string, issues: Issues, edits: Splice[]): void {
  const data = blocks(xml, 'sheetData')[0]
  if (!data) return
  const masters = new Map<string, string[]>()
  // the copies of each shared formula, the formulas in the order they are first met
  const copies = new Map<string, string[]>()
  const copiesOf = (si: string) => copies.get(si) ?? copies.set(si, []).get(si)!
  let row = -1
  let col = -1
  ROW_OR_CELL.lastIndex = data.open
  for (let m = ROW_OR_CELL.exec(xml); m && m.index < data.end; m = ROW_OR_CELL.exec(xml)) {
    const a = attributes(m[3])
    if (m[2] === 'row') {
      row = a.r !== undefined && /^\s*\d+\s*$/.test(a.r) ? Number(a.r) - 1 : row + 1
      col = -1
      continue
    }
    const ref = a.r && /^([A-Za-z]{1,3})(\d+)$/.exec(a.r)
    col = ref ? colToIndex(ref[1].toUpperCase()) : col + 1
    if (ref) row = Number(ref[2]) - 1
    if (m[4]) continue
    const open = m.index + m[0].length
    const close = xml.indexOf(`</${m[1]}c>`, open)
    if (close < 0 || close > data.end) break
    ROW_OR_CELL.lastIndex = close
    const f = CELL_F.exec(xml.slice(open, close))
    if (!f) continue
    const fa = attributes(f[2])
    const addr = formatAddress(row, col)
    if (!f[3]) {
      if (fa.t === 'shared' && fa.si !== undefined) copiesOf(fa.si).push(addr)
      continue
    }
    if (fa.t === 'shared' && fa.si !== undefined) copiesOf(fa.si)
    if (fa.t === 'dataTable') continue
    const r = repairPrefixes(readText(f[3]))
    if (!r) continue
    const at = open + f.index + f[0].length - f[3].length - `</${f[1]}f>`.length
    edits.push({ start: at, end: at + f[3].length, text: escapeText(r.formula) })
    report(issues, r.fns, { sheet, cell: addr })
    if (fa.t === 'shared' && fa.si !== undefined && !masters.has(fa.si)) masters.set(fa.si, r.fns)
  }
  // a shared formula's copies take the master's text, repaired with it: reported as the reader reports them
  for (const [si, list] of copies) {
    const fns = masters.get(si)
    if (fns) for (const addr of list) report(issues, fns, { sheet, cell: addr })
  }
}

/** A worksheet's part with its formulas repaired (`issues` given the calls on `sheet`, the workbook's name of it); the text itself where none is. */
export function repairSheet(xml: string, sheet: string, issues: Issues): string {
  const edits: Splice[] = []
  cells(xml, sheet, issues, edits)
  // conditional formats, 2007's and 2010's, at the first of their ranges, as the reader reports them
  for (const b of blocks(xml, 'conditionalFormatting')) {
    const sqref = sqrefOf(xml, b)
    const first = sqrefRanges(sqref).find((x) => x.rect)
    const where = { sheet, cell: first ? areaText(first.rect!) : sqref }
    formulas(xml, b.open, b.end, edits, (fns) => report(issues, fns, where))
    // a threshold of 2007's color scales, data bars and icon sets, in an attribute
    for (const m of xml.slice(b.open, b.end).matchAll(/<(?:[\w.-]+:)?cfvo\b[^>]*?\sval="([^"]*)"/g)) {
      const r = repairPrefixes(readText(m[1]))
      if (!r) continue
      const at = b.open + m.index + m[0].length - 1 - m[1].length
      edits.push({ start: at, end: at + m[1].length, text: escapeAttr(r.formula) })
      report(issues, r.fns, where)
    }
  }
  for (const b of blocks(xml, 'dataValidation')) {
    const where = { sheet, cell: sqrefOf(xml, b) }
    formulas(xml, b.open, b.end, edits, (fns) => report(issues, fns, where))
  }
  return spliced(xml, edits)
}

/** A table's part with its columns' formulas repaired, reported at the column's first cell, as the reader reports them (a totals formula at its last). */
export function repairTable(xml: string, sheet: string, issues: Issues): string {
  const root = blocks(xml, 'table')[0]
  const x = root ? fileRange(attrsOfTag(root.tag).ref ?? '') : null
  if (!x) return xml
  const edits: Splice[] = []
  blocks(xml, 'tableColumn').forEach((b, i) => {
    formulas(xml, b.open, b.end, edits, (fns, local) => report(issues, fns, { sheet, cell: formatAddress(local === 'totalsRowFormula' ? x.r2 : x.r1, x.c1 + i) }))
  })
  return spliced(xml, edits)
}

/** The workbook's part with its defined names repaired, reported by name; Excel's placeholders of newer functions (`_xlfn.` names) are none. */
export function repairNames(xml: string, issues: Issues): string {
  const edits: Splice[] = []
  for (const b of blocks(xml, 'definedName')) {
    const name = readText(attrsOfTag(b.tag).name ?? '')
    const raw = xml.slice(b.open, b.end)
    if (/^_xlfn\./i.test(name) || raw.includes('<')) continue
    const r = repairPrefixes(readText(raw))
    if (!r) continue
    edits.push({ start: b.open, end: b.end, text: escapeText(r.formula) })
    report(issues, r.fns, { name })
  }
  return spliced(xml, edits)
}
