/**
 * Carriers: names of the book that hold the references of the file's parts the model does not have (`openKept` adds
 * them, `writeKept` reads them back). Each sheet has one on its last cell, which a row or column inserted or deleted
 * anywhere on it moves and a rename or a deletion of the sheet changes, and which tells which sheet of the book the
 * file's sheet became. Each reference of a part the model does not have gets one of its own, which an operation of
 * the structure moves as Excel moves the reference: the series and labels of charts, the anchors of drawings, the
 * source and the place of pivot tables, sparklines, a sheet's AutoFilter and sort, the names the model does not take
 * (the print area and titles), page breaks, protected ranges and comment threads. Parts whose references are not
 * read (`unparsed`) are known only by their sheet's carrier: rows and columns inserted or deleted there are seen, a
 * range moved over them is not.
 */
import type { BookInput, NameInput } from '../../contract.js'
import { caseKey } from '../../collate.js'
import { indexToCol, MAX_COLS, MAX_ROWS } from '../../model.js'
import type { Rect } from '../../structure/area.js'
import type { Ast } from '../../syntax/ast.js'
import { FormulaError } from '../../syntax/ast.js'
import { parse } from '../../syntax/parse.js'
import { canonical, mapAst, quoteSheet, sheetKey, walk } from '../../syntax/print.js'
import { attr, attributes, element, elements } from '../xml/scan.js'
import { readText } from '../xml/text.js'
import { parseFormula } from '../read/formulas.js'
import type { Source } from './source.js'
import { attrsOf, childrenOf, ContentTypes, rootOf } from './xml.js'

/** The prefix of the carriers' names. */
export const PREFIX = '_xlcalc.keep.'

/** A reference of a part the model does not have, as a name of the book. */
export interface Carrier {
  name: string
  ref: string
  /** `ref`: the reference counts as written, its sheet too; `place`: the part lies on the sheet, only its cells count. */
  kind: 'ref' | 'place'
  /** The part, as a refusal names it. */
  what: string
}

export interface Carriers {
  /** Each of the source's sheets' carrier, in the book's order of them. */
  sheets: Carrier[]
  refs: Carrier[]
  /** By sheet of the source: its parts whose references no carrier holds. */
  unparsed: string[][]
  /** Parts whose references no carrier holds and that may refer to any sheet. */
  anywhere: string[]
  /** The tables and names parts take their data from by name (a pivot cache's source). */
  byName: { name: string; what: string }[]
  /** By sheet of the source: parts that name the sheet (a pivot chart's pivot table). */
  named: string[][]
}

/** Elements of a worksheet whose references are not read. */
const UNPARSED = new Set(['oleObjects', 'controls', 'dataConsolidate', 'scenarios', 'customSheetViews', 'cellWatches', 'webPublishItems', 'smartTags', 'AlternateContent'])

/** Extensions of a worksheet without references of cells, or with those the model or a carrier holds. */
const EXT_KNOWN = new Set(
  [
    '{78C0D931-6437-407d-A8EE-F0AAD7539E65}', // conditional formats
    '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}', // data validation
    '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}', // sparklines
    '{FC87AEE6-9EDD-4A0A-B7FB-166176984837}', // protected ranges
    '{A8765BA9-456A-4dab-B4F3-ACF838C121DE}', // slicers of pivot tables
    '{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}', // slicers of tables
    '{7E03D99C-DC04-49d9-9315-930204A7B6E9}', // timelines
    '{01252117-D84E-4E92-8308-4BE1C098FCBB}', // ignored errors
  ].map((u) => u.toUpperCase()),
)

/** A rectangle as an absolute reference: `$B$3`, `$B$3:$E$10`, `$C:$D`, `$3:$5`. */
export function absText(x: Rect): string {
  const cell = (r: number, c: number) => `$${indexToCol(c)}$${r + 1}`
  if (x.r1 === 0 && x.r2 === MAX_ROWS - 1) return `$${indexToCol(x.c1)}:$${indexToCol(x.c2)}`
  if (x.c1 === 0 && x.c2 === MAX_COLS - 1) return `$${x.r1 + 1}:$${x.r2 + 1}`
  return x.r1 === x.r2 && x.c1 === x.c2 ? cell(x.r1, x.c1) : `${cell(x.r1, x.c1)}:${cell(x.r2, x.c2)}`
}

/** A sheet's carrier: its last cell. */
const lastCell = (sheet: string) => `${quoteSheet(sheet)}!$XFD$1048576`

/** A reference of an `sqref` (space-separated cells and ranges) on a sheet, a union of them; null when one is no range. */
function sqrefText(sheet: string, sqref: string): string | null {
  const parts = sqref.trim().split(/\s+/).filter(Boolean)
  const out: string[] = []
  for (const p of parts) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(p)
    if (!m) return null
    out.push(`${quoteSheet(sheet)}!${p.replace(/\$/g, '').toUpperCase().replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`)
  }
  return out.length ? out.join(',') : null
}

export async function carriersOf(src: Source, book: BookInput): Promise<Carriers> {
  const names = book.sheets.map((s) => s.name)
  const known = new Set(names.map(sheetKey))
  const taken = new Set((book.names ?? []).map((n) => caseKey(n.name)))
  let n = 0
  const nextName = (): string => {
    let name = `${PREFIX}${n++}`
    while (taken.has(caseKey(name))) name = `${PREFIX}${n++}`
    return name
  }
  const out: Carriers = { sheets: [], refs: [], unparsed: names.map(() => []), anywhere: [], byName: [], named: names.map(() => []) }
  for (const [i, sheet] of names.entries()) {
    let name = `${PREFIX}s${i}`
    for (let k = 2; taken.has(caseKey(name)); k++) name = `${PREFIX}s${i}.${k}`
    out.sheets.push({ name, ref: lastCell(sheet), kind: 'place', what: `the sheet '${sheet}'` })
  }
  const place = (sheet: number, x: Rect | string | null, what: string) => {
    const ref = typeof x === 'string' ? x : x && `${quoteSheet(names[sheet])}!${absText(x)}`
    if (ref === null) out.unparsed[sheet].push(what)
    else out.refs.push({ name: nextName(), ref, kind: 'place', what })
  }
  /**
   * A file's formula of references (a chart's series, a name's, a pivot cache's source): its carrier, its references
   * absolute (the engine moves a name's relative reference with no row or column); none when it holds no reference of
   * the book's sheets.
   */
  const reference = (text: string, what: string) => {
    const f = parseFormula(text, { union: true })
    if (!f.ast) return void out.anywhere.push(what)
    if (!refersToBook(f.ast, known)) return
    const absolute = mapAst(f.ast, (x) => (x.k === 'ref' || x.k === 'ref3d' ? { ...x, ref: { ...x.ref, text: absText(x.ref) } } : undefined))
    out.refs.push({ name: nextName(), ref: canonical(absolute), kind: 'ref', what })
  }
  const types = new ContentTypes(src.types)
  for (const [i, s] of src.sheets.entries()) {
    const sheet = names[i]
    const root = rootOf(s.xml)
    for (const c of root ? childrenOf(s.xml, root) : []) {
      const a = () => attrsOf(s.xml, c)
      const text = () => s.xml.slice(c.start, c.end)
      if (UNPARSED.has(c.local)) out.unparsed[i].push(`the ${c.local} of '${sheet}'`)
      else if (c.local === 'autoFilter' || c.local === 'sortState') place(i, sqrefText(sheet, a().ref ?? ''), `the ${c.local === 'autoFilter' ? 'AutoFilter' : 'sort'} of '${sheet}'`)
      else if (c.local === 'rowBreaks' || c.local === 'colBreaks') {
        for (const b of elements(text(), 'brk')) {
          const id = Number(attributes(b.attrs).id)
          if (!(id >= 0)) continue
          // the line before the break and the one after it: either inserted or deleted moves it
          const x: Rect = c.local === 'rowBreaks' ? { r1: Math.max(0, id - 1), r2: Math.min(id, MAX_ROWS - 1), c1: 0, c2: MAX_COLS - 1 } : { r1: 0, r2: MAX_ROWS - 1, c1: Math.max(0, id - 1), c2: Math.min(id, MAX_COLS - 1) }
          place(i, x, `the page breaks of '${sheet}'`)
        }
      } else if (c.local === 'protectedRanges') for (const p of elements(text(), 'protectedRange')) place(i, sqrefText(sheet, attributes(p.attrs).sqref ?? ''), `the protected ranges of '${sheet}'`)
      else if (c.local === 'extLst') {
        for (const e of elements(text(), 'ext')) {
          const uri = (attributes(e.attrs).uri ?? '').toUpperCase()
          if (!EXT_KNOWN.has(uri)) out.unparsed[i].push(`an extension of '${sheet}' (${uri})`)
          else if (uri === '{05C60535-1F16-4FD2-B633-F4F36F0B64E0}') {
            for (const f of elements(e.body ?? '', 'f')) reference(readText(f.body ?? ''), `the sparklines of '${sheet}'`)
            for (const q of elements(e.body ?? '', 'sqref')) place(i, sqrefText(sheet, readText(q.body ?? '')), `the sparklines of '${sheet}'`)
          } else if (uri === '{FC87AEE6-9EDD-4A0A-B7FB-166176984837}') for (const q of elements(e.body ?? '', 'sqref')) place(i, sqrefText(sheet, readText(q.body ?? '')), `the protected ranges of '${sheet}'`)
        }
      }
    }
    for (const r of s.rels) {
      if (r.external) continue
      if (r.kind === 'drawing') {
        const xml = (await src.archive.text(r.target)) ?? ''
        for (const local of ['twoCellAnchor', 'oneCellAnchor'])
          for (const e of elements(xml, local)) {
            const at = (tag: string) => {
              const p = element(e.body ?? '', tag)
              const [col, row] = ['col', 'row'].map((k) => Number(element(p?.body ?? '', k)?.body))
              return p && Number.isInteger(col) && Number.isInteger(row) ? { row, col } : undefined
            }
            const from = at('from')
            const to = local === 'twoCellAnchor' ? at('to') : from
            const fits = (p: { row: number; col: number } | undefined) => p && p.row >= 0 && p.col >= 0 && p.row < MAX_ROWS && p.col < MAX_COLS
            place(i, fits(from) && fits(to) ? { r1: Math.min(from!.row, to!.row), r2: Math.max(from!.row, to!.row), c1: Math.min(from!.col, to!.col), c2: Math.max(from!.col, to!.col) } : null, `a drawing on '${sheet}' (${r.target})`)
          }
      } else if (r.kind === 'pivotTable') {
        const xml = (await src.archive.text(r.target)) ?? ''
        place(i, sqrefText(sheet, attributes(element(xml, 'location')?.attrs ?? '').ref ?? ''), `the pivot table ${r.target} on '${sheet}'`)
      } else if (r.kind === 'threadedComment') {
        const xml = (await src.archive.text(r.target)) ?? ''
        for (const t of elements(xml, 'threadedComment')) place(i, sqrefText(sheet, attributes(t.attrs).ref ?? ''), `the comment threads of '${sheet}'`)
      } else if (r.kind === 'vmlDrawing') {
        // a legacy drawing holds the notes, which the model has, and form controls, which it does not
        const xml = (await src.archive.text(r.target)) ?? ''
        if (elements(xml, 'ClientData').some((d) => attributes(d.attrs).ObjectType !== 'Note')) out.unparsed[i].push(`the form controls of '${sheet}'`)
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
    if (/^_xlfn\./i.test(name)) continue
    const decl = a.localSheetId === undefined ? undefined : src.decls[Number(a.localSheetId)]
    // a worksheet's name as the book has it (the reader repairs names Excel would not take)
    const scope = a.localSheetId === undefined ? '' : decl?.sheet !== undefined ? names[decl.sheet] : decl?.name
    if (held.has(`${scope === undefined ? '' : sheetKey(scope)}\n${caseKey(name)}`)) continue
    reference(readText(src.wb.slice(d.open, d.end - d.qname.length - 3)), `the name ${name}`)
  }
  for (const part of types.parts((t) => /drawingml\.chart\+xml$|chartex\+xml$/i.test(t))) {
    const xml = (await src.archive.text(part)) ?? ''
    for (const f of elements(xml, 'f')) reference(readText(f.body ?? ''), `the chart ${part}`)
    // a pivot chart's table, by its sheet's name
    const pivot = element(element(xml, 'pivotSource')?.body ?? '', 'name')
    const sheet = pivot && /^'?\[[^\]]*\]((?:[^'!]|'')+)'?!/.exec(readText(pivot.body ?? ''))?.[1]?.replace(/''/g, "'")
    if (sheet) {
      const k = names.findIndex((x) => sheetKey(x) === sheetKey(sheet))
      if (k >= 0) out.named[k].push(`the pivot chart ${part}`)
    }
  }
  for (const part of types.parts((t) => /pivotCacheDefinition\+xml$/i.test(t))) {
    const xml = (await src.archive.text(part)) ?? ''
    const what = `the source of the pivot cache ${part}`
    const ws = element(element(xml, 'cacheSource')?.body ?? '', 'worksheetSource')
    const a = attributes(ws?.attrs ?? '')
    if (ws && attr(a, 'id') === undefined) {
      if (a.name !== undefined) out.byName.push({ name: readText(a.name), what })
      else if (a.sheet !== undefined && a.ref !== undefined) reference(`${quoteSheet(readText(a.sheet))}!${a.ref}`, what)
    }
    for (const rs of elements(xml, 'rangeSet')) {
      const r = attributes(rs.attrs)
      if (attr(r, 'id') === undefined && r.sheet !== undefined && r.ref !== undefined) reference(`${quoteSheet(readText(r.sheet))}!${r.ref}`, what)
    }
  }
  // macro and dialog sheets hold formulas over the worksheets
  for (const d of src.decls) if (d.kind !== undefined && d.kind !== 'worksheet' && d.kind !== 'chartsheet') out.anywhere.push(`the ${d.kind} '${d.name}'`)
  return out
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
  return [...c.sheets, ...c.refs].map((x) => ({ name: x.name, ref: x.ref, hidden: true }))
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

/** Whether a carrier's reference lost its sheet, as a deleted sheet's references do (`#REF!$B$2`): each of its areas. */
export function lostSheet(text: string): boolean {
  let ast: Ast
  try {
    ast = parse(text, { union: true })
  } catch (e) {
    if (e instanceof FormulaError) return false
    throw e
  }
  const items = ast.k === 'union' ? ast.items : [ast]
  return items.every((x) => x.k === 'err' && x.sheet === undefined && x.ref !== undefined)
}

/** The sheet a sheet's carrier names now, and whether its cell is still the last one; undefined when the sheet is gone. */
export function sheetOfCarrier(text: string): { sheet: string; moved: boolean } | undefined {
  let ast: Ast
  try {
    ast = parse(text)
  } catch (e) {
    if (e instanceof FormulaError) return undefined
    throw e
  }
  if (ast.k === 'ref' && ast.sheet !== null) return { sheet: ast.sheet, moved: ast.ref.r1 !== MAX_ROWS - 1 || ast.ref.c1 !== MAX_COLS - 1 }
  if (ast.k === 'err' && ast.sheet !== undefined && ast.book === undefined) return { sheet: ast.sheet, moved: true }
  return undefined
}

