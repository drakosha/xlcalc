/**
 * What became of the file's parts: which sheet of the book each sheet of the file became (by the sheets' carriers),
 * and for each part with carriers that moved, the edits that write them back (`carriers.ts`). `writeKept` refuses
 * what it cannot write as Excel would: a part whose item refuses, a change of a sheet with parts whose references no
 * carrier reads, a sheet a pivot chart names deleted, a table or a name a pivot cache takes its data from deleted, a
 * sheet deleted with a pivot table or another part of its own the model does not have, the links to other workbooks
 * changed. The file would point at the wrong cells: no file is written then.
 */
import type { BookInput } from '../../contract.js'
import { caseKey } from '../../collate.js'
import { quoteSheet, sheetKey } from '../../syntax/print.js'
import { plain } from '../../syntax/links.js'
import { escapeAttr, escapeText } from '../xml/text.js'
import { columnPoints, DEFAULT_ROW_HEIGHT } from '../../column-width.js'
import { formatAddress } from '../../model.js'
import { entryAt, lines } from '../../layout/spans.js'
import { attributes } from '../xml/scan.js'
import { compared, sheetOfCarrier, type Carriers, type Edit, type Geometry } from './carriers.js'
import type { FormatArea } from './formats.js'
import { areasOf } from './refs.js'
import type { Source } from './source.js'

/**
 * Why `writeKept` refuses, for a program (`KeptRefusal.code`): `file`, a file it does not write over (a Strict Open XML
 * workbook, a sheet name Excel does not take, a worksheet in other namespaces); `carriers`, a hidden name `openKept`
 * gave the workbook is gone; `unread`, a sheet whose rows, columns or name changed has a part with references writeKept
 * does not read; `deleted`, cells, a sheet or a table a part of the file refers to or belongs to were deleted; `moved`, a
 * part would move or be cut by the change, which writeKept does not write; `links`, a link to another workbook was
 * removed, moved or pointed at another file, or one of a kind other than a workbook (DDE, OLE) changed; `font`, the default font changed; `threads`, the notes of a sheet with comment threads changed; `rules`,
 * the data validation of a sheet changed and the file has rules of it the workbook could not read (its `readXlsx` issues
 * `dataValidation`), which writing them again would lose, or its conditional formats changed and the file has rules of
 * it that cannot go with them as Excel would (a pivot table's, one read over part of its ranges, one whose formulas
 * would need making for a new first cell).
 */
export type KeptRefusalCode = 'file' | 'carriers' | 'unread' | 'deleted' | 'moved' | 'links' | 'font' | 'threads' | 'rules'

/** `writeKept`'s refusal: the change would leave a part of the file the model does not hold pointing at the wrong cells; `code` says why. */
export class KeptRefusal extends Error {
  override readonly name = 'KeptRefusal'
  constructor(
    why: string,
    readonly code: KeptRefusalCode,
  ) {
    super(`xlsx/keep: ${why}`)
  }
}

export interface Identity {
  /** By sheet of the book written: the file's sheet it is (an index of the book read), none for a new one. */
  source: (number | undefined)[]
  /** By sheet of the book read: the sheet of the book written it became; none for one deleted. */
  target: Map<number, number>
  /** By sheet of the book read: rows or columns were inserted or deleted on it. */
  lines: boolean[]
  /** The edits that write the moved references back to the parts. */
  edits: Edit[]
  /** By sheet of the book read: its cells with what only the file holds that moved, by address now on its sheet of the book written (the address they had), and the addresses they left. */
  cells: Map<number, Moved>
  /** By sheet of the book read: the rectangles of its cell formats that look as earlier ones where they lie now on its sheet of the book written; none for a sheet past `KEPT_FORMATS` or whose carriers a host dropped. */
  formats: Map<number, FormatArea[]>
}

export interface Moved {
  to: Map<string, string>
  left: Set<string>
}

/** The relationships a sheet deleted may have: parts of its own the model has or that go with it, or that hold no reference. */
const DELETABLE = new Set(['table', 'comments', 'vmlDrawing', 'printerSettings', 'hyperlink', 'threadedComment', 'drawing'])

const OPEN_KEPT = 'open the file with openKept and keep the hidden names it adds'

/** EMUs in a point. */
const EMU = 12_700

export function refusals(src: Source, book0: BookInput, book1: BookInput, c: Carriers): Identity {
  const names = new Map((book1.names ?? []).filter((n) => n.scope === undefined).map((n) => [caseKey(n.name), n]))
  const index = new Map(book1.sheets.map((s, j) => [sheetKey(s.name), j]))
  const out: Identity = { source: book1.sheets.map(() => undefined), target: new Map(), lines: [], edits: [], cells: new Map(), formats: new Map() }
  for (const [k, carrier] of c.sheets.entries()) {
    const was = book0.sheets[k].name
    const n = names.get(caseKey(carrier.name))
    const at = n && sheetOfCarrier(n.ref)
    const j = n ? at && index.get(sheetKey(at.sheet)) : index.get(sheetKey(was))
    if (j !== undefined && out.source[j] === undefined) {
      out.source[j] = k
      out.target.set(k, j)
    }
    out.lines[k] = !!at?.moved
    const renamed = j === undefined || book1.sheets[j].name !== was
    const changed = renamed || out.lines[k]
    const unread = [...c.unparsed[k], ...c.anywhere]
    if (!n && (unread.length || c.named[k].length)) throw new KeptRefusal(`the workbook has no name ${carrier.name}, which tells what became of the sheet '${was}': ${OPEN_KEPT}`, 'carriers')
    if (changed && unread.length) throw new KeptRefusal(`the sheet '${was}' had rows or columns inserted or deleted, or was renamed or deleted, and ${unread[0]} holds references writeKept does not read`, 'unread')
    for (const p of renamed ? c.named[k] : []) {
      if (j === undefined) throw new KeptRefusal(`the sheet '${was}' was deleted, and ${p.what} names it`, 'deleted')
      // a pivot chart's table by its sheet's new name, quoted as Excel quotes it
      const name = book1.sheets[j].name
      out.edits.push({ part: p.part, start: p.at[0], end: p.at[1], text: escapeText(`${plain(name) ? p.prefix + name : quoteSheet(p.prefix + name)}!${p.rest}`) })
    }
    if (j === undefined) {
      const own = src.sheets[k].rels.find((r) => !DELETABLE.has(r.kind))
      if (own) throw new KeptRefusal(`the sheet '${was}' was deleted, and it has a part the workbook does not hold (${own.kind} ${own.target})`, 'deleted')
    }
  }
  // the cells with what only the file holds: where each went on its sheet (one moved to another sheet, or deleted, leaves it)
  for (const [k, list] of c.cells.entries()) {
    const j = out.target.get(k)
    if (j === undefined || !list.length) continue
    const moved: Moved = { to: new Map(), left: new Set() }
    for (const cell of list) {
      const n = names.get(caseKey(cell.carrier.name))
      if (!n) throw new KeptRefusal(`the workbook has no name ${cell.carrier.name}, which tells where the cell ${cell.addr} of '${book0.sheets[k].name}' went: ${OPEN_KEPT}`, 'carriers')
      const area = areasOf(n.ref)?.[0]
      const at = area && 'rect' in area && sheetKey(area.sheet) === sheetKey(book1.sheets[j].name) ? formatAddress(area.rect.r1, area.rect.c1) : undefined
      if (at === cell.addr) continue
      moved.left.add(cell.addr)
      if (at !== undefined) moved.to.set(at, cell.addr)
    }
    if (moved.left.size) out.cells.set(k, moved)
  }
  // the rectangles of cell formats on their sheet now; a carrier gone leaves the sheet's cells their formats by address
  for (const [k, list] of c.formats.entries()) {
    const j = out.target.get(k)
    if (j === undefined || !list) continue
    const areas: FormatArea[] = []
    const kept = list.every(({ kind, s, carrier }) => {
      const n = names.get(caseKey(carrier.name))
      for (const a of (n && areasOf(n.ref)) || []) if ('rect' in a && sheetKey(a.sheet) === sheetKey(book1.sheets[j].name)) areas.push({ kind, s, rect: a.rect })
      return n !== undefined
    })
    if (kept) out.formats.set(k, areas)
  }
  /** A sheet's lines in EMUs as the book written has them: a row's height (its sheet's default where it has none), a column's width; hidden ones none. */
  const geometry = (k: number | undefined): Geometry => {
    const s = k === undefined ? undefined : book1.sheets[out.target.get(k)!]
    const rows = lines(s?.rows, 'r')
    const cols = lines(s?.columns, 'c')
    const fp = k === undefined ? {} : attributes(/<(?:[\w.-]+:)?sheetFormatPr\b([^>]*)>/.exec(src.sheets[k].xml)?.[1] ?? '')
    const height = Number(fp.defaultRowHeight) || DEFAULT_ROW_HEIGHT
    return {
      row: (r) => {
        const e = entryAt(rows, r)
        return e?.hidden ? 0 : (e?.height ?? height) * EMU
      },
      col: (c) => {
        const e = entryAt(cols, c)
        return e?.hidden ? 0 : columnPoints(e?.width) * EMU
      },
    }
  }
  for (const item of c.items) {
    // a part on a sheet deleted goes with it
    if (item.on !== undefined && !out.target.has(item.on)) continue
    const now = item.carriers.map((x) => {
      const n = names.get(caseKey(x.name))
      if (!n) throw new KeptRefusal(`the workbook has no name ${x.name}, which holds a reference of ${item.what}: ${OPEN_KEPT}`, 'carriers')
      return n.ref
    })
    if (item.carriers.every((x, k) => compared(now[k], x.kind) === compared(x.ref, x.kind))) continue
    const fix = item.fix(now, geometry(item.on))
    if ('refuse' in fix) throw new KeptRefusal(fix.refuse, fix.code)
    out.edits.push(...fix.edits)
  }
  // a pivot cache's source by the name of a table: a table renamed (the same range of the same sheet) takes the new name
  const tables = new Set(book1.sheets.flatMap((s) => (s.tables ?? []).map((t) => caseKey(t.name))))
  for (const b of c.byName) {
    if (tables.has(caseKey(b.name)) || (book1.names ?? []).some((n) => caseKey(n.name) === caseKey(b.name))) continue
    const k = book0.sheets.findIndex((s) => s.tables?.some((t) => caseKey(t.name) === caseKey(b.name)))
    const was = k < 0 ? undefined : book0.sheets[k].tables!.find((t) => caseKey(t.name) === caseKey(b.name))!
    const j = k < 0 ? undefined : out.target.get(k)
    const now = was && j !== undefined ? book1.sheets[j].tables?.find((t) => t.range === was.range && !book0.sheets[k].tables!.some((x) => caseKey(x.name) === caseKey(t.name))) : undefined
    if (!now) throw new KeptRefusal(`${b.what} is the table or name ${b.name}, which the workbook no longer has`, 'deleted')
    out.edits.push({ part: b.part, start: b.at[0], end: b.at[1], text: escapeAttr(now.name) })
  }
  // formulas read a link by its number: the file's links keep theirs, a new one comes after them
  const [l0, l1] = [book0.externalBooks ?? [], book1.externalBooks ?? []]
  if (l1.length < l0.length || l0.some((b, i) => b.target !== l1[i].target))
    throw new KeptRefusal("a link to another workbook was removed, moved or pointed at another file, which the file's formulas would read by another number", 'links')
  return out
}

/** A part's text with edits made (their spans in the text as the file has it, none crossing another). */
export function edited(text: string, edits: Edit[]): string {
  const list = [...edits].sort((a, b) => a.start - b.start || a.end - b.end)
  let out = ''
  let at = 0
  for (const e of list) {
    if (e.start < at) throw new Error(`xlsx/keep: edits of ${e.part} cross at ${e.start}`)
    out += text.slice(at, e.start) + e.text
    at = e.end
  }
  return out + text.slice(at)
}
