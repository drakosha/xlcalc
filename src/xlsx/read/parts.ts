/**
 * The parts of a sheet that are no cells' values: frozen panes (`<sheetView><pane state="frozen">`), notes (the comments
 * part a sheet links to, with its VML drawing), threads of comments (Excel 365's threaded comments, read as notes) and
 * hyperlinks (`<hyperlinks>`, their addresses in the sheet's relationships). What the model has no place for is reported.
 */
import type { CellInput, CellLink, CellNote, FrozenPanes } from '../../contract.js'
import { formatAddress, MAX_COLS, MAX_ROWS, parseAddress } from '../../model.js'
import { parseRange } from '../../structure/area.js'
import type { Issues } from '../issues.js'
import { attr, attributes, element, elements } from '../xml/scan.js'
import { readText } from '../xml/text.js'
import type { Rel } from './package.js'
import { stringItem } from './strings.js'

/** A sheet's parts as the archive holds them: its comments, VML drawing and threaded comments, the book's persons. */
export interface PartsSource {
  sheet: string
  head: string
  tail: string
  rels: Rel[]
  comments?: string
  vml?: string
  threads?: string
  persons?: string
}

/**
 * Frozen panes of the sheet's first view: the rows and columns before the split, counted from the first ones. A pane
 * frozen over a sheet scrolled down or right showed fewer of them; split panes without freezing are reported.
 */
export function readFrozen(head: string, sheet: string, issues: Issues): FrozenPanes | undefined {
  const view = element(element(head, 'sheetViews')?.body ?? '', 'sheetView')
  const pane = view?.body ? element(view.body, 'pane') : undefined
  if (!view || !pane) return undefined
  const a = attributes(pane.attrs)
  // a pane's state is a split unless it says otherwise
  if (a.state !== 'frozen' && a.state !== 'frozenSplit') {
    issues.add('frozenPanes', 'split panes that are not frozen are not in the model', { sheet })
    return undefined
  }
  const top = parseAddress(/^[A-Z]{1,3}[1-9]\d*$/.test(attributes(view.attrs).topLeftCell ?? '') ? attributes(view.attrs).topLeftCell! : 'A1')
  const [ys, xs] = [Math.trunc(Number(a.ySplit ?? 0)), Math.trunc(Number(a.xSplit ?? 0))]
  const out: FrozenPanes = {}
  if (ys > 0) out.rows = Math.min(top.row + ys, MAX_ROWS - 1)
  if (xs > 0) out.columns = Math.min(top.col + xs, MAX_COLS - 1)
  if ((ys > 0 && top.row > 0) || (xs > 0 && top.col > 0)) issues.add('frozenPanes', 'panes frozen over a sheet scrolled down or right are read from the first row and column', { sheet })
  return out.rows || out.columns ? out : undefined
}

/** The cells of a reference (`A1`, `A1:B3`), at most `max` of them listed; null when it is none. */
function refCells(ref: string, max: number): string[] | null {
  const r = parseRange(ref.toUpperCase())
  if (!r) return null
  if ((r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1) > max) return []
  const out: string[] = []
  for (let row = r.r1; row <= r.r2; row++) for (let c = r.c1; c <= r.c2; c++) out.push(formatAddress(row, c))
  return out
}

/** The person names of threaded comments by id (`persons/person.xml`). */
function persons(xml: string | undefined): Map<string, string> {
  return new Map(elements(xml ?? '', 'person').map((e) => {
    const a = attributes(e.attrs)
    return [a.id ?? '', readText(a.displayName ?? '')]
  }))
}

/**
 * Threads of comments by cell, each as a note: its first comment's author, its comments' text one after another, each
 * reply after an empty line with its author (`Ann: text`).
 */
function threads(src: PartsSource, issues: Issues): Map<string, CellNote> {
  const out = new Map<string, CellNote>()
  if (!src.threads) return out
  const who = persons(src.persons)
  const byId = new Map<string, string>()
  for (const e of elements(src.threads, 'threadedComment')) {
    const a = attributes(e.attrs)
    const text = readText(element(e.body ?? '', 'text')?.body ?? '')
    const author = who.get(a.personId ?? '')
    const parent = a.parentId === undefined ? undefined : byId.get(a.parentId)
    const cell = parent ?? a.ref?.toUpperCase()
    if (!cell || !refCells(cell, 1)?.length) continue
    if (a.id) byId.set(a.id, cell)
    const where = { sheet: src.sheet, cell }
    const note = out.get(cell)
    if (!note) {
      out.set(cell, author === undefined ? { text } : { author, text })
      issues.add('comment', "a thread of comments is read as a note: its first comment's author, every comment's text with the replies' authors; dates, mentions and the thread's state are not kept", where)
    } else note.text += `\n\n${author ?? ''}: ${text}`
  }
  return out
}

/**
 * Notes by cell from the comments part; formatted text is read as its plain text, a note always shown as one shown on
 * hover. The cells of `threads` hold Excel's placeholder notes for older readers: they are left to their threads.
 */
function notes(src: PartsSource, threads: Map<string, CellNote>, issues: Issues): Map<string, CellNote> {
  const out = new Map<string, CellNote>()
  if (!src.comments) return out
  const authors = elements(element(src.comments, 'authors')?.body ?? '', 'author').map((e) => readText(e.body ?? ''))
  for (const e of elements(element(src.comments, 'commentList')?.body ?? '', 'comment')) {
    const a = attributes(e.attrs)
    const cell = a.ref?.toUpperCase()
    if (!cell || !refCells(cell, 1)?.length || threads.has(cell)) continue
    const body = element(e.body ?? '', 'text')?.body ?? ''
    const { text } = stringItem(body)
    // runs of more than one look, or bold, italic, underlined or struck ones (the bold author's name of a note made in Excel's window)
    const looks = new Set(elements(body, 'rPr').map((r) => r.body ?? ''))
    if (looks.size > 1 || [...looks].some((l) => /<(?:[\w.-]+:)?(?:b|i|u|strike)[\s/>]/.test(l))) issues.add('comment', "a note's formatted text is read as its plain text", { sheet: src.sheet, cell })
    const author = authors[Number(a.authorId)]
    out.set(cell, author ? { author, text } : { text })
  }
  // the VML drawing shows the notes; one shown all the time is read as one shown on hover
  for (const shape of elements(src.vml ?? '', 'shape')) {
    const data = element(shape.body ?? '', 'ClientData')
    if (!data) continue
    const kind = attributes(data.attrs).ObjectType
    if (kind !== 'Note') {
      issues.add('drawing', 'pictures, charts and shapes are not in the model', { sheet: src.sheet })
      continue
    }
    if (!/<(?:[\w.-]+:)?Visible[\s/>]/.test(data.body ?? '')) continue
    const [row, col] = ['Row', 'Column'].map((k) => Number(element(data.body ?? '', k)?.body))
    const cell = Number.isInteger(row) && Number.isInteger(col) && row >= 0 && col >= 0 ? formatAddress(row, col) : undefined
    if (cell && !threads.has(cell)) issues.add('comment', 'a note shown all the time is read as one shown when the pointer rests on its cell', { sheet: src.sheet, cell })
  }
  return out
}

/** Hyperlinks over up to this many cells are given to each; a larger range to the cells it holds that the sheet has. */
const LINK_CELLS = 4096

/** Hyperlinks by cell: the address from the sheet's relationships, the place in the book, the tooltip. */
function links(src: PartsSource, cells: Record<string, CellInput>, issues: Issues): Map<string, CellLink> {
  const out = new Map<string, CellLink>()
  const rels = new Map(src.rels.map((r) => [r.id, r]))
  for (const e of elements(element(src.tail, 'hyperlinks')?.body ?? '', 'hyperlink')) {
    const a = attributes(e.attrs)
    const link: CellLink = {}
    const id = attr(a, 'id')
    const rel = id === undefined ? undefined : rels.get(id)
    if (rel?.external && rel.target) link.url = rel.target
    if (a.location) link.location = readText(a.location)
    if (a.tooltip) link.tooltip = readText(a.tooltip)
    const where = { sheet: src.sheet, cell: a.ref ?? '' }
    if (id !== undefined && link.url === undefined) issues.add('hyperlink', "a hyperlink's address that is no external relationship of the sheet is left out", where)
    if (link.url === undefined && link.location === undefined) continue
    let list = refCells(a.ref ?? '', LINK_CELLS)
    if (list && !list.length) {
      const r = parseRange(a.ref!.toUpperCase())!
      list = Object.keys(cells).filter((c) => {
        const p = parseAddress(c)
        return p.row >= r.r1 && p.row <= r.r2 && p.col >= r.c1 && p.col <= r.c2
      })
      issues.add('hyperlink', 'a hyperlink over more than 4,096 cells is read on the cells of its range that hold something', where)
    }
    for (const cell of list ?? []) if (!out.has(cell)) out.set(cell, link)
  }
  return out
}

/** Notes, threads of comments and hyperlinks put on the sheet's cells (a cell holding only them is made). */
export function readCellParts(src: PartsSource, cells: Record<string, CellInput>, issues: Issues): void {
  const thread = threads(src, issues)
  for (const [cell, note] of [...notes(src, thread, issues), ...thread]) cells[cell] = { ...cells[cell], note }
  for (const [cell, link] of links(src, cells, issues)) cells[cell] = { ...cells[cell], link }
}
