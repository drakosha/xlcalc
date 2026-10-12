/**
 * The parts of a sheet that are no cells' values: frozen panes (`<sheetView><pane state="frozen">`), notes (the comments
 * part a sheet links to, with its VML drawing), threads of comments (Excel 365's threaded comments, read as notes) and
 * hyperlinks (`<hyperlinks>`, their addresses in the sheet's relationships). What the model has no place for is reported.
 */
import type { CellInput, CellLink, CellNote, FrozenPanes } from '../../contract.js'
import { formatAddress, MAX_COLS, MAX_ROWS, parseAddress } from '../../model.js'
import { cellCount, cellsWithin, parseRange, type Rect } from '../../structure/area.js'
import type { Issues } from '../issues.js'
import { attr, attributes, element, elements, has } from '../xml/scan.js'
import { readText } from '../xml/text.js'
import type { Rel } from './package.js'
import { fileCell } from './refs.js'
import { stringItem } from './strings.js'

/** A sheet's parts as the archive holds them: its comments, VML drawing and threaded comments, the book's persons. */
export interface PartsSource {
  sheet: string
  /** The cells of a look of the sheet's looks in a rectangle, which a large hyperlink range goes to as to the sheet's cells. */
  looked?: (x: Rect) => Iterable<string>
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
    issues.add('frozenPanes', { key: 'splitPanes' }, { sheet })
    return undefined
  }
  const top = parseAddress(/^[A-Z]{1,3}[1-9]\d*$/.test(attributes(view.attrs).topLeftCell ?? '') ? attributes(view.attrs).topLeftCell! : 'A1')
  const [ys, xs] = [Math.trunc(Number(a.ySplit ?? 0)), Math.trunc(Number(a.xSplit ?? 0))]
  const out: FrozenPanes = {}
  if (ys > 0) out.rows = Math.min(top.row + ys, MAX_ROWS - 1)
  if (xs > 0) out.columns = Math.min(top.col + xs, MAX_COLS - 1)
  if ((ys > 0 && top.row > 0) || (xs > 0 && top.col > 0)) issues.add('frozenPanes', { key: 'frozenScrolled' }, { sheet })
  return out.rows || out.columns ? out : undefined
}

/** The person names of threaded comments by id (`persons/person.xml`). */
function persons(xml: string | undefined): Map<string, string> {
  return new Map(elements(xml ?? '', 'person').map((e) => {
    const a = attributes(e.attrs)
    return [a.id ?? '', readText(a.displayName ?? '')]
  }))
}

/** A note of the comments part: its cell, author and text, and whether its runs have more than one look. */
interface Comment {
  cell: string
  author: string
  text: string
  formatted: boolean
}

/** The cell of a note or a thread (`ref`), as Excel reads it; one that is no one cell is reported, as Excel does not open such a file (the set `parts/refs-held`). */
function commentCell(ref: string | undefined, sheet: string, issues: Issues): string | null {
  const cell = ref === undefined ? null : fileCell(ref)
  if (!cell) issues.add('comment', { key: 'noteCell' }, { sheet, cell: ref ?? '' })
  return cell
}

function comments(xml: string | undefined, sheet: string, issues: Issues): Comment[] {
  if (!xml) return []
  const authors = elements(element(xml, 'authors')?.body ?? '', 'author').map((e) => readText(e.body ?? ''))
  return elements(element(xml, 'commentList')?.body ?? '', 'comment').flatMap((e) => {
    const a = attributes(e.attrs)
    const cell = commentCell(a.ref, sheet, issues)
    if (!cell) return []
    const body = element(e.body ?? '', 'text')?.body ?? ''
    // runs of more than one look (the bold author's name of a note made in Excel's window); one look is the note's font
    return [{ cell, author: authors[Number(a.authorId)] ?? '', text: stringItem(body).text, formatted: new Set(elements(body, 'rPr').map((r) => r.body ?? '')).size > 1 }]
  })
}

/** The thread id of Excel's placeholder note for a thread, from its author `tc={id}`. */
const placeholderOf = (c: Comment) => (c.author.startsWith('tc=') ? c.author.slice(3) : undefined)

/**
 * Threads of comments by cell, each as a note: its first comment's author, its comments' text one after another, each
 * reply after an empty line with its author (`Ann: text`). A thread is on the cell of its placeholder note where there is
 * one: an Excel without threads moves the placeholder with its cell and leaves the thread's own cell as it was.
 */
function threads(src: PartsSource, placed: Map<string, string>, issues: Issues): Map<string, CellNote> {
  const out = new Map<string, CellNote>()
  if (!src.threads) return out
  const who = persons(src.persons)
  const byId = new Map<string, string>()
  for (const e of elements(src.threads, 'threadedComment')) {
    const a = attributes(e.attrs)
    const text = readText(element(e.body ?? '', 'text')?.body ?? '')
    const author = who.get(a.personId ?? '')
    const cell = a.parentId !== undefined ? byId.get(a.parentId) : (placed.get(a.id ?? '') ?? commentCell(a.ref, src.sheet, issues))
    if (!cell) continue
    if (a.id) byId.set(a.id, cell)
    const note = out.get(cell)
    if (!note) {
      out.set(cell, author === undefined ? { text } : { author, text })
      issues.add('comment', { key: 'noteThread' }, { sheet: src.sheet, cell })
    } else note.text += `\n\n${author ?? ''}: ${text}`
  }
  return out
}

/**
 * Notes by cell from the comments part; formatted text is read as its plain text, a note always shown as one shown on
 * hover. Excel's placeholder notes of the threads, which the threads take the place of, are left out.
 */
function notes(src: PartsSource, list: Comment[], threads: Map<string, CellNote>, issues: Issues): Map<string, CellNote> {
  const out = new Map<string, CellNote>()
  for (const c of list) {
    if (threads.has(c.cell) || placeholderOf(c) !== undefined) continue
    if (c.formatted) issues.add('comment', { key: 'noteFormatted' }, { sheet: src.sheet, cell: c.cell })
    out.set(c.cell, c.author ? { author: c.author, text: c.text } : { text: c.text })
  }
  // the VML drawing shows the notes; one shown all the time is read as one shown on hover
  for (const shape of elements(src.vml ?? '', 'shape')) {
    const data = element(shape.body ?? '', 'ClientData')
    if (!data) continue
    const kind = attributes(data.attrs).ObjectType
    if (kind !== 'Note') {
      issues.add('drawing', { key: 'drawings' }, { sheet: src.sheet })
      continue
    }
    if (!has(data.body ?? '', 'Visible')) continue
    const [row, col] = ['Row', 'Column'].map((k) => Number(element(data.body ?? '', k)?.body))
    const cell = Number.isInteger(row) && Number.isInteger(col) && row >= 0 && col >= 0 ? formatAddress(row, col) : undefined
    if (cell && !threads.has(cell)) issues.add('comment', { key: 'noteShown' }, { sheet: src.sheet, cell })
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
    if (id !== undefined && link.url === undefined) issues.add('hyperlink', { key: 'linkAddress' }, where)
    if (link.url === undefined && link.location === undefined) continue
    const r = parseRange((a.ref ?? '').toUpperCase())
    if (!r) continue
    if (cellCount(r) > LINK_CELLS) issues.add('hyperlink', { key: 'linkCells' }, where)
    for (const cell of cellsWithin(r, () => [...Object.keys(cells), ...(src.looked?.(r) ?? [])], LINK_CELLS)) if (!out.has(cell)) out.set(cell, link)
  }
  return out
}

/** Notes, threads of comments and hyperlinks put on the sheet's cells (a cell holding only them is made). */
export function readCellParts(src: PartsSource, cells: Record<string, CellInput>, issues: Issues): void {
  const list = comments(src.comments, src.sheet, issues)
  const placed = new Map(list.flatMap((c) => (placeholderOf(c) ? [[placeholderOf(c)!, c.cell] as const] : [])))
  const thread = threads(src, placed, issues)
  for (const [cell, note] of [...notes(src, list, thread, issues), ...thread]) cells[cell] = { ...cells[cell], note }
  for (const [cell, link] of links(src, cells, issues)) cells[cell] = { ...cells[cell], link }
}
