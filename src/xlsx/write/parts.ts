/**
 * The parts of a sheet that are no cells' values, as Excel writes them: frozen panes in the sheet's view, hyperlinks
 * (`<hyperlinks>`, their addresses as external relationships of the sheet), notes (a comments part, and a VML drawing
 * with a hidden box for each, which Excel needs to show them).
 */
import type { CellLink, CellNote, FrozenPanes } from '../../contract.js'
import { formatAddress, parseAddress } from '../../model.js'
import type { Issues } from '../issues.js'
import { escapeAttr, escapeText, needsPreserve } from '../xml/text.js'

const MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

/** Excel's limit on a note's text, as on a cell's. */
const MAX_NOTE = 32_767

/** The content of a sheet's view with frozen panes: the pane after the frozen rows and columns, active, and its selection. */
export function paneXml(f: FrozenPanes | undefined): string {
  if (!f) return ''
  const [rows, cols] = [f.rows ?? 0, f.columns ?? 0]
  const pane = rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight'
  const at = formatAddress(rows, cols)
  return `<pane${cols ? ` xSplit="${cols}"` : ''}${rows ? ` ySplit="${rows}"` : ''} topLeftCell="${at}" activePane="${pane}" state="frozen"/><selection pane="${pane}" activeCell="${at}" sqref="${at}"/>`
}

/** Cells in the order of a sheet (row by row). */
const inOrder = <T>(m: Map<string, T>): [string, T][] => [...m].sort((a, b) => {
  const [p, q] = [parseAddress(a[0]), parseAddress(b[0])]
  return p.row - q.row || p.col - q.col
})

/** Excel's limit on a hyperlink's tooltip (ScreenTip): it does not load a file with a longer one (the parts set `files`). */
const MAX_TIP = 255

/** `<hyperlinks>` of the cells' links; `rel` gives the relationship id of an address outside the book. */
export function hyperlinksXml(links: Map<string, CellLink>, rel: (url: string) => string, sheet: string, issues: Issues): string {
  if (!links.size) return ''
  const items = inOrder(links).map(([cell, l]) => {
    let tip = l.tooltip
    if (tip !== undefined && tip.length > MAX_TIP) {
      tip = tip.slice(0, MAX_TIP)
      issues.add('textLength', "a hyperlink's tooltip over Excel's 255 characters is cut", { sheet, cell })
    }
    const attrs = (l.url !== undefined ? ` r:id="${rel(l.url)}"` : '') + (l.location !== undefined ? ` location="${escapeAttr(l.location)}"` : '') + (tip !== undefined ? ` tooltip="${escapeAttr(tip)}"` : '')
    return `<hyperlink ref="${cell}"${attrs}/>`
  })
  return `<hyperlinks>${items.join('')}</hyperlinks>`
}

/** The comments part of a sheet's notes: authors once each (an empty one for notes without), then the notes in cell order. */
export function commentsXml(notes: Map<string, CellNote>, sheet: string, issues: Issues): string {
  const authors = new Map<string, number>()
  const list = inOrder(notes).map(([cell, n]) => {
    const author = n.author ?? ''
    if (!authors.has(author)) authors.set(author, authors.size)
    let text = n.text
    if (text.length > MAX_NOTE) {
      text = text.slice(0, MAX_NOTE)
      issues.add('textLength', "a note over Excel's 32,767 characters is cut", { sheet, cell })
    }
    return `<comment ref="${cell}" authorId="${authors.get(author)}"><text><t${needsPreserve(text) ? ' xml:space="preserve"' : ''}>${escapeText(text)}</t></text></comment>`
  })
  const names = [...authors.keys()].map((a) => (a ? `<author>${escapeText(a)}</author>` : '<author/>'))
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="${MAIN}"><authors>${names.join('')}</authors><commentList>${list.join('')}</commentList></comments>`
}

/**
 * The VML drawing of a sheet's notes: a hidden box beside each cell. Shape ids are unique through the book: the sheet's
 * blocks of 1,024 ids start at `block` (the idmap lists them), the notes count from one past its first id.
 */
export function vmlXml(notes: Map<string, CellNote>, block: number): string {
  const cells = inOrder(notes).map(([cell]) => parseAddress(cell))
  const last = Math.floor((block * 1024 + cells.length) / 1024)
  const blocks = Array.from({ length: last - block + 1 }, (_, i) => block + i).join(',')
  const shapes = cells.map(({ row, col }, i) => {
    // the box Excel draws for a new note: from the next column on, from the row above
    const top = Math.max(row - 1, 0)
    const anchor = `${col + 1}, 15, ${top}, ${row ? 10 : 2}, ${col + 3}, 15, ${top + 4}, 16`
    return (
      `<v:shape id="_x0000_s${block * 1024 + i + 1}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">` +
      '<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
      `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${row}</x:Row><x:Column>${col}</x:Column></x:ClientData></v:shape>`
    )
  })
  return (
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
    `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${blocks}"/></o:shapelayout>` +
    '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>' +
    `${shapes.join('')}</xml>`
  )
}

/** The blocks of shape ids a sheet's `count` notes take from `block` on: the next sheet's first block. */
export const nextBlock = (block: number, count: number): number => Math.floor((block * 1024 + count) / 1024) + 1
