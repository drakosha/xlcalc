/** End notes flow through pages independently of the worksheet grid. */
import { parseCell } from '../../a1.js'
import { DEFAULT_ROW_HEIGHT } from '../../column-width.js'
import { checkNumber } from '../../input/basic.js'
import { InputError } from '../../input-error.js'
import { formatAddress } from '../../model.js'
import { measureFont } from '../../numfmt/font-metrics.js'
import { advance } from '../../numfmt/glyphs.js'
import type { PrintNoteLine, PrintNoteMetrics, PrintNotes, PrintRect } from './types.js'

interface NoteCell {
  cell: string
  row: number
  col: number
  text: string
}

export function noteCells({ cells }: PrintNotes): NoteCell[] {
  if (!cells || typeof cells !== 'object' || Array.isArray(cells)) throw new InputError('notes.cells', 'expected cells by address')
  const out: NoteCell[] = []
  for (const cell in cells) {
    if (!Object.hasOwn(cells, cell)) continue
    const note = cells[cell]?.note
    if (note === undefined) continue
    const position = parseCell(cell)
    if (!position) throw new InputError('notes.cells', 'expected a cell address, got ' + JSON.stringify(cell))
    if (typeof note?.text !== 'string') throw new InputError('notes.cells.' + cell, 'expected note text')
    const { row, col } = position
    out.push({ cell: formatAddress(row - 1, col - 1), row, col, text: note.text })
  }
  return out.sort((a, b) => a.row - b.row || a.col - b.col)
}

function nominalMetrics(font: PrintNotes['font'], scale: number): PrintNoteMetrics {
  const size = font?.size ?? 11
  checkNumber(size, 'notes.font.size', Number.MIN_VALUE, Number.MAX_SAFE_INTEGER)
  const measured = font ? measureFont({ name: font.name, size: 11 }) : undefined
  return {
    lineHeight: DEFAULT_ROW_HEIGHT * (size / 11) * scale,
    measureText(text) {
      let total = 0
      for (const char of text) total += measured ? measured.advance(char.codePointAt(0)!) : advance(char)
      return total * size * scale / 2048
    },
  }
}

function* wrapParagraph(text: string, width: number, measure: PrintNoteMetrics['measureText']): Generator<string> {
  let line = ''
  let spaces = ''
  for (const [token] of text.matchAll(/[ \t]+|[^ \t-]+-?|-+/gu)) {
    if (/^[ \t]/.test(token)) {
      spaces = token
      continue
    }
    const joined = line ? line + spaces + token : token
    if (measure(joined) <= width) {
      line = joined
    } else {
      if (line) yield line
      line = token
      if (measure(line) > width) {
        // Keep a single grapheme on a line even when the printable column is narrower than it.
        const graphemes = typeof Intl.Segmenter === 'function'
          ? Array.from(new Intl.Segmenter(undefined, { granularity: 'grapheme' }).segment(token), ({ segment }) => segment)
          : token
        line = ''
        for (const char of graphemes) {
          if (line && measure(line + char) > width) {
            yield line
            line = ''
          }
          line += char
        }
      }
    }
    spaces = ''
  }
  yield line
}

function* noteLines(notes: readonly NoteCell[], width: number, measure: PrintNoteMetrics['measureText']) {
  for (const { cell, text } of notes) {
    yield { cell, kind: 'address', text: cell, labeled: true } as const
    if (!text) continue
    let labeled = true
    for (const paragraph of text.split(/\r\n?|\n/)) {
      for (const line of wrapParagraph(paragraph, width, measure)) {
        yield { cell, kind: 'body', text: line, labeled } as const
        labeled = false
      }
    }
  }
}

export function* notePages(notes: readonly NoteCell[], options: PrintNotes, body: PrintRect, scale: number): Generator<PrintNoteLine[]> {
  const labels = options.labels ?? { address: 'Cell:', body: 'Note:' }
  if (typeof labels.address !== 'string' || typeof labels.body !== 'string') throw new InputError('notes.labels', 'expected text labels')
  const metrics = options.metrics?.(scale) ?? nominalMetrics(options.font, scale)
  const { lineHeight } = metrics
  checkNumber(lineHeight, 'notes.metrics.lineHeight', Number.MIN_VALUE, Number.MAX_SAFE_INTEGER)
  const measure = (text: string) => {
    const width = metrics.measureText(text)
    checkNumber(width, 'notes.metrics.measureText', 0, Number.MAX_SAFE_INTEGER)
    return width
  }
  const labelWidths = { address: measure(labels.address), body: measure(labels.body) }
  const indent = Math.max(labelWidths.address, labelWidths.body)
  const textX = body.x + indent + measure(' ')
  const textWidth = Math.max(0, body.x + body.width - textX)
  let lines: PrintNoteLine[] = []
  let row = 0

  for (const { cell, kind, text, labeled } of noteLines(notes, textWidth, measure)) {
    // Excel leaves a blank line between notes, except before an entry that starts a fresh page.
    const gap = kind === 'address' && lines.length > 0 ? 1 : 0
    // A line can start at the bottom margin, with its glyphs extending past it.
    if ((row + gap) * lineHeight > body.height + 1e-8) {
      yield lines
      lines = []
      row = 0
    } else {
      row += gap
    }
    const y = body.y + row * lineHeight
    const bounds = { x: textX, y, width: textWidth, height: lineHeight }
    const line: PrintNoteLine = { cell, kind, text, bounds }
    if (labeled) {
      const width = labelWidths[kind]
      line.label = { text: labels[kind], bounds: { x: body.x + indent - width, y, width, height: lineHeight } }
    }
    lines.push(line)
    row++
  }
  if (lines.length) yield lines
}
