/**
 * The engine's answer to a rows case, the same for the predictions written before Excel runs a set and for the tests:
 * the case's book in a session, its texts typed (`setCellText`), entered into a selection (`fillCellText`) or pasted
 * (`pasteText`) in turn, its row fitted (`setLayout` with `{ rows, fit: true }`) when the case fits it, then the row's
 * height before and after and its cells' content and wrapping. The engine's functions come as arguments, so a test
 * passes the ones it tests (`src` or `dist`).
 */
import type { BookInput, CellInput, Session, SessionOptions } from '../../../src/contract.js'
import { formatAddress } from '../../../src/model.js'
import { caseCells, caseRow, cellKey, rowsBook, type RowCase } from './cases.js'

export interface RowsEngine {
  createSession(book: BookInput, opts?: SessionOptions): Session
}

/** The engine's answer: the row's height when the book opens and after (the default 15 for none, 0 for a hidden row) and each read cell's content (as `Formula` reads it) and wrapping. */
export interface RowsPrediction {
  before: number
  height: number
  cells: Record<string, { formula: string; wrap: boolean }>
}

/** What `Range.Formula` reads of a cell: its formula after `=`, else its value as text (TRUE, an error's text). */
function formulaOf(input: CellInput | undefined): string {
  if (input?.formula !== undefined) return `=${input.formula}`
  const v = input?.value
  return v === undefined || v === null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v)
}

/** The case's answer in a session of the locale `tag` whose clock stands at `now`. */
export function engineRows(e: RowsEngine, c: RowCase, tag: string, now: string): RowsPrediction {
  const book = { ...rowsBook([c]), locale: tag }
  const s = e.createSession(book, { clock: { now: `${now}T12:00:00Z`, seed: 1 } })
  const row = caseRow(0)
  const shown = () => {
    const r = s.getLayout().sheets[0].rows?.[String(row)]
    return r?.hidden || r?.height === 0 ? 0 : (r?.height ?? 15)
  }
  const at = (col: number, dr = 0) => formatAddress(row - 1 + dr, 1 + col)
  const before = shown()
  for (const [col, text] of c.typed ?? []) {
    // a text Excel does not take leaves the cells as they were, as the session's refusal does
    if (c.via === 'paste') s.pasteText('S', at(col), text)
    else if (c.via === 'fill') s.fillCellText('S', `${at(col)}:${at(c.fillTo!)}`, text)
    else s.setCellText('S', at(col), text)
  }
  if (!c.typed || c.fitAfter) s.setLayout([{ sheet: 'S', rows: String(row), fit: true }])
  const out = s.getBook()
  const styles = out.styles ?? []
  const cells = Object.fromEntries(
    caseCells(c).map((x) => {
      const input = out.sheets[0].cells[at(x[1], x[0])]
      const style = s.getViewport('S', at(x[1], x[0])).cells[at(x[1], x[0])]?.effectiveStyle
      return [cellKey(x), { formula: formulaOf(input), wrap: !!(style !== undefined && styles[style]?.alignment?.wrap) }]
    }),
  )
  const height = shown()
  s.close()
  return { before, height, cells }
}
