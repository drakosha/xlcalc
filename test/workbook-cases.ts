/**
 * The workbook cases (`bench/coverage/workbook/`) as the engine takes them: the book with what its file holds beyond
 * the cells (hidden rows and columns, the filter mode of a sheet whose autofilter has criteria), its steps applied
 * in a session as a host would apply them, and the differences of each read from Excel's.
 */
import type { Val } from '../bench/coverage/formulajs.js'
import { compare } from '../bench/coverage/compare.js'
import type { FileParts } from '../bench/coverage/case-book.js'
import type { Step, WorkbookCase } from '../bench/coverage/workbook/cases.js'
import { openSession, type BookInput, type ComputeResult, type Session, type SheetInput } from '../src/index.js'
import { patchLines, parseSpan } from '../src/layout/spans.js'
import { formatAddress, parseAddress } from '../src/model.js'
import { shownValue } from './spill-shown.js'

export const CLOCK = { now: '2026-10-07T12:00:00Z', seed: 1 }

/** A sheet with the parts of its file: hidden rows and columns, and the filter mode when its autofilter has criteria. */
function withParts(s: SheetInput, parts: NonNullable<FileParts['sheets']>[string] | undefined): SheetInput {
  if (!parts) return s
  let { rows, columns } = s
  for (const k of parts.hiddenRows ?? []) rows = patchLines(rows, 'r', ...parseSpan(k, 'r')!, { hidden: true })
  for (const k of parts.hiddenColumns ?? []) columns = patchLines(columns, 'c', ...parseSpan(k, 'c')!, { hidden: true })
  const filtered = !!parts.filter && Object.keys(parts.filter.keep).length > 0
  return { ...s, ...(rows ? { rows } : {}), ...(columns ? { columns } : {}), ...(filtered ? { filtered } : {}) }
}

/** The case's book as the engine takes it. */
export function bookOf(c: WorkbookCase): BookInput {
  const f = c.file ?? {}
  return { ...c.book, sheets: c.book.sheets.map((s) => withParts(s, f.sheets?.[s.name])), ...(f.date1904 ? { dateSystem: 1904 as const } : {}) }
}

/** Excel's autofilter as a host applies it: the rows of `range` below its header shown where `field` is `criterion`, else hidden; the sheet filtered. */
function filterRows(session: Session, sheet: string, range: string, field: number, criterion: string) {
  const [a, b] = range.split(':').map(parseAddress)
  const cells = session.book().sheets.find((s) => s.name === sheet)!.cells
  const rows = []
  for (let r = a.row + 1; r <= b.row; r++) rows.push({ sheet, rows: String(r + 1), set: { hidden: String(cells[formatAddress(r, a.col + field - 1)]?.value ?? '') !== criterion } })
  session.setLayout([{ sheet, filtered: true }, ...rows])
}

/** A step in the session; `read` takes the result. */
function apply(session: Session, s: Step, reads: ComputeResult[]) {
  if (s === 'read') reads.push(session.result())
  else if (s === 'full' || s === 'calc') return
  else if ('set' in s) session.setCells([{ sheet: s.set.sheet, cell: s.set.cell, input: s.set.value === null ? null : { value: s.set.value } }])
  else if ('formula' in s) session.setCells([{ sheet: s.formula.sheet, cell: s.formula.cell, input: { formula: s.formula.formula } }])
  else if ('hide' in s) {
    const { sheet, rows, columns, hidden } = s.hide
    session.setLayout([rows !== undefined ? { sheet, rows, set: { hidden } } : { sheet, columns: columns!, set: { hidden } }])
  } else filterRows(session, s.filter.sheet, s.filter.range, s.filter.field, s.filter.criterion)
}

/** The engine's result at each `read` step of the case. */
export function engineReads(c: WorkbookCase): ComputeResult[] {
  const session = openSession(bookOf(c), { clock: CLOCK })
  const reads: ComputeResult[] = []
  for (const s of c.steps ?? ['full', 'read']) apply(session, s, reads)
  session.close()
  return reads
}

/** Cells of the case's rectangles where the engine does not show Excel's answer, and cells to show that show other text. */
export function differences(c: WorkbookCase, excel: Record<string, Val>, shown: Record<string, string> | undefined, r: Pick<ComputeResult, 'sheets'>): string[] {
  const out: string[] = []
  for (const [k, text] of Object.entries(shown ?? {})) {
    const [sheet, addr] = k.split('!')
    const ours = r.sheets.find((s) => s.name === sheet)!.cells[addr]?.text ?? ''
    if (ours !== text) out.push(`${k} shows ${JSON.stringify(text)} in Excel, ${JSON.stringify(ours)} here`)
  }
  for (const g of c.read) {
    const [p, q = p] = g.range.split(':').map(parseAddress)
    for (let row = p.row; row <= q.row; row++)
      for (let col = p.col; col <= q.col; col++) {
        const addr = formatAddress(row, col)
        const x = excel[`${g.sheet}!${addr}`] ?? null
        const engine = shownValue(r, g.sheet, addr)
        const ok = x === null || engine === null ? x === engine : compare(x, engine).ok
        if (!ok) out.push(`${g.sheet}!${addr}: Excel ${JSON.stringify(x)}, engine ${JSON.stringify(engine)}`)
      }
  }
  return out
}
