/**
 * The workbook cases (`bench/coverage/workbook/`) as the engine takes them: the book with what its file holds beyond
 * the cells (hidden rows and columns, the filter mode of a sheet whose autofilter has criteria), its steps applied
 * in a session as a host would apply them, and the differences of each read from Excel's.
 */
import type { Val } from '../bench/coverage/formulajs.js'
import { compare } from '../bench/coverage/compare.js'
import type { FileParts } from '../bench/coverage/case-book.js'
import type { Step, WorkbookAnswer, WorkbookCase } from '../bench/coverage/workbook/cases.js'
import { takenBook } from '../bench/coverage/stored.js'
import { createSession, type BookInput, type CachedValues, type ComputeResult, type Session, type SheetInput } from '../src/index.js'
import { patchLines, parseSpan } from '../src/layout/spans.js'
import { formatAddress, parseAddress } from '../src/model.js'
import { readDifferences } from './spill-shown.js'

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
  const iteration = f.iterate && { ...(f.iterate.count !== undefined ? { maxIterations: f.iterate.count } : {}), ...(f.iterate.delta !== undefined ? { maxChange: f.iterate.delta } : {}) }
  const book = takenBook(c.book)
  return { ...book, sheets: book.sheets.map((s) => withParts(s, f.sheets?.[s.name])), ...(f.date1904 ? { dateSystem: 1904 as const } : {}), ...(iteration ? { iteration } : {}) }
}

/** The values the case's file stores (`Sheet!A1` → value), as `readXlsx` gives them: none where it stores none. */
export function cachedOf(c: WorkbookCase): CachedValues {
  const stored = Object.entries(c.file?.cached ?? {})
  return { sheets: c.book.sheets.map((s) => ({ name: s.name, cells: Object.fromEntries(stored.filter(([k]) => k.startsWith(`${s.name}!`)).map(([k, value]) => [k.slice(s.name.length + 1), { value }])) })) }
}

/** Excel's autofilter as a host applies it: the rows of `range` below its header shown where `field` is `criterion`, else hidden; the sheet filtered. */
function filterRows(session: Session, sheet: string, range: string, field: number, criterion: string) {
  const [a, b] = range.split(':').map(parseAddress)
  const cells = session.getBook().sheets.find((s) => s.name === sheet)!.cells
  const rows = []
  for (let r = a.row + 1; r <= b.row; r++) rows.push({ sheet, rows: String(r + 1), set: { hidden: String(cells[formatAddress(r, a.col + field - 1)]?.value ?? '') !== criterion } })
  session.setLayout([{ sheet, filtered: true }, ...rows])
}

/** A step in the session; `read` takes the result. */
function apply(session: Session, s: Step, reads: ComputeResult[]) {
  if (s === 'read') reads.push(session.getResult())
  else if (s === 'full' || s === 'rebuild' || s === 'calc' || s === 'save') return
  else if ('set' in s) session.setCells([{ sheet: s.set.sheet, cell: s.set.cell, input: s.set.value === null ? null : { value: s.set.value } }])
  else if ('formula' in s) session.setCells([{ sheet: s.formula.sheet, cell: s.formula.cell, input: { formula: s.formula.formula, ...(s.formula.array ? { array: s.formula.array } : {}) } }])
  else if ('hide' in s) {
    const { sheet, rows, columns, hidden } = s.hide
    session.setLayout([rows !== undefined ? { sheet, rows, set: { hidden } } : { sheet, columns: columns!, set: { hidden } }])
  } else filterRows(session, s.filter.sheet, s.filter.range, s.filter.field, s.filter.criterion)
}

/** The engine's result at each `read` step of the case. */
export function engineReads(c: WorkbookCase): ComputeResult[] {
  const session = createSession(bookOf(c), { clock: CLOCK, cached: cachedOf(c) })
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
  return [...out, ...readDifferences(c.read, excel, r, (x, engine) => compare(x, engine).ok)]
}

/** Whether the engine throws on the case or differs from Excel's reads `excel` at some read. */
export function differsOrThrows(c: WorkbookCase, excel: WorkbookAnswer['reads']): boolean {
  try {
    const reads = engineReads(c)
    return excel.some((a, i) => differences(c, a.values, a.shown, reads[i]).length > 0)
  } catch {
    return true
  }
}

/** A predicted read as Excel's answer: the values, and the shown text of `show:` keys. */
const predictedRead = (p: Record<string, Val | string>): [Record<string, Val>, Record<string, string>] => [
  Object.fromEntries(Object.entries(p).filter(([k]) => !k.startsWith('show:'))) as Record<string, Val>,
  Object.fromEntries(Object.entries(p).flatMap(([k, v]) => (k.startsWith('show:') ? [[k.slice(5), v as string]] : []))),
]

/**
 * A held-out case's predictions against Excel's answers `excel` and the engine: the cells `missed` pins that Excel answered
 * as predicted after all (none, as a miss stays pinned), and the differences of the engine from the predictions, the
 * missed cells holding Excel's answer instead, as the engine gives it since (none).
 */
export function predictionCheck(c: WorkbookCase, excel: WorkbookAnswer['reads'], missed: Record<string, string>): [string[], string[]] {
  const answered = Object.keys(missed).filter((k) => !c.predicted!.some((p, i) => JSON.stringify(p[k]) !== JSON.stringify(excel[i].values[k])))
  const reads = engineReads(c)
  const kept = c.predicted!.map((p, i) => ({ ...p, ...Object.fromEntries(Object.keys(missed).map((k) => [k, excel[i].values[k]])) }))
  return [answered, kept.flatMap((p, i) => differences(c, ...predictedRead(p), reads[i]))]
}

/**
 * A recalculation, a full calculation or an edit computes Excel's cycles again from the values they hold, the
 * engine's from 0 (`BookInput.iteration`): in the sets `iterate` and `spill-cycles-held-2`, a cell of a read after the first may differ where
 * Excel's value moved from its previous read, and must match where it did not.
 */
export const goesOn = (set: string, excel: WorkbookAnswer['reads'], read: number, cell: string) =>
  (set === 'iterate' || set === 'spill-cycles-held-2') && read > 0 && JSON.stringify(excel[read].values[cell]) !== JSON.stringify(excel[read - 1].values[cell])
