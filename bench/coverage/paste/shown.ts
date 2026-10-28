/**
 * What a book of the step `paste` shows, comparable across files: over a window of each sheet each cell's contents,
 * note, hyperlink, style and number format (of its entry, the sheet's looks or its lines), and the sheet's merged areas,
 * conditional formats, data validation and column widths; and the engine's outcome of a case as `test/paste-excel.test.ts`
 * compares it, written and read back.
 */
import type { BookInput, PasteKind, SheetInput } from '../../../src/contract.js'
import { createSession, toLocal, validateFormula } from '../../../src/index.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import { lookOf } from '../../../src/layout/lines.js'
import { entryAt, lines } from '../../../src/layout/spans.js'
import { inputIn } from '../../../src/looks/sheet.js'
import { formatAddress } from '../../../src/model.js'
import type { PasteCase } from './cases.js'

/** The kinds of paste the engine has. */
export const ENGINE_KINDS: readonly string[] = ['copy', 'all', 'values', 'formulas', 'formats', 'validation', 'valuesAndNumberFormats', 'formulasAndNumberFormats', 'columnWidths', 'allExceptBorders', 'notes', 'allMergingConditionalFormats', 'painter'] satisfies (PasteKind | 'copy' | 'painter')[]

/** Whether the engine has a case's paste. */
export const engineHas = (c: PasteCase): boolean => ENGINE_KINDS.includes(c.what)

const WINDOW = { rows: 30, cols: 20 }

/** A formula in one spelling: Excel saves `T!A1` where the writer writes `'T'!A1`. */
function canonical(formula: string): string {
  const v = validateFormula(formula)
  return v.ok ? v.canonical : formula
}

/** What a sheet shows over the window and holds besides. */
function shownSheet(book: BookInput, sheet: SheetInput) {
  const [rows, cols] = [lines(sheet.rows, 'r'), lines(sheet.columns, 'c')]
  const cells: Record<string, unknown> = {}
  for (let r = 0; r < WINDOW.rows; r++)
    for (let c = 0; c < WINDOW.cols; c++) {
      const a = formatAddress(r, c)
      const cell = sheet.cells[a]
      const look = lookOf(inputIn(sheet, a, r, c), entryAt(rows, r), entryAt(cols, c))
      const style = look.style === undefined ? {} : (book.styles?.[look.style] ?? {})
      const x = {
        ...(cell?.formula !== undefined ? { formula: canonical(cell.formula) } : {}),
        ...(cell?.value !== undefined ? { value: cell.value } : {}),
        ...(cell?.array !== undefined ? { array: cell.array } : {}),
        ...(cell?.note ? { note: cell.note.text } : {}),
        ...(cell?.link ? { link: cell.link.url ?? cell.link.location } : {}),
        ...(Object.keys(style).length ? { style } : {}),
        ...(look.numberFormat !== undefined && look.numberFormat !== 'General' && look.numberFormat !== 0 ? { numberFormat: look.numberFormat } : {}),
      }
      if (Object.keys(x).length) cells[a] = x
    }
  const sorted = <T>(xs: T[] | undefined) => (xs ?? []).map((x) => JSON.stringify(x)).sort()
  return {
    cells,
    merges: [...(sheet.merges ?? [])].sort(),
    rules: sorted(sheet.conditionalFormats?.map((f) => ({ ...f, ranges: [...f.ranges].sort() }))),
    validations: sorted(sheet.validations),
    widths: Object.fromEntries(Object.entries(sheet.columns ?? {}).flatMap(([k, v]) => (v.width === undefined ? [] : [[k, v.width]]))),
    heights: Object.fromEntries(Object.entries(sheet.rows ?? {}).flatMap(([k, v]) => (v.height === undefined || Number(k) > WINDOW.rows ? [] : [[k, v.height]]))),
    tables: (sheet.tables ?? []).map((t) => `${t.name} ${t.range} ${t.columns.map((c) => c.name).join(',')}`),
  }
}

/** A book as a file shows it: every sheet by name. */
export const shownBook = (book: BookInput) => Object.fromEntries(book.sheets.map((s) => [s.name, shownSheet(book, s)]))

/** What a book shows (`shownBook`). */
export type Shown = ReturnType<typeof shownBook>

/** What the file of a book shows, written and read back. */
export const shownFile = async (data: Uint8Array) => shownBook((await readXlsx(data)).book)

/**
 * The engine's outcome of a case in a session, as a host pastes (a session fits the rows a copy laid its looks on): what
 * its book shows after the paste, written and read back, or why it is refused.
 */
export async function engineOutcome(book: BookInput, c: PasteCase): Promise<{ refused: string } | { shown: ReturnType<typeof shownBook> }> {
  // the harness pastes with Paste Special, `all` too; `copy` is a copy without it; the Format Painter is a paste of formats
  const paste = c.what === 'painter' ? 'formats' : (c.what as PasteKind)
  const op = { op: 'copyRange' as const, from: c.from, to: c.to, ...(c.what === 'copy' ? {} : { paste }), ...(c.transpose ? { transpose: true } : {}), ...(c.operation ? { operation: c.operation } : {}), ...(c.skipBlanks ? { skipBlanks: true } : {}) }
  const s = createSession(book, { clock: { now: '2026-10-27T12:00:00Z', seed: 1 } })
  const r = s.editStructure(op)
  return r.ok ? { shown: await shownFile((await writeXlsx({ book: s.getBook() })).data) } : { refused: r.message }
}

/** The formulas of a case's `local` cells in a book that shows them, as Excel reads them in the book's language (`FormulaLocal`). */
export function localFormulas(shown: Shown, c: PasteCase, locale: string): Record<string, string> {
  const cells = shown[c.to.sheet].cells as Record<string, { formula?: string }>
  return Object.fromEntries((c.local ?? []).map((a) => {
    const f = cells[a]?.formula
    const t = f === undefined ? null : toLocal(f, locale)
    return [a, t?.ok ? `=${t.formula}` : '']
  }))
}
