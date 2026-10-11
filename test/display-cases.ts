/**
 * The display cases measured in Excel (`test/excel/display/`, a locale's in a directory of its tag: `run.ts display
 * <set> [locale]`) as books: each case's cell where the harness put it, in a column of its width (a merged area of
 * its span), computed in the locale of the answers; and how an engine's shown text and colour compare with Excel's.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { place, type StoredDisplay } from '../bench/coverage/display.js'
import type { DisplayAnswer, DisplayCase } from '../bench/coverage/display/cases.js'
import { localeDir, TAGS } from '../bench/coverage/locales.js'
import type { BookInput, CellDisplay, CellInput, CellValue, SheetInput } from '../src/contract.js'
import { formatAddress, indexToCol } from '../src/model.js'
import './locales.js'

export const CLOCK = { now: '2026-10-07T12:00:00Z', seed: 1 }

export interface DisplaySetFile extends StoredDisplay {
  name: string
  /** The locale Excel showed the cells in. */
  locale: string
}

const dir = new URL('./excel/display/', import.meta.url)
const read = (sub: string, locale: DisplaySetFile['locale']): DisplaySetFile[] => {
  const d = new URL(sub, dir)
  if (!existsSync(d)) return []
  return readdirSync(d)
    .filter((f) => f.endsWith('.json'))
    .sort()
    .map((f) => ({ name: f.slice(0, -5), locale, ...(JSON.parse(readFileSync(new URL(f, d), 'utf8')) as StoredDisplay) }))
}

/** The stored answers of Excel in each locale that has them. */
export const DISPLAY_SETS: DisplaySetFile[] = TAGS.flatMap((tag) => read(`./${localeDir(tag)}`, tag))

/** A case's cell as the engine takes it: a built-in format by its number. */
export function formatted(c: DisplayCase): CellInput {
  // a stored case names the cell's number format `format`, as the engine's input did when it was stored
  const { format, ...cell } = c.cell
  const numberFormat = c.numFmtId ?? format
  return numberFormat === undefined ? cell : { ...cell, numberFormat }
}

/** The book of a set's cases, as the harness lays them out, and each case's address. */
export function displayBook(cases: DisplayCase[], locale: string): { book: BookInput; at: Map<string, string> } {
  const sheet: Required<Pick<SheetInput, 'cells' | 'columns' | 'merges'>> & SheetInput = { name: 'S', cells: {}, columns: {}, merges: [] }
  const at = new Map<string, string>()
  const placed = place(cases)
  for (const c of cases) {
    const { row, col } = placed.get(c.id)!
    const addr = formatAddress(row, col)
    at.set(c.id, addr)
    sheet.cells[addr] = formatted(c)
    const span = c.span ?? 1
    for (let k = 0; k < span; k++) if (c.width !== undefined) sheet.columns[indexToCol(col + k)] = { width: c.width }
    if (span > 1) sheet.merges.push(`${addr}:${formatAddress(row, col + span - 1)}`)
  }
  return { book: { locale, sheets: [sheet] }, at }
}

/**
 * Whether the engine shows what Excel showed. Excel reads the automatic colour and `[Black]` alike (black is left
 * out of the answers), and through AppleScript gives no colour for an exact 0 in a format of fewer than three
 * sections or for text, where the screen shows the section's colour (`docs/formats.md`): there any colour passes.
 */
export function sameShown(excel: DisplayAnswer, engine: CellDisplay, value: CellValue): boolean {
  if (excel.text !== engine.text) return false
  const color = engine.color === '#000000' ? undefined : engine.color
  return excel.color === color || (excel.color === undefined && (value === 0 || typeof value === 'string' || typeof value === 'boolean'))
}
