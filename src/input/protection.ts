/**
 * The checks of protection, a sheet's (`SheetProtection`) and a cell's (`CellProtection`), and their canonical forms:
 * only what differs from Excel's defaults, so `{ locked: true }` is no protection of a cell and `{ formatCells: false }`
 * Excel's default protection of a sheet. The xlsx reader and writer take the fields and defaults from here.
 */
import type { CellProtection, CellStyle, SheetProtection } from '../contract.js'
import { checkBoolean, object, onlyKeys } from './basic.js'

/** What a protected sheet allows by default, as Excel's Protect Sheet dialog offers it: selecting cells, nothing else. */
export const SHEET_PROTECTION: Readonly<Record<keyof SheetProtection, boolean>> = {
  selectLockedCells: true,
  selectUnlockedCells: true,
  formatCells: false,
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  insertHyperlinks: false,
  deleteColumns: false,
  deleteRows: false,
  sort: false,
  autoFilter: false,
  pivotTables: false,
  objects: false,
  scenarios: false,
}

/** A cell's protection by default: locked, not hidden. */
export const CELL_PROTECTION: Readonly<Record<keyof CellProtection, boolean>> = { locked: true, hidden: false }

/** The boolean fields of `o` that differ from `defaults`: `o` itself when none is a default or undefined. */
function differing<T extends object>(o: T, defaults: Readonly<Record<keyof T, boolean>>): T {
  const keys = Object.keys(o) as (keyof T)[]
  const kept = keys.filter((k) => o[k] !== undefined && o[k] !== defaults[k])
  return kept.length === keys.length ? o : (Object.fromEntries(kept.map((k) => [k, o[k]])) as T)
}

/** A sheet's protection as the workbook holds it: the fields that differ from Excel's default protection. */
export const protectionTaken = (p: SheetProtection): SheetProtection => differing(p, SHEET_PROTECTION)

/** A style as the workbook holds it: its protection without the defaults, none when nothing differs; the style itself when it has none to drop. */
export function styleTaken(s: CellStyle): CellStyle {
  if (s.protection === undefined) return s
  const protection = differing(s.protection, CELL_PROTECTION)
  const empty = !Object.keys(protection).length
  if (protection === s.protection && !empty) return s
  const out: CellStyle = { ...s, protection }
  if (empty) delete out.protection
  return out
}

/** A sheet's protection: its fields of what users may do, each a boolean. */
export function checkSheetProtection(v: unknown, p: string) {
  const o = object(v, p, 'a protection such as { formatCells: true } ({} protects the sheet as Excel does by default)')
  onlyKeys(o, p, Object.keys(SHEET_PROTECTION))
  for (const [k, x] of Object.entries(o)) if (x !== undefined) checkBoolean(x, `${p}.${k}`)
}
