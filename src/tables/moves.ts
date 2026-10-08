/**
 * Tables under cut and paste, copy and paste (the set `tables/ops`): a table whose cells all move goes with them; a
 * formula that leaves the table its references without a table name meant gets the table's name (`[@Qty]` pasted
 * outside reads `Sales[@Qty]`). Pasting over a table's header row, or cutting part of one, would rename or break its
 * columns: refused with `table`.
 */
import type { BookInput, CellInput, RestructureResult, SheetInput, TableInput } from '../contract.js'
import { formatAddress, parseAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import { contains, meets, type Rect } from '../structure/area.js'
import { rewriteTableRefs } from './rewrite.js'
import { tableAt, tableKey, tablesOf, type Table, type Tables } from './model.js'

const headerRow = (t: Table): Rect | null => (t.headers ? { r1: t.r1, r2: t.r1, c1: t.c1, c2: t.c2 } : null)

/** Why a paste over `dst` of sheet `s` is refused: it meets the header row of a table that does not move along. */
function headerRefusal(tables: Tables, s: number, dst: Rect, moving: Set<Table>): string | null {
  const t = tables.list.find((x) => x.s === s && !moving.has(x) && headerRow(x) && meets(headerRow(x)!, dst))
  return t ? `the paste would land on the header row of the table ${t.name}` : null
}

/** The formula with its references without a table name given `t`'s name (only those change). */
const qualified = (formula: string, tables: Tables, t: Table) => rewriteTableRefs(formula, tables, t, null, () => undefined, true)

/** Cut `src` of sheet `from` and paste it at `to`, `dr` rows and `dc` columns away: the tables it carries, refusals. */
export function moveTables(book: BookInput, from: string, src: Rect, to: string, dr: number, dc: number): { refuse?: string; before(b: BookInput): BookInput; after(b: BookInput): BookInput } {
  const tables = tablesOf(book)
  const [fs, ts] = [book.sheets.findIndex((s) => sheetKey(s.name) === sheetKey(from)), book.sheets.findIndex((s) => sheetKey(s.name) === sheetKey(to))]
  const dst: Rect = { r1: src.r1 + dr, r2: src.r2 + dr, c1: src.c1 + dc, c2: src.c2 + dc }
  const moving = new Set(tables.list.filter((t) => t.s === fs && contains(src, t)))
  const keep = { before: (b: BookInput) => b, after: (b: BookInput) => b }
  const cut = tables.list.find((t) => t.s === fs && !moving.has(t) && headerRow(t) && meets(headerRow(t)!, src))
  if (cut) return { ...keep, refuse: `the cut takes part of the header row of the table ${cut.name}` }
  const over = headerRefusal(tables, ts, dst, moving)
  if (over) return { ...keep, refuse: over }
  const land = [...moving].map((t) => ({ t, x: { r1: t.r1 + dr, r2: t.r2 + dr, c1: t.c1 + dc, c2: t.c2 + dc } }))
  const onto = land.find(({ x }) => tables.list.some((o) => o.s === ts && !moving.has(o) && meets(o, x)))
  if (onto) return { ...keep, refuse: `the table ${onto.t.name} would land on another table` }
  if (!moving.size && !tables.list.some((t) => t.s === fs && meets(t, src))) return keep
  // the table a moved cell is in after the move: one that moved along, or one that stays
  const after = (r: number, c: number): string | undefined => {
    const m = land.find(({ x }) => r >= x.r1 && r <= x.r2 && c >= x.c1 && c <= x.c2)
    if (m) return m.t.name
    const t = tableAt(tables, ts, r, c)
    return t && !moving.has(t) ? t.name : undefined
  }
  return {
    before(b) {
      const sheet = b.sheets[fs]
      let cells: Record<string, CellInput> | undefined
      for (const [addr, cell] of Object.entries(sheet.cells)) {
        if (cell.formula === undefined) continue
        const { row, col } = parseAddress(addr)
        if (row < src.r1 || row > src.r2 || col < src.c1 || col > src.c2) continue
        const t = tableAt(tables, fs, row, col)
        if (!t || moving.has(t) || after(row + dr, col + dc) === t.name) continue
        const formula = qualified(cell.formula, tables, t)
        if (formula !== cell.formula) (cells ??= { ...sheet.cells })[addr] = { ...cell, formula }
      }
      if (!cells) return b
      const sheets = [...b.sheets]
      sheets[fs] = { ...sheet, cells }
      return { ...b, sheets }
    },
    after(b) {
      if (!moving.size) return b
      const gone = new Set([...moving].map((t) => tableKey(t.name)))
      const sheets: SheetInput[] = b.sheets.map((s) => {
        const left = (s.tables ?? []).filter((t) => !gone.has(tableKey(t.name)))
        const out: SheetInput = { ...s, tables: left }
        if (!left.length) delete out.tables
        return out
      })
      const moved: TableInput[] = land.map(({ t, x }) => ({ ...t.input, range: `${formatAddress(x.r1, x.c1)}:${formatAddress(x.r2, x.c2)}` }))
      sheets[ts] = { ...sheets[ts], tables: [...(sheets[ts].tables ?? []), ...moved] }
      return { ...b, sheets }
    },
  }
}

/**
 * A copy of `src` of sheet `from` pasted over `target` of `to`: refused over a table's header row; afterwards the pasted
 * formulas that leave their source's table name it.
 */
export function copyTables(book: BookInput, from: string, src: Rect, to: string, target: Rect): { refuse?: string; after(r: RestructureResult): RestructureResult } {
  const tables = tablesOf(book)
  const [fs, ts] = [book.sheets.findIndex((s) => sheetKey(s.name) === sheetKey(from)), book.sheets.findIndex((s) => sheetKey(s.name) === sheetKey(to))]
  const over = headerRefusal(tables, ts, target, new Set())
  if (over) return { refuse: over, after: (r) => r }
  if (!tables.list.some((t) => t.s === fs && meets(t, src))) return { after: (r) => r }
  const [h, w] = [src.r2 - src.r1 + 1, src.c2 - src.c1 + 1]
  return {
    after(r) {
      if (!r.ok) return r
      const sheet = r.book.sheets[ts]
      let cells: Record<string, CellInput> | undefined
      for (const [addr, cell] of Object.entries(sheet.cells)) {
        if (cell.formula === undefined) continue
        const { row, col } = parseAddress(addr)
        if (row < target.r1 || col < target.c1 || row > Math.max(target.r2, target.r1 + h - 1) || col > Math.max(target.c2, target.c1 + w - 1)) continue
        const t = tableAt(tables, fs, src.r1 + ((row - target.r1) % h), src.c1 + ((col - target.c1) % w))
        if (!t || tableAt(tables, ts, row, col) === t) continue
        const formula = qualified(cell.formula, tables, t)
        if (formula !== cell.formula) (cells ??= { ...sheet.cells })[addr] = { ...cell, formula }
      }
      if (!cells) return r
      const sheets = [...r.book.sheets]
      sheets[ts] = { ...sheet, cells }
      return { ...r, book: { ...r.book, sheets } }
    },
  }
}
