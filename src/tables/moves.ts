/**
 * Tables under cut and paste, copy and paste (the sets `tables/ops` and `ops-held`): a table whose cells all move goes
 * with them; a formula that leaves the table its references without a table name meant gets the table's name (`[@Qty]`
 * pasted outside reads `Sales[@Qty]`). A paste over a table's header cells renames their columns to what lands there,
 * the text of a value or of a formula's value; header cells cut away leave their columns named `Column1` and so on.
 */
import type { BookInput, CellInput, SheetInput, TableInput } from '../contract.js'
import { createState } from '../engine.js'
import { currentClock } from '../env.js'
import { spend } from '../exec/budget.js'
import type { Message } from '../locale/messages.js'
import { formatAddress, parseAddress, readAddress } from '../model.js'
import { clockSerial } from '../result.js'
import { ERROR_TEXT } from '../syntax.js'
import { isErr } from '../values.js'
import { headerEdits } from './headers.js'
import { contains, cornersText, hasCell, meets, shiftRect, type Rect } from '../structure/area.js'
import { copyTransform, tiles } from '../structure/copy.js'
import type { Applied } from '../structure/geometry.js'
import { cellsIn, gridOf } from '../structure/grid.js'
import { rewriteFormula, sheetIndex, sheetNames } from '../structure/rewrite.js'
import { rewriteTableRefs } from './rewrite.js'
import { tableAt, tableKey, tablesOf, type Table, type Tables } from './model.js'

const headerRow = (t: Table): Rect | null => (t.headers ? { r1: t.r1, r2: t.r1, c1: t.c1, c2: t.c2 } : null)

/**
 * The book after a paste with the columns of the header cells in `areas` (sheet index and rectangle) named by what the
 * paste left there, as Excel renames them: a value's text, a formula's value as text (the book is computed for it), a
 * cleared cell `Column1`. `still` are the tables whose header cells moved along with them.
 */
function renamedByPaste(tables: Tables, after: BookInput, areas: [number, Rect][], still: (t: Table) => boolean): BookInput {
  const cells: { s: number; addr: string }[] = []
  for (const t of tables.list) {
    const head = headerRow(t)
    if (!head || !still(t)) continue
    for (let c = t.c1; c <= t.c2; c++) if (areas.some(([s, x]) => s === t.s && meets(x, { r1: t.r1, r2: t.r1, c1: c, c2: c }))) cells.push({ s: t.s, addr: formatAddress(t.r1, c) })
  }
  if (!cells.length) return after
  const clock = currentClock()
  const state = createState(after, { now: clockSerial(clock), seed: clock.seed })
  const text = (v: unknown) => (isErr(v as never) ? ERROR_TEXT[(v as { code: keyof typeof ERROR_TEXT }).code] : String(state.ctx.rules.toText(v as never)))
  const edits = cells.map(({ s, addr }) => {
    const input = after.sheets[s].cells[addr]
    const { row, col } = parseAddress(addr)
    const v = input?.formula !== undefined ? state.rt.cell(s, row, col) : (input?.value ?? null)
    const value = v === null || v === undefined || v === '' ? null : typeof v === 'string' ? v : typeof v === 'object' && 'date' in (v as object) ? (v as never) : text(v)
    return { sheet: after.sheets[s].name, cell: addr, input: value === null ? null : { value } }
  })
  return headerEdits(after, edits, text) ?? after
}

/** The formula with its references without a table name given `t`'s name (only those change). */
const qualified = (formula: string, tables: Tables, t: Table) => rewriteTableRefs(formula, tables, t, null, () => undefined, true)

/** Cut `src` of sheet `from` and paste it at `to`, `dr` rows and `dc` columns away: the tables it carries, refusals. */
export function moveTables(book: BookInput, from: string, src: Rect, to: string, dr: number, dc: number): { refuse?: Message; before(b: BookInput): BookInput; after(b: BookInput): BookInput } {
  const tables = tablesOf(book)
  const [fs, ts] = [sheetIndex(book, from), sheetIndex(book, to)]
  const dst = shiftRect(src, dr, dc)
  const moving = new Set(tables.list.filter((t) => t.s === fs && contains(src, t)))
  const keep = { before: (b: BookInput) => b, after: (b: BookInput) => b }
  const land = [...moving].map((t) => ({ t, x: shiftRect(t, dr, dc) }))
  const onto = land.find(({ x }) => tables.list.some((o) => o.s === ts && !moving.has(o) && meets(o, x)))
  if (onto) return { ...keep, refuse: { key: 'tableLands', table: onto.t.name } }
  if (!moving.size && !tables.list.some((t) => (t.s === fs && meets(t, src)) || (t.s === ts && meets(t, dst)))) return keep
  // the table a moved cell is in after the move: one that moved along, or one that stays
  const after = (r: number, c: number): string | undefined => {
    const m = land.find(({ x }) => hasCell(x, r, c))
    if (m) return m.t.name
    const t = tableAt(tables, ts, r, c)
    return t && !moving.has(t) ? t.name : undefined
  }
  return {
    before(b) {
      const sheet = b.sheets[fs]
      let cells: Record<string, CellInput> | undefined
      const g = gridOf(sheet)
      for (const i of cellsIn(g, src)) {
        const [addr, cell, row, col] = [g.addr[i], g.cell[i], g.row[i], g.col[i]]
        if (cell.formula === undefined) continue
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
    after(done) {
      // header cells cut away or pasted over name their columns anew
      const b = renamedByPaste(tables, done, [[fs, src], [ts, dst]], (t) => !moving.has(t))
      if (!moving.size) return b
      const gone = new Set([...moving].map((t) => tableKey(t.name)))
      const sheets: SheetInput[] = b.sheets.map((s) => {
        const left = (s.tables ?? []).filter((t) => !gone.has(tableKey(t.name)))
        const out: SheetInput = { ...s, tables: left }
        if (!left.length) delete out.tables
        return out
      })
      // each moved table as the move left it (its calculated columns rewritten with its cells), at its new place
      const own = (t: Table) => b.sheets[fs].tables!.find((ti) => tableKey(ti.name) === tableKey(t.name))!
      const moved: TableInput[] = land.map(({ t, x }) => ({ ...own(t), range: cornersText(x) }))
      sheets[ts] = { ...sheets[ts], tables: [...(sheets[ts].tables ?? []), ...moved] }
      return { ...b, sheets }
    },
  }
}

/**
 * A copy of `src` of sheet `from` pasted over `target` of `to`: a table the copy takes whole becomes a new table at each
 * paste, named after it with the book's count of tables and two (`Sales3` from a book of one table), its formulas'
 * references to the table taken to the new one; the other pasted formulas that leave their source's table name it, and header cells pasted over rename their
 * columns.
 */
export function copyTables(book: BookInput, from: string, src: Rect, to: string, target: Rect): { after(r: Applied): Applied } {
  const tables = tablesOf(book)
  const [fs, ts] = [sheetIndex(book, from), sheetIndex(book, to)]
  const [h, w] = [src.r2 - src.r1 + 1, src.c2 - src.c1 + 1]
  // the tiles of the paste and the area they cover, as copyRange lays them (the set `paste-names-held`: one copy at the
  // top-left cell of a target that is no multiple of the source, the target's other cells left as they were)
  const { offsets, area: pasted } = tiles(src, target)
  const renamed = (r: Applied): Applied => (r.ok ? { ...r, book: renamedByPaste(tables, r.book, [[ts, pasted]], () => true) } : r)
  if (!tables.list.some((t) => t.s === fs && meets(t, src))) return { after: renamed }
  // the tables copied whole onto each tile
  const whole = tables.list.filter((t) => t.s === fs && contains(src, t))
  const taken = tables.list.map((t) => t.name)
  const copies: { t: Table; x: Rect; name: string }[] = []
  for (const [dr, dc] of offsets)
    for (const t of whole) {
      const x = shiftRect(t, dr, dc)
      if (tables.list.some((o) => o.s === ts && meets(o, x)) || copies.some((o) => meets(o.x, x))) continue
      // Excel numbers a copy by the book's count of tables and two (the set `more`: Sales3 in a book of one table)
      let n = taken.length + 2
      while (taken.some((k) => tableKey(k) === tableKey(`${t.name}${n}`))) n++
      taken.push(`${t.name}${n}`)
      copies.push({ t, x, name: `${t.name}${n}` })
    }
  return {
    after(r) {
      if (!r.ok) return r
      const sheet = r.book.sheets[ts]
      let cells: Record<string, CellInput> | undefined
      // a sheet the paste made: its formulas alone are read for their places, which no other step reads
      const p = { r: 0, c: 0 }
      for (const addr in sheet.cells) {
        const cell = sheet.cells[addr]
        spend(1)
        if (cell.formula === undefined) continue
        readAddress(addr, p)
        const [row, col] = [p.r, p.c]
        if (!hasCell(pasted, row, col)) continue
        const t = tableAt(tables, fs, src.r1 + ((row - target.r1) % h), src.c1 + ((col - target.c1) % w))
        if (!t || tableAt(tables, ts, row, col) === t) continue
        // inside a copy of its table: the references to the table are to the copy, those without a name stay so
        const copy = copies.find((c) => c.t === t && hasCell(c.x, row, col))
        const formula = copy ? rewriteTableRefs(cell.formula, tables, t, null, (x) => (x === t ? { table: copy.name } : undefined)) : qualified(cell.formula, tables, t)
        if (formula !== cell.formula) (cells ??= { ...sheet.cells })[addr] = { ...cell, formula }
      }
      const sheets = [...r.book.sheets]
      sheets[ts] = { ...sheet, ...(cells ? { cells } : {}) }
      // a copy's calculated columns: relative references move by its offset as its cells' do, references to the table go to the copy
      const names = sheetNames(r.book)
      const template = (f: string, t: Table, x: Rect, name: string) => rewriteTableRefs(rewriteFormula(f, to, to, copyTransform(x.r1 - t.r1, x.c1 - t.c1), names), tables, t, null, (y) => (y === t ? { table: name } : undefined))
      if (copies.length)
        sheets[ts] = {
          ...sheets[ts],
          tables: [...(sheets[ts].tables ?? []), ...copies.map(({ t, x, name }) => ({ ...t.input, name, range: cornersText(x), columns: t.input.columns.map((c) => (c.formula === undefined ? c : { ...c, formula: template(c.formula, t, x, name) })) }))],
        }
      return renamed({ ...r, book: { ...r.book, sheets } })
    },
  }
}
