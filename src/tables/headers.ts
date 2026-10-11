/**
 * A table's header cell typed over renames its column, as Excel does (the set `tables/ops`): the structured references
 * to the column follow, a number or a boolean becomes its text, a cleared cell names the column `Column1` (the first
 * free number), and a name another column has already gives the column further right a number: `Price2`, the first
 * free one. Columns are named from left to right, so a typed name keeps it and the column to its right gives way.
 */
import type { BookInput, CellInput, CellReplace, NumberFormat, Scalar } from '../contract.js'
import { InputError } from '../input.js'
import { formatAddress, inputScalar, parseAddress } from '../model.js'
import { sheetKey } from '../syntax.js'
import { rewriteBookTables } from './rewrite.js'
import { freeColumnName } from './shift.js'
import { tableKey, tablesOf } from './model.js'

/**
 * The book with `edits` applied when some of them type over header cells: the columns renamed, the references to them
 * rewritten, the header cells holding the names; null when no edit touches a header cell. `text` writes a number or a
 * boolean as the book's locale does, by the cell's own number format when it has one.
 */
export function headerEdits(book: BookInput, edits: CellReplace[], text: (v: Scalar, format?: NumberFormat) => string): BookInput | null {
  const tables = tablesOf(book).list.filter((t) => t.headers)
  if (!tables.length) return null
  const typed = new Map<string, Map<number, string | null>>()
  edits.forEach((e, i) => {
    const s = book.sheets.findIndex((x) => sheetKey(x.name) === sheetKey(e.sheet))
    const { row, col } = parseAddress(e.cell)
    const t = tables.find((x) => x.s === s && x.r1 === row && col >= x.c1 && col <= x.c2)
    if (!t) return
    if (e.input?.formula !== undefined) throw new InputError(`edits[${i}].input.formula`, `${e.sheet}!${e.cell} is a header cell of the table ${t.name}: it holds its column's name, not a formula`)
    const v = inputScalar(e.input?.value)
    const name = v === null || v === '' ? null : typeof v === 'string' ? v : text(v, e.input?.format)
    if (!typed.has(t.name)) typed.set(t.name, new Map())
    typed.get(t.name)!.set(col - t.c1, name)
  })
  if (!typed.size) return null
  // the new names, left to right
  const renamed = new Map<string, string[]>()
  for (const t of tables) {
    const own = typed.get(t.name)
    if (!own) continue
    const names: (string | null)[] = t.cols.map((c, i) => (own.has(i) ? own.get(i)! : c))
    const done: string[] = []
    names.forEach((n, i) => {
      const others = [...done, ...names.slice(i + 1).filter((x): x is string => x !== null)]
      let name = n ?? freeColumnName(others, book.locale)
      if (done.some((d) => tableKey(d) === tableKey(name))) {
        let k = 2
        while (others.some((d) => tableKey(d) === tableKey(`${name}${k}`))) k++
        name = `${name}${k}`
      }
      done.push(name)
    })
    renamed.set(tableKey(t.name), done)
  }
  // references follow their columns' new names, read with the old ones
  const out = rewriteBookTables(book, (t, spec) => {
    const names = renamed.get(tableKey(t.name))
    if (!names || !spec.cols) return undefined
    const to = (c: string) => names[t.index.get(tableKey(c))!] ?? c
    const cols = spec.cols.map(to) as [string] | [string, string]
    return cols.some((c, i) => c !== spec.cols![i]) ? { cols } : undefined
  })
  const sheets = out.sheets.map((sheet, s) => {
    const cells: Record<string, CellInput> = { ...sheet.cells }
    for (const e of edits) {
      if (sheetKey(e.sheet) !== sheetKey(sheet.name)) continue
      if (e.input === null) delete cells[e.cell]
      else cells[e.cell] = e.input
    }
    const list = sheet.tables?.map((ti) => {
      const names = renamed.get(tableKey(ti.name))
      if (!names) return ti
      const t = tables.find((x) => x.s === s && tableKey(x.name) === tableKey(ti.name))!
      names.forEach((n, i) => {
        const addr = formatAddress(t.r1, t.c1 + i)
        const { value: _, formula: __, ...look } = cells[addr] ?? {}
        cells[addr] = { ...look, value: n }
      })
      return { ...ti, columns: ti.columns.map((c, i) => ({ ...c, name: names[i] })) }
    })
    return { ...sheet, cells, ...(list ? { tables: list } : {}) }
  })
  return { ...out, sheets }
}
