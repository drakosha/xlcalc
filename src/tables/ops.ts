/**
 * Operations on a book's tables (`editStructure`): a table added over cells, renamed, deleted with its cells, or turned
 * into plain cells, with the structured references to it rewritten as Excel rewrites them (the set `tables/ops`).
 */
import type { BookInput, StructureOp, TableInput } from '../contract.js'
import { checkTables } from './check.js'
import { InputError } from '../input.js'
import { formatAddress } from '../model.js'
import { say, type Message } from '../locale/messages.js'
import { localeBooleans } from '../locale/registry.js'
import { nameRule } from '../syntax.js'
import { meets, parseRange, rectCells } from '../structure/area.js'
import type { Applied } from '../structure/geometry.js'
import { parseMerge } from '../layout/merges.js'
import { convertedText, rewriteBookTables } from './rewrite.js'
import { sheetIndex } from '../structure/rewrite.js'
import { tableKey, tablesOf } from './model.js'

type TableOp = Extract<StructureOp, { op: 'addTable' | 'renameTable' | 'deleteTable' | 'convertTable' }>

/** The book with the table of key `key` replaced by `next` (or removed when null) on its sheet. */
function withTable(book: BookInput, key: string, next: TableInput | null): BookInput {
  return {
    ...book,
    sheets: book.sheets.map((s) => {
      if (!s.tables?.some((t) => tableKey(t.name) === key)) return s
      const tables = s.tables.flatMap((t) => (tableKey(t.name) !== key ? [t] : next ? [next] : []))
      const out = { ...s, tables }
      if (!tables.length) delete (out as { tables?: TableInput[] }).tables
      return out
    }),
  }
}

/** A name a table can take in the book: Excel's rules, not another table's or a defined name (but `own`). */
function nameRefusal(book: BookInput, name: string, own?: string): Message | null {
  // not ИСТИНА in ru-RU, as Excel for Windows (Excel for Mac keeps such a table's name)
  const rule = nameRule(name, localeBooleans(book.locale))
  if (rule) return { key: 'tableName', name, rule }
  const k = tableKey(name)
  const table = tablesOf(book).list.find((t) => tableKey(t.name) === k && (own === undefined || tableKey(own) !== k))
  if (table) return { key: 'tableNameTaken', table: table.name }
  if ((book.names ?? []).some((n) => tableKey(n.name) === k)) return { key: 'tableNameDefined', name }
  return null
}

export function tableOp(book: BookInput, op: TableOp): Applied {
  if (op.op === 'addTable') return addTable(book, op.sheet, op.table)
  const t = tablesOf(book).byName.get(tableKey(op.table))
  if (!t) throw new InputError('op.table', `unknown table ${JSON.stringify(op.table)}`)
  const key = tableKey(t.name)
  if (op.op === 'renameTable') {
    if (typeof op.name !== 'string') throw new InputError('op.name', 'expected a table name')
    const why = nameRefusal(book, op.name, t.name)
    if (why) return { ok: false, code: 'table_name', message: say(book.locale, why) }
    // Excel keeps the name when only its case changes (the set `tables/ops`, rename-case)
    if (tableKey(op.name) === key) return { ok: true, book, geometry: [] }
    const renamed = rewriteBookTables(book, (x) => (tableKey(x.name) === key ? { table: op.name } : undefined))
    return { ok: true, book: withTable(renamed, key, { ...t.input, name: op.name }), geometry: [] }
  }
  if (op.op === 'deleteTable') {
    // the table goes with its cells
    const out = withTable(rewriteBookTables(book, (x) => (tableKey(x.name) === key ? 'lost' : undefined)), key, null)
    const cells = { ...out.sheets[t.s].cells }
    for (const a of rectCells(t)) delete cells[a]
    const sheets = [...out.sheets]
    sheets[t.s] = { ...sheets[t.s], cells }
    return { ok: true, book: { ...out, sheets }, geometry: [{ kind: 'clear', s: t.s, area: { r1: t.r1, r2: t.r2, c1: t.c1, c2: t.c2 } }] }
  }
  // convertTable: references become the cells they meant
  const sheet = book.sheets[t.s].name
  const out = rewriteBookTables(book, (x, spec, at) => (tableKey(x.name) === key ? { text: convertedText(x, spec, at ? at.r : null, sheet) } : undefined))
  return { ok: true, book: withTable(out, key, null), geometry: [] }
}

/** A table over cells of `sheet`: its header cells take the columns' names; refused over another table or a merged area. */
function addTable(book: BookInput, sheet: string, table: TableInput): Applied {
  const s = sheetIndex(book, sheet)
  if (s < 0) throw new InputError('op.sheet', `unknown sheet ${JSON.stringify(sheet)}`)
  if (typeof table !== 'object' || table === null || typeof table.name !== 'string') throw new InputError('op.table', 'expected a table such as { name: "Sales", range: "A1:C9", columns: [...] }')
  const why = nameRefusal(book, table.name)
  if (why) return { ok: false, code: 'table_name', message: say(book.locale, why) }
  const x = typeof table.range === 'string' ? parseRange(table.range) : null
  const over = x && tablesOf(book).list.find((t) => t.s === s && meets(t, x))
  if (over) return { ok: false, code: 'table', message: say(book.locale, { key: 'tableOverTable', range: table.range, table: over.name }) }
  const merge = x && (book.sheets[s].merges ?? []).find((m) => meets(parseMerge(m)!, x))
  if (merge) return { ok: false, code: 'merged', message: say(book.locale, { key: 'tableOverMerge', range: table.range, area: merge }) }
  const cells = { ...book.sheets[s].cells }
  if (x && table.headers !== false && Array.isArray(table.columns))
    table.columns.forEach((c, i) => {
      const addr = formatAddress(x.r1, x.c1 + i)
      if (c && typeof c.name === 'string') cells[addr] = { ...cells[addr], value: c.name, formula: undefined }
      if (cells[addr].formula === undefined) delete cells[addr].formula
    })
  // a malformed table (its fields, range, columns) is an input error, as in a book
  checkTables({ ...(book.names ? { names: book.names } : {}), sheets: [{ ...book.sheets[s], cells, tables: [table] }] }, 'op')
  const sheets = [...book.sheets]
  sheets[s] = { ...sheets[s], cells, tables: [...(sheets[s].tables ?? []), table] }
  return { ok: true, book: { ...book, sheets }, geometry: [] }
}
