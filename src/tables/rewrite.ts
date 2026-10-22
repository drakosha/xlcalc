/**
 * Structured references rewritten in formula text, in place, as tables and their columns change: a table renamed, a
 * column renamed or deleted, a table deleted (#REF!) or turned into cells (references to its cells). Only the references
 * that change are rewritten; a formula that does not parse is left as written.
 */
import type { BookInput, CellInput, TableInput } from '../contract.js'
import { spend } from '../exec/budget.js'
import { formatAddress, readAddress } from '../model.js'
import { FormulaError, parse, quoteSheet } from '../syntax.js'
import { eachToken } from '../syntax/links.js'
import { caseKey } from '../collate.js'
import { specText, type TableSpec } from '../syntax/table.js'
import { tableAt, tableKey, tablesOf, type Table, type Tables } from './model.js'
import { mapFormulas } from '../cond/structure.js'
import { originOf } from '../cond/rules.js'
import { withFormulas } from '../validation/structure.js'
import { anchorOf } from '../validation/rules.js'

/** What becomes of a reference: unchanged, a new name for its table or columns, lost (#REF!), or other text. */
export type TableChange = undefined | { table?: string; cols?: [string] | [string, string] } | 'lost' | { text: string }

/** A change of the references to tables: `t` the table referred to (as it was), the cell's sheet and row (null in a name). */
export type TableEdit = (t: Table, spec: TableSpec, at: { s: number; r: number } | null) => TableChange

/**
 * The formula with its structured references changed by `edit`; `enclosing` is the table holding the formula's cell
 * (it gives the references written without a table name). `qualify` names the table in those too: the formula leaves it.
 */
export function rewriteTableRefs(text: string, tables: Tables, enclosing: Table | undefined, at: { s: number; r: number } | null, edit: TableEdit, qualify = false): string {
  if (!text.includes('[')) {
    const key = caseKey(text)
    if (!tables.list.some((t) => key.includes(caseKey(t.name)))) return text
  }
  try {
    parse(text, { union: at === null })
  } catch (e) {
    if (e instanceof FormulaError) return text
    throw e
  }
  return eachToken(text, (tok) => {
    let now: string | undefined
    if (tok.t === 'table') {
      const t = tok.table === null ? enclosing : tables.byName.get(tableKey(tok.table))
      if (!t) return undefined
      const ch = edit(t, tok.spec, at)
      const named = tok.table !== null || qualify
      if (ch === 'lost') now = '#REF!'
      else if (ch && 'text' in ch) now = ch.text
      // a new name for the table leaves a reference written without it as it is
      else if ((ch && (ch.cols || named)) || (qualify && tok.table === null)) now = `${named ? (ch?.table ?? t.name) : ''}[${specText(ch?.cols ? { ...tok.spec, cols: ch.cols } : tok.spec)}]`
    } else if (tok.t === 'ident' && tok.sheet === undefined) {
      // a table's name alone is its data
      const t = tables.byName.get(tableKey(tok.v))
      const ch = t ? edit(t, { areas: [] }, at) : undefined
      if (ch === 'lost') now = '#REF!'
      else if (ch && 'text' in ch) now = ch.text
      else if (ch?.table) now = ch.table
    }
    return now
  })
}

/**
 * Every formula of the book rewritten by `edit` (cells, names, the formulas of calculated columns, of conditional formats
 * and of data validation), with the tables of `book` itself; the same objects where nothing changed.
 */
export function rewriteBookTables(book: BookInput, edit: TableEdit): BookInput {
  const tables = tablesOf(book)
  let changed = false
  const p = { r: 0, c: 0 }
  const sheets = book.sheets.map((sheet, s) => {
    let cells: Record<string, CellInput> | undefined
    for (const addr in sheet.cells) {
      const cell = sheet.cells[addr]
      // a session's operation stops here at its deadline or abort (`exec/budget.ts`)
      spend(1)
      if (cell.formula === undefined) continue
      readAddress(addr, p)
      const formula = rewriteTableRefs(cell.formula, tables, tableAt(tables, s, p.r, p.c), { s, r: p.r }, edit)
      if (formula !== cell.formula) (cells ??= { ...sheet.cells })[addr] = { ...cell, formula }
    }
    let list: TableInput[] | undefined
    sheet.tables?.forEach((ti, i) => {
      const t = tables.byName.get(tableKey(ti.name))!
      const columns = ti.columns.map((c) => {
        if (c.formula === undefined) return c
        const formula = rewriteTableRefs(c.formula, tables, t, { s, r: t.r1 + (t.headers ? 1 : 0) }, edit)
        return formula === c.formula ? c : { ...c, formula }
      })
      if (columns.some((c, j) => c !== ti.columns[j])) (list ??= [...sheet.tables!])[i] = { ...ti, columns }
    })
    // rules' formulas as formulas of their first cell
    const conditionalFormats = mapList(sheet.conditionalFormats, (rule) => {
      const o = originOf(rule)
      return changedFormulas(rule, mapFormulas, (text) => rewriteTableRefs(text, tables, tableAt(tables, s, o.r, o.c), { s, r: o.r }, edit))
    })
    const validations = mapList(sheet.validations, (v) => {
      const o = anchorOf(v)
      return withFormulas(v, (text) => rewriteTableRefs(text, tables, tableAt(tables, s, o.r, o.c), { s, r: o.r }, edit))
    })
    if (!cells && !list && conditionalFormats === sheet.conditionalFormats && validations === sheet.validations) return sheet
    changed = true
    return { ...sheet, ...(cells ? { cells } : {}), ...(list ? { tables: list } : {}), ...(conditionalFormats ? { conditionalFormats } : {}), ...(validations ? { validations } : {}) }
  })
  let names = book.names
  const nm = book.names?.map((n) => {
    const ref = rewriteTableRefs(n.ref, tables, undefined, null, edit)
    return ref === n.ref ? n : { ...n, ref }
  })
  if (nm && nm.some((n, i) => n !== book.names![i])) {
    names = nm
    changed = true
  }
  return changed ? { ...book, sheets, ...(names ? { names } : {}) } : book
}

/** Each item through `f`; the same array when none changes. */
function mapList<T>(xs: T[] | undefined, f: (x: T) => T): T[] | undefined {
  const out = xs?.map(f)
  return out && out.some((x, i) => x !== xs![i]) ? out : xs
}

/** The item with its formulas through `f` by `map`; the same item when no formula changes. */
function changedFormulas<T>(x: T, map: (x: T, f: (text: string) => string) => T, f: (text: string) => string): T {
  let hit = false
  const next = map(x, (text) => {
    const out = f(text)
    hit ||= out !== text
    return out
  })
  return hit ? next : x
}

const abs = (r: number, c: number) => formatAddress(r, c).replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2')

/**
 * The text a reference becomes when its table turns into cells (Excel's Convert to Range): its cells with `$` and the
 * sheet, the formula's row kept relative (`'S'!$B4`; `row` null in a name, relative to A1); a part the table lacks (a
 * totals row) gives its whole range, as Excel converts it, and a column it lacks `#REF!`.
 */
export function convertedText(t: Table, spec: TableSpec, row: number | null, sheet: string): string {
  // a column the table lacks (the formula is #NAME?) is lost, so the formula stays an error
  if (spec.cols?.some((c) => !t.index.has(tableKey(c)))) return '#REF!'
  const head = t.headers ? t.r1 : -1
  const first = t.r1 + (t.headers ? 1 : 0)
  const last = t.r2 - (t.totals ? 1 : 0)
  const has = (a: TableSpec['areas'][number]) => spec.areas.includes(a)
  let r1 = first
  let r2 = last
  let relative = false
  // a part the table lacks (#REF!) becomes the whole table, columns and all
  if ((has('headers') && !has('data') && head < 0) || (has('totals') && !has('data') && !t.totals)) return `${quoteSheet(sheet)}!${abs(t.r1, t.c1)}:${abs(t.r2, t.c2)}`
  if (has('all')) [r1, r2] = [t.r1, t.r2]
  else if (has('row')) {
    ;[r1, r2] = [row ?? 0, row ?? 0]
    relative = true
  } else if (has('headers') && !has('data')) r1 = r2 = head
  else if (has('totals') && !has('data')) r1 = r2 = t.r2
  else {
    if (has('headers') && head >= 0) r1 = head
    if (has('totals') && t.totals) r2 = t.r2
  }
  let [c1, c2] = [t.c1, t.c2]
  // in a name, Excel turns the formula's row of some columns into the whole row of the table (the set `ops-held`)
  if (spec.cols && !(row === null && has('row'))) {
    const [a, b = a] = spec.cols.map((c) => t.index.get(tableKey(c))!)
    ;[c1, c2] = [t.c1 + Math.min(a, b), t.c1 + Math.max(a, b)]
  }
  const cell = (r: number, c: number) => {
    const [, col, num] = /^([A-Z]+)(\d+)$/.exec(formatAddress(r, c))!
    return `$${col}${relative ? '' : '$'}${num}`
  }
  const ref = r1 === r2 && c1 === c2 ? cell(r1, c1) : `${cell(r1, c1)}:${cell(r2, c2)}`
  return `${quoteSheet(sheet)}!${ref}`
}
