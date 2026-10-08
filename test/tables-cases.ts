/**
 * A table case of `bench/coverage/tables/` taken by the engine as Excel took it: the book opened in a session, each
 * step applied (a cell typed through `setCells`, an operation through `restructure`, a refusal noted), and the answer
 * read as the harness reads Excel's: the rectangles' values, their formulas as stored text, the tables.
 */
import type { Val } from '../bench/coverage/formulajs.js'
import { cellsOf } from '../bench/coverage/case-book.js'
import { compare } from '../bench/coverage/compare.js'
import type { TableAnswer, TableCase, TableShape } from '../bench/coverage/tables/cases.js'
import { InputError, openSession, type BookInput, type ComputeResult } from '../src/index.js'
import { parseAddress } from '../src/model.js'
import { sheetKey, validate } from '../src/syntax.js'
import { tableAt, tablesOf } from '../src/tables/model.js'
import { shownValue } from './spill-shown.js'

export const CLOCK = { now: '2026-10-07T12:00:00Z', seed: 1 }

export interface EngineAnswer {
  values: Record<string, Val>
  /** Each formula of the read rectangles, as stored (canonical) text. */
  formulas: Record<string, string>
  tables: TableShape[]
  /** The steps the engine refused, by index: its message. */
  refused: Record<number, string>
  book: BookInput
  result: ComputeResult
}

/** The tables of a book as Excel lists them. */
export const shapes = (book: BookInput): TableShape[] =>
  book.sheets.flatMap((s) => (s.tables ?? []).map((t) => ({ sheet: s.name, name: t.name, range: t.range.includes(':') ? t.range : `${t.range}:${t.range}`, headers: t.headers !== false, totals: t.totals === true, columns: t.columns.map((c) => c.name) })))

export function engineAnswer(c: TableCase): EngineAnswer {
  const s = openSession(c.book, { clock: CLOCK })
  const refused: Record<number, string> = {}
  ;(c.steps ?? []).forEach((step, i) => {
    if ('op' in step) {
      const r = s.restructure(step.op)
      if (!r.ok) refused[i] = `${r.code}: ${r.message}`
      return
    }
    const { sheet, cell, value: v } = step.edit
    const input = v === null ? null : typeof v === 'object' ? { formula: v.formula } : { value: v }
    try {
      s.setCells(cellsOf(cell).addrs.map((a) => ({ sheet, cell: a, input })))
    } catch (e) {
      if (!(e instanceof InputError)) throw e
      refused[i] = e.message
    }
  })
  const book = s.book()
  const result = s.result()
  const values: Record<string, Val> = {}
  const formulas: Record<string, string> = {}
  for (const g of c.read) {
    const sheet = result.sheets.find((x) => sheetKey(x.name) === sheetKey(g.sheet))
    if (!sheet) continue
    for (const a of cellsOf(g.range).addrs) {
      const v = shownValue(result, sheet.name, a)
      if (v) values[`${g.sheet}!${a}`] = v
      const f = sheet.cells[a]?.formula
      if (f !== undefined) formulas[`${g.sheet}!${a}`] = f
    }
  }
  s.close()
  return { values, formulas, tables: shapes(book), refused, book, result }
}

/** Excel's formula text (`=[@Qty]*2`, as Excel shows it) at a cell of the book as stored text, read with the book's tables. */
export function storedText(book: BookInput, sheet: string, addr: string, shown: string): string {
  const tables = tablesOf(book)
  const s = book.sheets.findIndex((x) => sheetKey(x.name) === sheetKey(sheet))
  const { row, col } = parseAddress(addr)
  const enclosing = tableAt(tables, s, row, col)?.name
  const r = validate(shown.replace(/^=/, ''), { sheets: book.sheets.map((x) => x.name), tables: tables.list.map((t) => ({ name: t.name, columns: t.cols })), ...(enclosing ? { enclosing } : {}) })
  return r.ok ? r.canonical : `(${r.code}) ${shown}`
}

/** What differs between the engine's answer and Excel's, as lines (empty when they agree). */
export function differences(c: TableCase, x: TableAnswer): string[] {
  const out: string[] = []
  if (x.rejected) return ['Excel did not load the book']
  const got = engineAnswer(c)
  // Excel refuses some steps without an error (rows deleted from two tables): the book stays as it was, which the
  // values, formulas and tables below check; every step it answered with an error the engine refuses too
  const missed = Object.keys(x.refused ?? {}).filter((i) => !(i in got.refused))
  if (missed.length) out.push(`steps Excel refused and the engine took: ${missed.map((i) => `${i} (${x.refused![Number(i)]})`).join(', ')}`)
  const shape = (t: TableAnswer['tables'][number]) => JSON.stringify(t)
  const ex = x.tables.map(shape).sort()
  const en = got.tables.map(shape).sort()
  if (JSON.stringify(ex) !== JSON.stringify(en)) out.push(`tables: Excel ${ex.join(' ')}, engine ${en.join(' ')}`)
  for (const k of new Set([...Object.keys(x.values), ...Object.keys(got.values)])) {
    const want = x.values[k]
    const have = got.values[k]
    if (!want || !have) {
      if (want !== have) out.push(`${k}: Excel ${JSON.stringify(want ?? 'blank')}, engine ${JSON.stringify(have ?? 'blank')}`)
    } else if (!compare(want, have).ok) out.push(`${k}: Excel ${JSON.stringify(want)}, engine ${JSON.stringify(have)}`)
  }
  for (const [k, shown] of Object.entries(x.formulas)) {
    const [sheet, addr] = k.split('!')
    const stored = storedText(got.book, sheet, addr, shown)
    if (stored !== got.formulas[k]) out.push(`${k}: Excel ${shown} (${stored}), engine ${got.formulas[k] ?? 'no formula'}`)
  }
  return out
}
