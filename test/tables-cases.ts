/**
 * A table case of `bench/coverage/tables/` taken by the engine as Excel took it: the book opened in a session, each
 * step applied (a cell typed through `setCells`, an operation through `editStructure`, a name added through `setNames`,
 * a refusal noted), and the answer read as the harness reads Excel's: the rectangles' values, their formulas as stored
 * text, the tables, the names.
 */
import type { Val } from '../bench/coverage/formulajs.js'
import { cellsOf } from '../bench/coverage/case-book.js'
import { compare } from '../bench/coverage/compare.js'
import type { TableAnswer, TableCase, TableShape } from '../bench/coverage/tables/cases.js'
import { compute, enUS, InputError, LOCALES, createSession, quoteSheetIfNeeded, toInvariant, toLocal, type BookInput, type ComputeResult, type Locale } from '../src/index.js'
import { parseAddress } from '../src/model.js'
import { sheetKey, validate } from '../src/syntax.js'
import { tableAt, tablesOf } from '../src/tables/model.js'
import { shownValue } from './spill-shown.js'
import { nameIssues } from './diagnostics.js'

export const CLOCK = { now: '2026-10-07T12:00:00Z', seed: 1 }

export interface EngineAnswer {
  values: Record<string, Val>
  /** Each formula of the read rectangles, as stored (canonical) text. */
  formulas: Record<string, string>
  tables: TableShape[]
  /** The steps the engine refused, by index: its message. */
  refused: Record<number, string>
  /** The names the book computes with, a sheet's own as `Sheet!name`, as Excel lists them. */
  names: string[]
  book: BookInput
  result: ComputeResult
}

/** The tables of a book as Excel lists them. */
export const shapes = (book: BookInput): TableShape[] =>
  book.sheets.flatMap((s) => (s.tables ?? []).map((t) => ({ sheet: s.name, name: t.name, range: t.range.includes(':') ? t.range : `${t.range}:${t.range}`, headers: t.headers !== false, totals: t.totals === true, columns: t.columns.map((c) => c.name) })))

/** The formula text of a case's locale, in which its typed formulas (`{ local }`) and Excel's local text are; a book without one, the run's (`tag`). */
const localeOf = (c: TableCase, tag = 'en-US') => LOCALES.get(c.book.locale ?? tag)!

export function engineAnswer(c: TableCase): EngineAnswer {
  const s = createSession(c.book, { clock: CLOCK })
  const refused: Record<number, string> = {}
  ;(c.steps ?? []).forEach((step, i) => {
    if ('name' in step) {
      // Excel's Names.Add takes a name or refuses it: a name left out with an issue is refused
      const before = s.getBook().names ?? []
      try {
        s.setNames([...before, step.name])
      } catch (e) {
        if (!(e instanceof InputError)) throw e
        refused[i] = e.message
        return
      }
      const issue = nameIssues(s.getResult()).find((x) => x.name === step.name.name && x.scope === step.name.scope)
      if (issue) {
        refused[i] = issue.reason
        s.setNames(before)
      }
      return
    }
    if ('op' in step) {
      const r = s.editStructure(step.op)
      if (!r.ok) refused[i] = `${r.code}: ${r.message}`
      return
    }
    const { sheet, cell, value: v } = step.edit
    const local = v !== null && typeof v === 'object' && 'local' in v ? toInvariant(v.local, localeOf(c)) : null
    if (local && !local.ok) throw new Error(`${c.id}: ${local.message}`)
    const input = v === null ? null : typeof v === 'object' ? { formula: local?.ok ? local.formula : (v as { formula: string }).formula } : { value: v }
    // a cell typed over keeps its style and number format, as in Excel
    const look = (a: string) => {
      const { value: _, formula: __, ...rest } = s.getBook().sheets.find((x) => sheetKey(x.name) === sheetKey(sheet))?.cells[a] ?? {}
      return rest
    }
    try {
      s.setCells(cellsOf(cell).addrs.map((a) => ({ sheet, cell: a, input: input && { ...look(a), ...input } })))
    } catch (e) {
      if (!(e instanceof InputError)) throw e
      refused[i] = e.message
    }
  })
  const book = s.getBook()
  const result = s.getResult()
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
  const left = new Set(nameIssues(result).map((x) => `${x.scope ?? ''}!${x.name}`))
  const names = (book.names ?? []).filter((n) => !left.has(`${n.scope ?? ''}!${n.name}`)).map((n) => (n.scope ? `${n.scope}!${n.name}` : n.name))
  return { values, formulas, tables: shapes(book), refused, names, book, result }
}

/** The table holding a cell of the book. */
function enclosingOf(book: BookInput, sheet: string, addr: string): string | undefined {
  const s = book.sheets.findIndex((x) => sheetKey(x.name) === sheetKey(sheet))
  const { row, col } = parseAddress(addr)
  return tableAt(tablesOf(book), s, row, col)?.name
}

/** Excel's formula text (`=[@Qty]*2`, as Excel shows it) at a cell of the book as stored text, read with the book's tables. */
export function storedText(book: BookInput, sheet: string, addr: string, shown: string): string {
  const tables = tablesOf(book)
  const s = book.sheets.findIndex((x) => sheetKey(x.name) === sheetKey(sheet))
  const { row, col } = parseAddress(addr)
  const enclosing = tableAt(tables, s, row, col)?.name
  const r = validate(shown.replace(/^=/, ''), { sheets: book.sheets.map((x) => x.name), tables: tables.list.map((t) => ({ name: t.name, columns: t.cols })), ...(enclosing ? { table: enclosing } : {}) })
  return r.ok ? r.canonical : `(${r.code}) ${shown}`
}

/** What differs between the engine's answer and Excel's in a run in the locale `tag`, as lines (empty when they agree). */
export function differences(c: TableCase, x: TableAnswer, tag?: string): string[] {
  const out: string[] = []
  // Excel did not load the book over a name it does not take: the engine agrees when it leaves a name of the book out
  if (x.rejected) return nameIssues(compute(c.book, { display: false })).length ? [] : ['Excel did not load the book']
  const got = engineAnswer(c)
  // Excel refuses some steps without an error (rows deleted from two tables): the book stays as it was, which the
  // values, formulas and tables below check; every step it answered with an error the engine refuses too
  const missed = Object.keys(x.refused ?? {}).filter((i) => !(i in got.refused))
  if (missed.length) out.push(`steps Excel refused and the engine took: ${missed.map((i) => `${i} (${x.refused![Number(i)]})`).join(', ')}`)
  const shape = (t: TableAnswer['tables'][number]) => JSON.stringify(t)
  const ex = x.tables.map(shape).sort()
  const en = got.tables.map(shape).sort()
  if (JSON.stringify(ex) !== JSON.stringify(en)) out.push(`tables: Excel ${ex.join(' ')}, engine ${en.join(' ')}`)
  // the names in any order (Excel lists them sorted), a sheet's name unquoted
  const nameList = (xs: string[]) => xs.map((n) => n.replace(/^'(.*)'!/, '$1!')).sort()
  if (x.names && JSON.stringify(nameList(x.names)) !== JSON.stringify(nameList(got.names))) out.push(`names: Excel ${x.names.join(' ')}, engine ${got.names.join(' ')}`)
  for (const k of new Set([...Object.keys(x.values), ...Object.keys(got.values)])) {
    const want = x.values[k]
    const have = got.values[k]
    if (!want || !have) {
      if (want !== have) out.push(`${k}: Excel ${JSON.stringify(want ?? 'blank')}, engine ${JSON.stringify(have ?? 'blank')}`)
    } else if (!compare(want, have).ok) out.push(`${k}: Excel ${JSON.stringify(want)}, engine ${JSON.stringify(have)}`)
  }
  for (const [k, typed] of Object.entries(x.formulas)) {
    let shown = typed
    const [sheet, addr] = k.split('!')
    const stored = storedText(got.book, sheet, addr, shown)
    if (stored !== got.formulas[k]) out.push(`${k}: Excel ${shown} (${stored}), engine ${got.formulas[k] ?? 'no formula'}`)
    // the text Excel shows, en-US and in the user's language: toLocal with the table of the cell
    if (got.formulas[k] === undefined) continue
    const table = enclosingOf(got.book, sheet, addr)
    // (the stored text quotes every sheet name, Excel only the names that need it; a reference across sheets stays quoted)
    const unquoted = (f: string, locale: Locale) => f.replace(/'((?:[^':]|'')+)'!/g, (_, name: string) => `${quoteSheetIfNeeded(name.replaceAll("''", "'"), locale)}!`)
    // nor the spaces typed between a reference's items (`[[#Data], [Price]]`)
    shown = shown.replace(/([,;]) +\[/g, '$1[')
    const en = toLocal(got.formulas[k], enUS, { table })
    if (en.ok && `=${unquoted(en.formula, enUS)}` !== shown) out.push(`${k}: Excel shows ${shown}, toLocal ${en.formula}`)
    const local = x.local?.[k]
    const loc = local === undefined ? null : toLocal(got.formulas[k], localeOf(c, tag), { table })
    if (loc?.ok && `=${unquoted(loc.formula, localeOf(c, tag))}` !== local?.replace(/([,;]) +\[/g, '$1[')) out.push(`${k}: Excel ${c.book.locale ?? tag} shows ${local}, toLocal ${loc.formula}`)
  }
  return out
}
