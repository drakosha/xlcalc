/**
 * A pivot case played on the engine alone, without Excel's books: the source rows of `data.ts` on `D`, a cache refreshed
 * from them, `PT1` at `P!A3` and each operation as the engine takes it (fields placed by their names, value fields
 * named as Excel names them, grouping and calculated fields and items through `editStructure`'s operations on the
 * cache, edits of the source and refreshes), each PivotTable laid out again after each. A phase ends at each `read`:
 * the cells of every PivotTable's area and the operations the engine refused. The prediction of a held-out set before
 * Excel runs it (`predict.ts --drive`); operations it does not take are refused as `unsupported`.
 */
import type { BookInput, CellInput, PivotAxisField, PivotCache, PivotDataField, PivotFunction, PivotShowAs, PivotTable } from '../../../src/contract.js'
import { createState } from '../../../src/engine.js'
import { editPivotCache } from '../../../src/pivot/cache-ops.js'
import { wordsOf } from '../../../src/pivot/captions.js'
import { refreshPivots } from '../../../src/pivot/engine.js'
import { groupedItems } from '../../../src/pivot/items.js'
import type { PivotCacheOp } from '../../../src/pivot/hook.js'
import { refreshCache } from '../../../src/pivot/refresh.js'
import { sourceGrid } from '../../../src/pivot/source.js'
import { itemText } from '../../../src/pivot/values.js'
import { clockOptions } from '../../../src/result.js'
import type { Applied } from '../../../src/structure/geometry.js'
import { stateValues } from '../../../src/structure/values.js'
import type { PivotCase, PivotOp } from './cases.js'
import { DATA, sourceBook, sourceRange } from './data.js'

const CLOCK = clockOptions({ now: '2026-10-30', seed: 0 })
/** The harness's names of Show Values As that the contract spells otherwise. */
const SHOW: Record<string, PivotShowAs> = { percentOfCol: 'percentOfColumn', percentOfParentCol: 'percentOfParentColumn' }
const valuesOf = (b: BookInput) => stateValues(createState(b, CLOCK))

/** The settings of a `table` operation the engine takes, each set on the PivotTable. */
const TABLE_PROPS: Record<string, (t: PivotTable, v: never) => PivotTable> = {
  customLists: (t, v: boolean) => ({ ...t, customListSort: v }),
  pagesWrap: (t, v: number) => ({ ...t, pagesWrap: v }),
  pagesAcross: (t, v: boolean) => ({ ...t, pagesAcrossFirst: v }),
  rowGrand: (t, v: boolean) => ({ ...t, grandTotals: { ...t.grandTotals, rows: v } }),
  columnGrand: (t, v: boolean) => ({ ...t, grandTotals: { ...t.grandTotals, columns: v } }),
  classic: (t, v: boolean) => ({ ...t, classicLayout: v }),
}

/** A phase played: each PivotTable's area by its name, and the operations refused (`op <index>: <message>`). */
export interface DrivenPhase {
  pivots: Record<string, { area: string } | { error: string }>
  book: BookInput
  errors: string[]
}

const isoOf = (d: number) => new Date(Math.round((d - 25_569) * 86_400_000)).toISOString().slice(0, 19)

/** Plays a case: its phases, each with the book at its end. */
export function drive(c: PivotCase): DrivenPhase[] {
  if (c.from || c.file) throw new Error('a case from a saved book is not driven')
  const cols = DATA[c.data ?? 'sales']()
  let book = sourceBook(cols, c.source ?? 'range')
  const source = c.source === 'table' ? { type: 'table' as const, table: 'Sales' } : c.source === 'name' ? { type: 'name' as const, name: 'SalesData' } : { type: 'range' as const, sheet: 'D', range: sourceRange(cols) }
  const grid = sourceGrid(book, source, valuesOf(book))
  if (typeof grid === 'string') throw new Error(grid)
  const made = refreshCache({ id: 1, source, fields: [] }, [], grid)
  if (typeof made === 'string') throw new Error(made)
  const cache: PivotCache = made.cache
  const empty = (name: string, at: string): PivotTable => ({ name, cache: 1, location: { range: at, firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 0 }, fields: (book.pivotCaches?.[0] ?? cache).fields.map(() => ({})) })
  book = { ...book, pivotCaches: [cache], sheets: book.sheets.map((s) => (s.name === 'P' ? { ...s, pivotTables: [empty('PT1', 'A3')] } : s)) }
  const out: DrivenPhase[] = []
  let errors: string[] = []
  const tableOf = (name = 'PT1') => book.sheets.find((s) => s.name === 'P')!.pivotTables!.find((t) => t.name === name)!
  const fieldOf = (name: string) => book.pivotCaches![0].fields.findIndex((f) => f.name.toLowerCase() === name.toLowerCase())
  const setTable = (name: string, f: (t: PivotTable) => PivotTable) => {
    book = { ...book, sheets: book.sheets.map((s) => (s.name === 'P' ? { ...s, pivotTables: s.pivotTables!.map((t) => (t.name === name ? f(t) : t)) } : s)) }
  }
  const take = (k: number, r: Applied) => {
    if (r.ok) book = r.book
    else errors.push(`op ${k}: ${r.message}`)
  }
  const laid = (k: number, fromSource = false, labelSort?: number[]) => take(k, refreshPivots(book, labelSort ? [{ sheet: 'P', name: 'PT1' }] : undefined, valuesOf(book), 0, fromSource, 'spill', labelSort))
  const cacheOp = (k: number, op: PivotCacheOp) => {
    try {
      take(k, editPivotCache(book, op, valuesOf(book), 0))
    } catch (e) {
      errors.push(`op ${k}: ${(e as Error).message}`)
    }
  }
  /** The place of an item on a PivotTable's field by its label. */
  const itemAt = (t: PivotTable, f: number, label: string) => {
    const values = groupedItems(book.pivotCaches![0].fields[f]) ?? []
    return (t.fields[f].items ?? []).findIndex((x) => x.item !== undefined && itemText(values[x.item] ?? null).toLowerCase() === label.toLowerCase())
  }
  c.ops.forEach((op: PivotOp, k) => {
    const pt = op.pt ?? 'PT1'
    if ('read' in op) {
      out.push(phase(book, errors))
      errors = []
      return
    }
    if ('create' in op) {
      const at = op.at.split('!')[1]
      book = { ...book, sheets: book.sheets.map((s) => (s.name === 'P' ? { ...s, pivotTables: [...s.pivotTables!, empty(op.create, at)] } : s)) }
      return laid(k)
    }
    if ('rows' in op || 'columns' in op) {
      const axis = 'rows' in op ? 'rows' : 'columns'
      const names = 'rows' in op ? op.rows : (op as { columns: string[] }).columns
      setTable(pt, (t) => {
        const fs = names.map(fieldOf)
        const other = axis === 'rows' ? 'columns' : 'rows'
        return { ...t, [other]: (t[other] ?? []).filter((f) => !fs.includes(f as number)), [axis]: [...(t[axis] ?? []).filter((f) => !fs.includes(f as number)), ...fs] }
      })
      return laid(k)
    }
    if ('pages' in op) {
      const fs = op.pages.map(fieldOf)
      setTable(pt, t => {
        const pages = [...(t.pages ?? [])]
        for (const field of fs) if (!pages.some(p => p.field === field)) pages.unshift({ field })
        return { ...t, rows: t.rows?.filter(f => !fs.includes(f as number)), columns: t.columns?.filter(f => !fs.includes(f as number)), pages }
      })
      return laid(k)
    }
    if ('page' in op || 'hide' in op) {
      const f = fieldOf('page' in op ? op.page : op.hide)
      const all = 'page' in op && op.items === '(All)'
      const selected = all ? [] : (op.items as string[]).map(label => itemAt(tableOf(pt), f, label))
      if (selected.some(p => p < 0)) { errors.push(`op ${k}: unknown item`); return }
      setTable(pt, t => {
        const fields = t.fields.slice()
        if ('hide' in op || all || selected.length > 1)
          fields[f] = { ...fields[f], items: fields[f].items?.map((x, p) => ({ ...x, hidden: 'hide' in op ? x.hidden || selected.includes(p) : !all && !selected.includes(p) })) }
        const pages = 'page' in op ? t.pages?.map(p => p.field === f ? { field: f, ...(selected.length === 1 ? { item: selected[0] } : {}) } : p) : t.pages
        return { ...t, fields, ...(pages ? { pages } : {}) }
      })
      return laid(k)
    }
    if ('value' in op) {
      const f = fieldOf(op.value.field)
      const fn = (op.value.fn ?? 'sum') as PivotFunction
      if (fn !== 'sum' && (book.pivotCaches![0].calculatedItems ?? []).length && ['average', 'stdDev', 'stdDevp', 'var', 'varp'].includes(fn)) {
        errors.push(`op ${k}: Averages, standard deviations, and variances are not supported when a PivotTable report has calculated items.`)
        return
      }
      const v = op.value
      const d: PivotDataField = { field: f, name: v.caption ?? wordsOf(undefined).valueName(fn, book.pivotCaches![0].fields[f].name), ...(fn !== 'sum' ? { function: fn } : {}) }
      if (v.show && v.show !== 'normal') d.showAs = (SHOW[v.show] ?? v.show) as PivotShowAs
      if (v.base) d.baseField = fieldOf(v.base)
      if (v.item === '(previous)' || v.item === '(next)') d.baseItem = v.item === '(previous)' ? 'previous' : 'next'
      else if (v.item !== undefined && d.baseField !== undefined && itemAt(tableOf(pt), d.baseField, v.item) >= 0) d.baseItem = itemAt(tableOf(pt), d.baseField, v.item)
      setTable(pt, (t) => ({ ...t, values: [...(t.values ?? []), d] }))
      return laid(k)
    }
    if ('move' in op) {
      const f = fieldOf(op.move)
      setTable(pt, (t) => {
        const off = { ...t, rows: (t.rows ?? []).filter((x) => x !== f), columns: (t.columns ?? []).filter((x) => x !== f), pages: (t.pages ?? []).filter((p) => p.field !== f) }
        if (op.to === 'hidden') return off
        if (op.to === 'page') return { ...off, pages: [...off.pages, { field: f }] }
        const axis = op.to === 'row' ? 'rows' : 'columns'
        const list = [...off[axis]]
        list.splice(op.position === undefined ? list.length : op.position - 1, 0, f)
        return { ...off, [axis]: list }
      })
      return laid(k)
    }
    if ('valuesOn' in op) {
      // the value fields' own field (Σ Values) on rows or columns, at its place among the axis's fields
      setTable(pt, (t) => {
        const rows: PivotAxisField[] = (t.rows ?? []).filter((f) => f !== 'values')
        const columns: PivotAxisField[] = (t.columns ?? []).filter((f) => f !== 'values')
        const axis = op.valuesOn === 'row' ? rows : columns
        const at = op.position === undefined ? axis.length : op.position - 1
        if ((t.values?.length ?? 0) > 1) axis.splice(at, 0, 'values')
        return { ...t, rows, columns, valuesOnRows: op.valuesOn === 'row', ...(op.position === undefined ? {} : { valuesPosition: at }) }
      })
      return laid(k)
    }
    if ('layout' in op) {
      setTable(pt, (t) => ({ ...t, layout: op.layout }))
      return laid(k)
    }
    if ('sort' in op && op.order !== 'manual') {
      const f = fieldOf(op.sort)
      setTable(pt, (t) => {
        const v = op.by === undefined ? undefined : (t.values ?? []).findIndex((x) => x.name === op.by)
        if (v === -1) throw new Error(`Unknown value field ${op.by}`)
        const fields = t.fields.slice()
        fields[f] = { ...fields[f], sort: { order: op.order as 'ascending' | 'descending', ...(v === undefined ? {} : { by: { value: v } }) } }
        return { ...t, fields }
      })
      return laid(k, false, op.by === undefined ? [f] : undefined)
    }
    if ('table' in op && Object.keys(op.table).every((k) => k in TABLE_PROPS)) {
      setTable(pt, (t) => Object.entries(op.table).reduce((x, [k, v]) => TABLE_PROPS[k](x, v as never), t))
      return laid(k)
    }
    if ('group' in op) {
      const field = fieldOf(op.group)
      const date = (v: unknown) => (typeof v === 'object' && v !== null && 'd' in v ? isoOf((v as { d: number }).d) : undefined)
      const group = op.periods
        ? { by: 'dates' as const, periods: op.periods, ...(op.by !== undefined ? { days: op.by } : {}), ...(op.start !== undefined ? { start: date(op.start) } : {}), ...(op.end !== undefined ? { end: date(op.end) } : {}) }
        : { by: 'numbers' as const, interval: op.by ?? 1, ...(op.start !== undefined ? { start: op.start as number } : {}), ...(op.end !== undefined ? { end: op.end as number } : {}) }
      return cacheOp(k, { op: 'groupPivotField', sheet: 'P', pivot: pt, field, group })
    }
    if ('groupItems' in op) {
      const field = fieldOf(op.groupItems)
      const t = tableOf(pt)
      return cacheOp(k, { op: 'groupPivotField', sheet: 'P', pivot: pt, field, group: { by: 'items', items: op.items.map((x) => itemAt(t, field, x)) } })
    }
    if ('ungroup' in op) return cacheOp(k, { op: 'ungroupPivotField', sheet: 'P', pivot: pt, field: fieldOf(op.ungroup) })
    if ('calcField' in op) return cacheOp(k, { op: 'setPivotCalculatedField', sheet: 'P', pivot: pt, name: op.calcField, formula: op.formula })
    if ('calcItem' in op) return cacheOp(k, { op: 'setPivotCalculatedItem', sheet: 'P', pivot: pt, field: fieldOf(op.field), name: op.calcItem, formula: op.formula })
    if ('subtotals' in op) {
      const f = fieldOf(op.subtotals)
      setTable(pt, (t) => {
        const fields = t.fields.slice()
        const { subtotals: _, ...rest } = fields[f]
        fields[f] = op.fns === 'automatic' ? rest : { ...rest, subtotals: op.fns === 'none' ? [] : op.fns }
        return { ...t, fields }
      })
      return laid(k)
    }
    if ('edit' in op && !op.edit.startsWith('P!')) {
      const [sheet, cell] = op.edit.split('!')
      const to = op.to
      // the cell keeps its format, as Value2 leaves it
      const edited = (was: CellInput | undefined): CellInput => {
        const { value: _, formula: __, ...kept } = was ?? {}
        return to === null ? kept : typeof to === 'object' ? { ...kept, ...('formula' in to ? { formula: to.formula } : { value: to.typed }) } : { ...kept, value: to }
      }
      book = { ...book, sheets: book.sheets.map((s) => (s.name === sheet ? { ...s, cells: { ...s.cells, [cell]: edited(s.cells[cell]) } } : s)) }
      return
    }
    if ('cache' in op && Object.keys(op.cache).every((k) => k === 'missingItems')) {
      const retainItems = op.cache.missingItems === 'none' ? 0 : op.cache.missingItems === 'max' ? 32500 : undefined
      book = { ...book, pivotCaches: book.pivotCaches!.map((c) => ({ ...c, retainItems })) }
      return
    }
    if ('refresh' in op || 'refreshAll' in op) return laid(k, true)
    errors.push(`op ${k}: unsupported`)
  })
  out.push(phase(book, errors))
  return out
}

function phase(book: BookInput, errors: string[]): DrivenPhase {
  const pivots = Object.fromEntries((book.sheets.find((s) => s.name === 'P')!.pivotTables ?? []).map((t) => [t.name, { area: t.location.range }]))
  return { pivots, book, errors }
}
