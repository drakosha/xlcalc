/**
 * A pivot case played on the engine alone, without Excel's books: the source rows of `data.ts` on `D`, a cache refreshed
 * from them, `PT1` at `P!A3` and each operation as the engine takes it (fields placed by their names, value fields
 * named as Excel names them, grouping and calculated fields and items through `editStructure`'s operations on the
 * cache, edits of the source and refreshes), each PivotTable laid out again after each. A phase ends at each `read`:
 * the cells of every PivotTable's area and the operations the engine refused. The prediction of a held-out set before
 * Excel runs it (`predict.ts --drive`); operations it does not take are refused as `unsupported`.
 */
import type { BookInput, CellInput, PivotCache, PivotDataField, PivotFunction, PivotShowAs, PivotTable } from '../../../src/contract.js'
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
import { placedValues } from '../../../src/structure/pivot-values.js'
import { pivotOp } from '../../../src/structure/pivot-ops.js'
import { parseRange } from '../../../src/structure/area.js'
import { pivotArea } from '../../../src/pivot/area.js'
import { formatAddress, parseAddress } from '../../../src/model.js'
import { patched } from '../../../src/structure/pivots.js'
import type { PivotCase, PivotOp } from './cases.js'
import { DATA, sourceBook, sourceRange } from './data.js'

const CLOCK = clockOptions({ now: '2026-10-30', seed: 0 })
/** The harness's names of Show Values As that the contract spells otherwise. */
const SHOW: Record<string, PivotShowAs> = { percentOfCol: 'percentOfColumn', percentOfParentCol: 'percentOfParentColumn' }
const valuesOf = (b: BookInput) => stateValues(createState(b, CLOCK))

/** A phase played: each PivotTable's area by its name, and the operations refused (`op <index>: <message>`). */
export interface DrivenPhase {
  pivots: Record<string, { area: string } | { error: string }>
  book: BookInput
  errors: string[]
}

const isoOf = (d: number) => new Date(Math.round((d - 25_569) * 86_400_000)).toISOString().slice(0, 19)

/** The book a case starts from: the source rows on `D`, a cache refreshed from them and `PT1` at `P!A3`. */
export function startBook(c: PivotCase): BookInput {
  const cols = DATA[c.data ?? 'sales']()
  const book = sourceBook(cols, c.source ?? 'range')
  const source = c.source === 'table' ? { type: 'table' as const, table: 'Sales' } : c.source === 'name' ? { type: 'name' as const, name: 'SalesData' } : { type: 'range' as const, sheet: 'D', range: sourceRange(cols) }
  const grid = sourceGrid(book, source, valuesOf(book))
  if (typeof grid === 'string') throw new Error(grid)
  const made = refreshCache({ id: 1, source, fields: [] }, [], grid)
  if (typeof made === 'string') throw new Error(made)
  return { ...book, pivotCaches: [made.cache], sheets: book.sheets.map((s) => (s.name === 'P' ? { ...s, pivotTables: [emptyTable(made.cache, 'PT1', 'A3')] } : s)) }
}

const emptyTable = (cache: PivotCache, name: string, at: string): PivotTable => ({ name, cache: cache.id, location: { range: at, firstHeaderRow: 1, firstDataRow: 1, firstDataColumn: 0 }, fields: cache.fields.map(() => ({})) })

/** Plays a case: its phases, each with the book at its end; a case from a saved book (`from`) starts from `saved`, that book as Excel saved it. */
export function drive(c: PivotCase, saved?: BookInput): DrivenPhase[] {
  if (c.file || (c.from && !saved)) throw new Error('a case from a file, or from a saved book without it, is not driven')
  let book = c.from ? saved! : startBook(c)
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
  /** A cell of a PivotTable's area (`P!A5`), which an edit there is an entry into. */
  const onPivot = (cell: string) => {
    const [sheet, at] = cell.split('!')
    const { row, col } = parseAddress(at)
    return (book.sheets.find((s) => s.name === sheet)?.pivotTables ?? []).some((t) => {
      const x = pivotArea(t)
      return x !== null && row >= x.r1 && row <= x.r2 && col >= x.c1 && col <= x.c2
    })
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
      // over the cache of PT1, as addPivotTable makes it; COM with Excel's alerts off answers its question to replace the cells
      const [sheet, cell] = op.at.split('!')
      return take(k, pivotOp(book, { op: 'addPivotTable', sheet, cell, name: op.create, source: { cache: book.pivotCaches![0].id }, replace: true }, valuesOf(book)))
    }
    if ('addFields' in op || 'cubeField' in op) {
      // placed as Orientation places them
      const add = 'addFields' in op ? op.addFields : { [op.to === 'row' ? 'rows' : 'columns']: [op.cubeField] }
      for (const axis of ['rows', 'columns'] as const) {
        const fs = (add[axis] ?? []).map(fieldOf)
        const other = axis === 'rows' ? 'columns' : 'rows'
        setTable(pt, (t) => placedValues({ ...t, [other]: (t[other] ?? []).filter((f) => !fs.includes(f as number)), [axis]: [...(t[axis] ?? []).filter((f) => !fs.includes(f as number)), ...fs] }))
      }
      return laid(k)
    }
    if ('cellFormat' in op || 'merge' in op || 'addTable' in op) {
      const [sheet, range] = ('cellFormat' in op ? op.cellFormat : 'merge' in op ? op.merge : op.addTable).split('!')
      book = {
        ...book,
        sheets: book.sheets.map((s) => {
          if (s.name !== sheet) return s
          if ('cellFormat' in op) return { ...s, cells: { ...s.cells, [range]: { ...s.cells[range], numberFormat: op.numberFormat } } }
          if ('merge' in op) return { ...s, merges: [...(s.merges ?? []), range] }
          const x = parseRange(range)!
          const columns = Array.from({ length: x.c2 - x.c1 + 1 }, (_, c) => ({ name: String(s.cells[formatAddress(x.r1, x.c1 + c)]?.value ?? `Column${c + 1}`) }))
          return { ...s, tables: [...(s.tables ?? []), { name: op.name, range, columns }] }
        }),
      }
      return
    }
    if ('rows' in op || 'columns' in op) {
      const axis = 'rows' in op ? 'rows' : 'columns'
      const names = 'rows' in op ? op.rows : (op as { columns: string[] }).columns
      setTable(pt, (t) => {
        const fs = names.map(fieldOf)
        const other = axis === 'rows' ? 'columns' : 'rows'
        return placedValues({ ...t, [other]: (t[other] ?? []).filter((f) => !fs.includes(f as number)), [axis]: [...(t[axis] ?? []).filter((f) => !fs.includes(f as number)), ...fs] })
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
      setTable(pt, (t) => placedValues({ ...t, values: [...(t.values ?? []), d] }))
      return laid(k)
    }
    if ('dropValue' in op) {
      setTable(pt, (t) => placedValues({ ...t, values: (t.values ?? []).filter((d) => d.name !== op.dropValue) }))
      return laid(k)
    }
    if ('valuesOn' in op) {
      // Σ Values moved as editPivotTable moves it by valuesOnRows and valuesPosition
      const set = { valuesOnRows: op.valuesOn === 'row' || null, valuesPosition: op.position === undefined ? undefined : op.position - 1 }
      setTable(pt, (t) => placedValues(patched(t, set), set))
      return laid(k)
    }
    if ('move' in op) {
      const f = fieldOf(op.move)
      setTable(pt, (t) => {
        const off = { ...t, rows: (t.rows ?? []).filter((x) => x !== f), columns: (t.columns ?? []).filter((x) => x !== f), pages: (t.pages ?? []).filter((p) => p.field !== f) }
        if (op.to === 'hidden') return placedValues(off)
        if (op.to === 'page') return placedValues({ ...off, pages: [...off.pages, { field: f }] })
        const axis = op.to === 'row' ? 'rows' : 'columns'
        const list = [...off[axis]]
        list.splice(op.position === undefined ? list.length : op.position - 1, 0, f)
        return placedValues({ ...off, [axis]: list })
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
    if ('table' in op && Object.keys(op.table).every((k) => k === 'customLists')) {
      setTable(pt, (t) => ({ ...t, customListSort: op.table.customLists }))
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
    if ('edit' in op && !onPivot(op.edit)) {
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
