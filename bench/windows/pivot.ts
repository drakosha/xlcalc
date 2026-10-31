/**
 * Pivot tables in Excel for Windows (`bench/coverage/pivot/`): each case's book (the rows of `data.ts`, or a saved
 * book of another set) opened and computed; `pivotSteps` turns the case's description into steps of `pivot.ps1`, one
 * primitive of the object model each, and Excel builds the pivot tables, reads them at the end of each phase and saves
 * the book. The answers go to `test/excel/win/pivot/[<tag>/]<set>.json`, the books Excel saved beside them
 * (`<set>/<id>.xlsx`, and `<set>/<id>-<k>.xlsx` for the copy saved at the end of phase k before the last).
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import type { CellValue } from '../../src/contract.js'
import { writeXlsx } from '../../src/xlsx.js'
import { FILTER_TYPES, type Agg, type Axis, type FieldProps, type PivotAnswer, type PivotArea, type PivotCase, type PivotOp, type PivotPhase, type PivotSet, type Period, type ShowAs, type TableProps } from '../coverage/pivot/cases.js'
import { cellsOf } from '../coverage/case-book.js'
import { DATA, sourceBook, sourceText } from '../coverage/pivot/data.js'
import { localeDir } from '../coverage/locales.js'
import { fileAnswers, secondsSince, setModule } from '../coverage/stored.js'
import { caseFiles } from '../coverage/xlsx/sent.js'
import { comEdit, fileJobs } from './books.js'
import { valOf, winFile, type WinCell } from './functions.js'
import { metaOf, type Step, type Tag, type WinEnv } from './remote.js'

const ROOT = fileURLToPath(new URL('../..', import.meta.url))

/** `XlConsolidationFunction`. */
const FN: Record<Agg, number> = { sum: -4157, count: -4112, average: -4106, max: -4136, min: -4139, product: -4149, countNums: -4113, stdDev: -4155, stdDevp: -4156, var: -4164, varp: -4165 }
/** The index of each function in `PivotField.Subtotals` (1 is Automatic). */
const SUBTOTAL: Record<Agg, number> = { sum: 2, count: 3, average: 4, max: 5, min: 6, product: 7, countNums: 8, stdDev: 9, stdDevp: 10, var: 11, varp: 12 }
/** `XlPivotFieldCalculation`. */
const SHOW: Record<ShowAs, number> = { normal: -4143, difference: 2, percent: 3, percentDiff: 4, runTotal: 5, percentOfRow: 6, percentOfCol: 7, percentOfTotal: 8, index: 9, percentOfParentRow: 10, percentOfParentCol: 11, percentOfParent: 12, percentOfRunningTotal: 13, rankAscending: 14, rankDescending: 15 }
/** `XlPivotFieldOrientation`. */
const AXIS: Record<Axis, number> = { hidden: 0, row: 1, column: 2, page: 3 }
/** The periods in the order of `Range.Group`'s array. */
const PERIODS: readonly Period[] = ['seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years']
/** `XlTimelineLevel`, `XlPivotTableMissingItems`, `XlLayoutRowType`. */
const LEVEL = { years: 0, quarters: 1, months: 2, days: 3 } as const
const MISSING = { default: -1, none: 0, max: 32500 } as const
const ROW_LAYOUT = { compact: 0, tabular: 1, outline: 2 } as const
/** `XlSlicerSort`, `XlSlicerCrossFilterType`. */
const SLICER_SORT = { source: 1, ascending: 2, descending: 3 } as const
const CROSS_FILTER = { none: 1, last: 2, dimmed: 3, hidden: 4 } as const
/** `XlSortOrder`, and `xlManual` for no sorting on update. */
const SORT_ORDER = { ascending: 1, descending: 2, manual: -4135 } as const
/** The pivot table's and the cache's version, 8 as Excel 365 makes them (COM refuses `xlPivotTableVersionCurrent`, -1). */
const VERSION = 8

type Props = [string, unknown][]

/** A table's settings as the object model names them. */
function tableProps(t: TableProps): Props {
  const p: Props = []
  if (t.rowGrand !== undefined) p.push(['RowGrand', t.rowGrand])
  if (t.columnGrand !== undefined) p.push(['ColumnGrand', t.columnGrand])
  if (t.grandTotalName !== undefined) p.push(['GrandTotalName', t.grandTotalName])
  if (t.rowHeader !== undefined) p.push(['CompactLayoutRowHeader', t.rowHeader])
  if (t.columnHeader !== undefined) p.push(['CompactLayoutColumnHeader', t.columnHeader])
  if (t.errorText !== undefined) p.push(['DisplayErrorString', true], ['ErrorString', t.errorText])
  if (t.emptyText !== undefined) p.push(['DisplayNullString', true], ['NullString', t.emptyText])
  if (t.multipleFilters !== undefined) p.push(['AllowMultipleFilters', t.multipleFilters])
  if (t.mergeLabels !== undefined) p.push(['MergeLabels', t.mergeLabels])
  if (t.customLists !== undefined) p.push(['SortUsingCustomLists', t.customLists])
  if (t.valuesRow !== undefined) p.push(['ShowValuesRow', t.valuesRow])
  if (t.classic !== undefined) p.push(['InGridDropZones', t.classic])
  return p
}

/** A field's settings as the object model names them; the form is `LayoutCompactRow` with `LayoutForm` (`xlTabular` 0, `xlOutline` 1). */
function fieldProps(f: FieldProps): Props {
  const p: Props = []
  if (f.caption !== undefined) p.push(['Caption', f.caption])
  if (f.form) p.push(['LayoutCompactRow', f.form === 'compact'], ['LayoutForm', f.form === 'tabular' ? 0 : 1])
  if (f.blankLine !== undefined) p.push(['LayoutBlankLine', f.blankLine])
  if (f.subtotalsTop !== undefined) p.push(['LayoutSubtotalLocation', f.subtotalsTop ? 1 : 2])
  if (f.repeatLabels !== undefined) p.push(['RepeatLabels', f.repeatLabels])
  if (f.format !== undefined) p.push(['NumberFormat', f.format])
  if (f.showAllItems !== undefined) p.push(['ShowAllItems', f.showAllItems])
  return p
}

/** What each step of a case is for, to read its answer: an operation (by index, -1 the first pivot table), a read of a pivot table, a watched range, a slicer, the copy or the save. */
type Role = { op: number } | { read: string } | { readAll: true } | { watch: string } | { slicer: string } | { external: string[] } | { end: true }

/** A case's steps and what each is for. */
export interface Planned {
  steps: Step[]
  roles: Role[]
}

/**
 * The steps of `pivot.ps1` for a case: `PT1` made over the source unless the book has it, each operation, and at the
 * end of each phase the pivot tables read, the ranges watched, the slicers read and a copy saved (the book saved at the
 * end). `book` gives the file of a set's saved book (`basic/sum`) for an `external` operation.
 */
export function pivotSteps(c: PivotCase, book: (name: string) => Uint8Array): Planned {
  const steps: Step[] = [{ do: 'full' }]
  const roles: Role[] = [{ op: -2 }]
  const sheets = new Map<string, string>([['PT1', 'P']])
  const slicers: string[] = []
  const add = (role: Role, ...s: Step[]) => s.forEach((x) => (steps.push(x), roles.push(role)))
  let copies = 0
  const phase = (last: boolean) => {
    // formulas typed over the pivot tables settle in a full calculation
    add({ op: -2 }, { do: 'full' })
    if (c.file) add({ readAll: true }, { do: 'pivot', op: 'readAll' })
    else for (const [pt, sheet] of sheets) add({ read: pt }, { do: 'pivot', op: 'read', sheet, pt })
    for (const w of c.watch ?? []) {
      const [sheet, range] = w.split('!')
      add({ watch: w }, { do: 'values', sheet, range })
    }
    for (const name of slicers) add({ slicer: name }, { do: 'pivot', op: 'slicerRead', field: name })
    add({ end: true }, last ? { do: 'saveAs' } : { do: 'pivot', op: 'copy', k: ++copies })
  }
  if (!c.from && !c.file) {
    const cols = DATA[c.data ?? 'sales']()
    add({ op: -1 }, { do: 'pivot', op: 'create', sheet: 'P', at: 'A3', pt: 'PT1', source: sourceText(c.source ?? 'range', cols), version: VERSION })
  }
  c.ops.forEach((o, k) => {
    if ('read' in o) return phase(false)
    const pt = o.pt ?? 'PT1'
    const at = { do: 'pivot', sheet: sheets.get(pt) ?? 'P', pt }
    const set = (on: string, props: Props, more: object = {}): Step => ({ ...at, op: 'set', on, props, ...more })
    const call = (on: string, method: string, args: unknown[], more: object = {}): Step => ({ ...at, op: 'call', on, method, args, ...more })
    const role = { op: k }
    if ('create' in o) {
      const [sheet, cell] = o.at.split('!')
      sheets.set(o.create, sheet)
      const source = o.source ? { source: sourceText(o.source, DATA[c.data ?? 'sales']()) } : { cacheOf: 'PT1', cacheSheet: sheets.get('PT1') }
      add(role, { do: 'pivot', op: 'create', sheet, at: cell, pt: o.create, version: VERSION, ...source })
    } else if ('rows' in o || 'columns' in o || 'pages' in o) {
      const [axis, fields]: [Axis, string[]] = 'rows' in o ? ['row', o.rows] : 'columns' in o ? ['column', o.columns] : ['page', o.pages]
      for (const f of fields) add(role, set('field', [['Orientation', AXIS[axis]]], { field: f }))
    } else if ('move' in o) add(role, set('field', [['Orientation', AXIS[o.to]], ...(o.position ? [['Position', o.position] as [string, unknown]] : [])], { field: o.move }))
    else if ('value' in o) {
      const v = o.value
      const props: Props = []
      if (v.format !== undefined) props.push(['NumberFormat', v.format])
      if (v.show) props.push(['Calculation', SHOW[v.show]])
      if (v.base !== undefined) props.push(['BaseField', v.base])
      if (v.item !== undefined) props.push(['BaseItem', v.item])
      add(role, { ...at, op: 'data', field: v.field, caption: v.caption ?? null, fn: FN[v.fn ?? 'sum'], props })
    } else if ('valueDefault' in o) add(role, { ...at, op: 'dataDefault', field: o.valueDefault })
    else if ('dropValue' in o) add(role, set('data', [['Orientation', AXIS.hidden]], { field: o.dropValue }))
    else if ('addFields' in o) add(role, call('table', 'AddFields', [o.addFields.rows ?? null, o.addFields.columns ?? null, null, true]))
    else if ('cubeField' in o) add(role, set('cubeField', [['Orientation', AXIS[o.to]]], { field: o.cubeField }))
    else if ('cellFormat' in o) {
      const [sheet, range] = o.cellFormat.split('!')
      add(role, { ...at, op: 'set', on: 'range', rangeSheet: sheet, range, props: [['NumberFormat', o.numberFormat]] })
    } else if ('merge' in o || 'addTable' in o) {
      const [sheet, range] = ('merge' in o ? o.merge : o.addTable).split('!')
      add(role, { do: 'struct', op: 'merge' in o ? 'mergeCells' : 'addTable', sheet, range, target: 'merge' in o ? '' : o.name })
    }
    else if ('valuesOn' in o) add(role, set('values', [['Orientation', AXIS[o.valuesOn]], ...(o.position ? [['Position', o.position] as [string, unknown]] : [])]))
    else if ('layout' in o) add(role, call('table', 'RowAxisLayout', [ROW_LAYOUT[o.layout]]))
    else if ('repeatAll' in o) add(role, call('table', 'RepeatAllLabels', [o.repeatAll ? 2 : 1]))
    else if ('table' in o) add(role, set('table', tableProps(o.table)))
    else if ('props' in o) add(role, set('field', fieldProps(o.props), { field: o.field }))
    else if ('subtotals' in o) {
      const pairs = o.fns === 'automatic' ? [[1, true]] : o.fns === 'none' ? [[1, false]] : [[1, false], ...o.fns.map((f) => [SUBTOTAL[f], true])]
      add(role, { ...at, op: 'put', on: 'field', field: o.subtotals, name: 'Subtotals', props: pairs })
    } else if ('hide' in o) for (const item of o.items) add(role, set('item', [['Visible', false]], { field: o.hide, item }))
    else if ('page' in o) {
      if (o.items === '(All)') add(role, set('field', [['CurrentPage', '(All)']], { field: o.page }))
      else if (o.items.length === 1) add(role, set('field', [['EnableMultiplePageItems', false], ['CurrentPage', o.items[0]]], { field: o.page }))
      else add(role, set('field', [['EnableMultiplePageItems', true]], { field: o.page }), { ...at, op: 'only', on: 'field', field: o.page, items: o.items })
    } else if ('filter' in o) add(role, { ...at, op: 'filter', field: o.filter, type: FILTER_TYPES[o.type], dataField: o.by ?? null, value1: o.value1 ?? null, value2: o.value2 ?? null, wholeDays: o.wholeDays ?? null })
    else if ('sort' in o) add(role, call('field', 'AutoSort', [SORT_ORDER[o.order], o.by ?? o.sort], { field: o.sort }))
    else if ('item' in o) add(role, set('item', [['Position', o.position]], { field: o.of, item: o.item }))
    else if ('calcField' in o) add(role, call('calcFields', 'Add', [o.calcField, o.formula, true]))
    else if ('calcItem' in o) add(role, call('calcItems', 'Add', [o.calcItem, o.formula, true], { field: o.field }))
    else if ('group' in o) add(role, { ...at, op: 'group', field: o.group, start: o.start ?? null, end: o.end ?? null, by: o.by ?? null, periods: o.periods ? PERIODS.map((p) => o.periods!.includes(p)) : null })
    else if ('groupItems' in o) add(role, { ...at, op: 'groupItems', field: o.groupItems, items: o.items })
    else if ('collapse' in o) {
      const detail: Props = [['ShowDetail', o.expand ?? false]]
      if (o.items === 'all') add(role, set('field', detail, { field: o.collapse }))
      else for (const item of o.items) add(role, set('item', detail, { field: o.collapse, item }))
    } else if ('ungroup' in o) add(role, { ...at, op: 'ungroup', field: o.ungroup })
    else if ('slicer' in o) {
      slicers.push(o.name)
      add(role, { ...at, op: 'slicer', field: o.slicer, name: o.name, kind: 1, at: o.on ?? 'P', also: (o.pivots ?? []).map((p) => [sheets.get(p) ?? 'P', p]) })
    } else if ('select' in o) add(role, { ...at, op: 'only', on: 'slicerCache', field: o.select, items: o.items })
    else if ('deleteSlicer' in o) add(role, call('slicerCache', 'Delete', [], { field: o.deleteSlicer }))
    else if ('addSheet' in o) add(role, { do: 'struct', op: 'insertSheet', target: o.addSheet })
    else if ('deleteSheet' in o) add(role, { do: 'struct', op: 'deleteSheet', sheet: o.deleteSheet })
    else if ('slicerSettings' in o) {
      const props: Props = []
      if (o.order) props.push(['SortItems', SLICER_SORT[o.order]])
      if (o.noData) props.push(['CrossFilterType', CROSS_FILTER[o.noData]])
      add(role, set('slicerCache', props, { field: o.slicerSettings }))
    } else if ('link' in o) add(role, { do: 'pivot', op: 'link', field: o.link, sheet: sheets.get(o.pivot) ?? 'P', pt: o.pivot, off: o.off ?? false })
    else if ('timeline' in o) {
      slicers.push(o.name)
      add(role, { ...at, op: 'slicer', field: o.timeline, name: o.name, kind: 2, at: 'P', also: [] })
      if (o.from !== undefined && o.to !== undefined) add(role, call('timeline', 'SetFilterDateRange', [{ d: o.from }, { d: o.to }], { field: o.name }))
      if (o.level) add(role, set('timelineView', [['Level', LEVEL[o.level]]], { field: o.name }))
    } else if ('cache' in o) {
      const props: Props = []
      if (o.cache.missingItems) props.push(['MissingItemsLimit', MISSING[o.cache.missingItems]])
      if (o.cache.refreshOnLoad !== undefined) props.push(['RefreshOnFileOpen', o.cache.refreshOnLoad])
      add(role, set('cache', props))
    } else if ('edit' in o) {
      const [sheet, cell] = o.edit.split('!')
      add(role, { do: 'set', sheet, cell, ...(o.to !== null && typeof o.to === 'object' && 'typed' in o.to ? { kind: 'f2', value: o.to.typed } : comEdit(o.to)) })
    } else if ('drill' in o) {
      const [sheet, cell] = o.drill.split('!')
      add(role, { ...at, op: 'drill', sheet, cell })
    } else if ('refresh' in o) add(role, { ...at, op: 'refresh' })
    else if ('refreshAll' in o) add(role, { do: 'pivot', op: 'refreshAll' })
    else if ('define' in o) add(role, { do: 'table', line: ['name', '', o.define, `=${o.ref}`].join('\t') })
    else if ('resize' in o) add(role, { do: 'pivot', op: 'resize', table: o.resize, range: o.range })
    else if ('external' in o) add({ external: Object.keys(o.cells) }, { do: 'pivot', op: 'external', source: Buffer.from(book(o.external)).toString('base64'), cells: o.cells })
  })
  phase(true)
  return { steps, roles }
}

const cellValue = (c: WinCell): CellValue => {
  const v = valOf(c)
  if (!v) return null
  return v.t === 'e' ? { error: v.v as Extract<CellValue, { error: string }>['error'] } : (v.v as CellValue)
}

const errorOf = (x: unknown): string | undefined => (x && typeof x === 'object' && 'error' in x ? String((x as { error: unknown }).error) : undefined)

/** A pivot table's area as `pivot.ps1` reads it. */
function areaOf(got: unknown): PivotArea {
  const [table, area, body, values, text, formats] = got as [string, string, string, WinCell[][], string[][], string[][]]
  return { table, area, body, values: values.map((r) => r.map(cellValue)), text, formats }
}

/** A case's answer from its steps' answers, phase by phase. */
export function pivotAnswer(plan: Planned, answers: unknown[]): PivotAnswer {
  const phases: PivotPhase[] = []
  let cur: PivotPhase = { pivots: {} }
  plan.roles.forEach((role, i) => {
    const got = answers[i]
    const error = errorOf(got)
    if ('op' in role) {
      if (error) (cur.errors ??= []).push(`${role.op === -1 ? 'create' : role.op === -2 ? 'calculate' : `op ${role.op}`}: ${error}`)
      else if (got !== null && role.op >= 0) (cur.results ??= {})[`op ${role.op}`] = got
    } else if ('read' in role) cur.pivots[role.read] = error ? { error } : areaOf(got)
    else if ('readAll' in role) {
      if (error) cur.pivots['*'] = { error }
      else for (const [sheet, name, area] of got as [string, string, unknown][]) cur.pivots[`${sheet}!${name}`] = errorOf(area) ? { error: errorOf(area)! } : areaOf(area)
    } else if ('watch' in role) {
      const [sheet, range] = role.watch.split('!')
      const vals = (got as WinCell[][]).flat()
      cellsOf(range).addrs.forEach((a, j) => ((cur.cells ??= {})[`${sheet}!${a}`] = cellValue(vals[j])))
    } else if ('slicer' in role) (cur.slicers ??= {})[role.slicer] = error ? { error } : (got as NonNullable<PivotPhase['slicers']>[string])
    else if ('external' in role) {
      const reads: Record<string, Record<string, CellValue>> = {}
      const errors: string[] = []
      for (const x of got as unknown[]) {
        const e = errorOf(x)
        if (e) errors.push(e)
        else {
          const [label, vals] = x as [string, WinCell[]]
          reads[label] = Object.fromEntries(vals.map((v, j) => [role.external[j], cellValue(v)]))
        }
      }
      cur.external = { reads, ...(errors.length ? { errors } : {}) }
    } else {
      phases.push(cur)
      cur = { pivots: {} }
    }
  })
  return { phases }
}

/** A set's saved book (`basic/sum`), as the harness stored it. */
const savedBook = (tag: Tag) => (name: string) => new Uint8Array(readFileSync(winFile(`pivot/${localeDir(tag)}${name}.xlsx`)))

export async function stepPivot(set: string, tag: Tag, e: WinEnv): Promise<void> {
  const cases = (await setModule<PivotSet>('pivot', set)).buildPivotCases()
  const saved = savedBook(tag)
  const { files, sent: parts } = await caseFiles(cases, async (c) => (c.file ? new Uint8Array(readFileSync(resolve(ROOT, c.file))) : c.from ? saved(c.from) : (await writeXlsx({ book: sourceBook(DATA[c.data ?? 'sales'](), c.source ?? 'range') }, { values: 'omit' })).data))
  const plans = new Map(cases.map((c) => [c.id, pivotSteps(c, saved)]))
  // a case runs again when its file or its steps change (the steps carry the books of external operations)
  const sent = Object.fromEntries(cases.map((c) => [c.id, `${parts[c.id]} ${createHash('sha256').update(JSON.stringify(plans.get(c.id)!.steps)).digest('hex').slice(0, 16)}`]))
  const file = winFile(`pivot/${localeDir(tag)}${set}.json`)
  const dir = file.slice(0, -'.json'.length)
  const store = fileAnswers<PivotCase, PivotAnswer>(file, cases, sent)
  const t0 = Date.now()
  const save = () => store.save({ ...metaOf(e, tag), date: new Date().toISOString().slice(0, 10), seconds: secondsSince(t0) }, { cases })
  await fileJobs(
    `pivot ${set} (${tag})`,
    `pivot-${set}`,
    store,
    files,
    (c) => plans.get(c.id)!.steps,
    (c, a) => {
      if (a.copies) mkdirSync(dir, { recursive: true })
      a.copies?.forEach((x, k) => writeFileSync(join(dir, `${c.id}-${k + 1}.xlsx`), x))
      return pivotAnswer(plans.get(c.id)!, a.answers)
    },
    save,
  )
  pruneSaved(dir, new Set(cases.map((c) => c.id)))
}

/** Removes the saved books of cases the set no longer has (`<id>.xlsx`, `<id>-<k>.xlsx`). */
function pruneSaved(dir: string, ids: Set<string>) {
  if (!existsSync(dir)) return
  for (const f of readdirSync(dir)) {
    const id = f.replace(/(-\d+)?\.xlsx$/, '')
    if (!ids.has(id) && !ids.has(f.slice(0, -'.xlsx'.length))) rmSync(join(dir, f))
  }
}
