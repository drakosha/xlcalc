/** Bounded queries executed in the session: no whole-sheet value copy; a search keeps the texts it reads with the state (`search.ts`). */
import type { CellReadField, CellValue, CellView, Clock, DependentCell, DependentsOptions, DependentsResult, FindCellsOptions, PageOptions, QueryPage, RangeRef, RangeSummary, SummaryOptions, ReadCell, ReadCellsOptions, ReadCellsResult } from '../contract.js'
import { sheetIndexOf, toCellValue, valueAt, type EngineState, type Formula } from '../engine.js'
import { CalcError, paced, type Steps } from '../exec/budget.js'
import { readersOf, referencesOf } from '../graph.js'
import { checkSheet } from '../input.js'
import { InputError } from '../input-error.js'
import { array, checkBoolean, checkNumber, checkString, object, onlyKeys } from '../input/basic.js'
import { checkAddress, checkRange } from '../input/range.js'
import { issueDetail } from '../issues.js'
import { byPlace, colOf, formatAddress, MAX_COLS, MAX_ROWS, parseAddress, placeOf, rowOf } from '../model.js'
import { cellsAt, outValue } from '../result.js'
import { rectText, meets } from '../structure/area.js'
import { CompensatedSum } from '../functions/statistical/sum.js'
import { checkBase } from './answer.js'
import { inputIn } from '../looks/sheet.js'
import { cellValue } from './looks.js'
import { checkLineNumbers, sessionEditTexts } from './viewport.js'
import { FIND_CALL_OPTIONS, READ_OPTIONS } from '../strict-args.js'
import { finderOf, type Finder, type Ref } from './find.js'
import { earlierOf, firstIn, nextIn, readOrder, spillsIn } from './read-order.js'
import { textsOf } from './search.js'

const PAGE = ['maxCells', 'cursor', 'count'] as const
const FIELDS: CellReadField[] = ['value', 'text', 'formula', 'editText', 'style', 'numberFormat', 'note', 'link', 'issue']
/** A query's ranges checked: rectangles of the book's sheets. */
export const refsOf = (s: EngineState, refs: RangeRef[]) =>
  array(refs, 'ranges', 'an array of sheet/range references').map((v, i): Ref => {
    const p = `ranges[${i}]`,
      o = object(v, p, 'a sheet and range')
    onlyKeys(o, p, ['sheet', 'range'])
    return { ...checkRange(o.range, `${p}.range`), s: checkSheet(o.sheet, (n) => sheetIndexOf(s, n), `${p}.sheet`) }
  })

/** A cursor's parts: the revision, the results before it, the query's signature, the count of all (-1 when not counted), and the position the query goes on after. */
const isCursor = (c: unknown): c is [number, number, string, number, ...number[]] => Array.isArray(c) && c.length >= 5 && c.slice(3).every(Number.isSafeInteger) && c[1] >= 0 && Number.isSafeInteger(c[1])

class Page<T> {
  readonly items: T[] = []
  readonly max: number
  readonly count: boolean
  /** The results before the cursor, the position of the last of them in the query's order, and the count of all when a page before counted them. */
  readonly offset: number = 0
  readonly from: number[] | null = null
  readonly total: number | null = null
  matched = 0
  /** The results the page takes, which its items hold once made. */
  taken = 0
  more = false
  /** The position of the last result of the page. */
  last: number[] = []
  constructor(
    readonly revision: number,
    readonly clock: Clock,
    readonly signature: string,
    o: PageOptions,
  ) {
    this.max = o.maxCells ?? 1000
    checkNumber(this.max, 'opts.maxCells', 1, 10000, true)
    if (o.count !== undefined) checkBoolean(o.count, 'opts.count')
    this.count = o.count === true
    if (o.cursor !== undefined) {
      let c: unknown
      try {
        c = JSON.parse(checkString(o.cursor, 'opts.cursor'))
      } catch {
        throw new InputError('opts.cursor', 'expected a cursor returned by this query')
      }
      if (!isCursor(c)) throw new InputError('opts.cursor', 'expected a cursor returned by this query')
      if (c[0] !== revision) throw new CalcError('stale_revision', 'the cursor belongs to another revision')
      if (c[2] !== signature) throw new InputError('opts.cursor', 'the cursor belongs to another query')
      this.offset = c[1]
      this.total = c[3] >= 0 ? c[3] : null
      this.from = c.slice(4) as number[]
    }
  }
  /** A result at `position` of the query's order: whether it goes on the page, or the page is full. */
  wanted(position: number[]): boolean {
    this.matched++
    if (this.taken < this.max) {
      this.taken++
      this.last = position
      return true
    }
    this.more = true
    return false
  }
  get done(): boolean {
    return this.more && (!this.count || this.total !== null)
  }
  result(): QueryPage {
    const total = this.total ?? (this.count ? this.offset + this.matched : -1)
    return { revision: this.revision, clock: { ...this.clock }, truncated: this.more, ...(this.more ? { cursor: JSON.stringify([this.revision, this.offset + this.taken, this.signature, total, ...this.last]) } : {}), ...(this.count ? { total } : {}) }
  }
}

/**
 * A cell's fields `fields`. Those of its input from its input; the others from `view`, the cell shown as `getCell` gives
 * it, when `text` is asked; without, from its value and input alone, no result of the cell built (`view` null).
 */
function projected(s: EngineState, i: number, r: number, c: number, addr: string, fields: CellReadField[], edit: ReturnType<typeof sessionEditTexts> | null, view: CellView | undefined | null): ReadCell {
  const out: ReadCell = { sheet: s.sheetNames[i], cell: addr }
  const input = inputIn(s.input.sheets[i], addr, r, c)
  for (const f of fields) {
    if (f === 'issue') out.issue = issueDetail(s, i, r, c)
    else if (f === 'editText') out.editText = edit!(i, r, c, input)
    else if (f === 'style' && input?.style !== undefined) out.style = input.style
    else if (f === 'note' && input?.note) out.note = { ...input.note }
    else if (f === 'link' && input?.link) out.link = { ...input.link }
    else if (view !== null) {
      if (view![f as keyof CellView] !== undefined) (out as unknown as Record<string, unknown>)[f] = view![f as keyof CellView]
    } else if (f === 'value') out.value = cellValue(s, i, r, c)
    else if (f === 'formula') {
      const id = s.at.get(placeOf(i, r, c))
      const text = id === undefined ? input?.formula : s.formulas[id]!.text
      if (text !== undefined) out.formula = text
    } else if (f === 'numberFormat' && input?.numberFormat !== undefined) out.numberFormat = input.numberFormat
  }
  return out
}

/** A cell of a page: its sheet, row, column and address. */
type Cell = [s: number, r: number, c: number, addr: string]

/** The fields `fields` of a page's cells (`projected`); with `text`, their views made together (`cellsAt`). */
function* projectedAll(s: EngineState, cells: Cell[], fields: CellReadField[], edit: ReturnType<typeof sessionEditTexts> | null): Steps<ReadCell[]> {
  const views = fields.includes('text') ? yield* cellsAt(s, cells, true) : null
  return cells.map(([i, r, c, addr], k) => projected(s, i, r, c, addr, fields, edit, views ? views[k] : null))
}

/** A query's ranges, fields and rows checked, and its page; `key`, what else makes the query the one a cursor belongs to. */
function queryOf(s: EngineState, revision: number, clock: Clock, given: RangeRef[], o: ReadCellsOptions, key: unknown[]) {
  const refs = refsOf(s, given)
  const fields = o.fields === undefined ? (['value', 'text', 'formula'] as CellReadField[]) : (array(o.fields, 'opts.fields', 'an array of cell field names') as CellReadField[])
  if (fields.some((f) => !FIELDS.includes(f))) throw new InputError('opts.fields', 'unknown cell field')
  if (o.rows !== undefined) checkLineNumbers(o.rows, 'opts.rows', MAX_ROWS)
  const rows = o.rows?.map((n) => n - 1) ?? null
  const page = new Page<ReadCell>(revision, clock, JSON.stringify([refs, fields, o.rows ?? null, ...key]), o)
  return { refs, fields, rows, page, edit: fields.includes('editText') ? sessionEditTexts(s, clock.now) : null }
}

export function* readQuery(s: EngineState, revision: number, clock: Clock, given: RangeRef[], eo?: ReadCellsOptions): Steps<ReadCellsResult> {
  const { refs, fields, rows, page, edit } = queryOf(s, revision, clock, given, checkBase<ReadCellsOptions>(eo, revision, READ_OPTIONS), [])
  const cells: Cell[] = []
  for (const cell of readOrder(s, refs, rows, page.from)) {
    if (!cell) {
      yield
      continue
    }
    const [i, r, c, n, addr] = cell
    if (page.wanted([n, r * MAX_COLS + c])) cells.push([i, r, c, addr])
    if (page.done) break
  }
  page.items.push(...(yield* projectedAll(s, cells, fields, edit)))
  return { ...page.result(), cells: page.items }
}

/** The position in Find's order of the cell a search begins after (`FindCellsOptions.after`), checked: a cell of the ranges. */
function startOf(s: EngineState, refs: Ref[], finder: Finder, after: unknown): number[] {
  const o = object(after, 'opts.after', 'a cell such as { sheet: "Sheet1", cell: "D7" }')
  onlyKeys(o, 'opts.after', ['sheet', 'cell'])
  const i = checkSheet(o.sheet, (n) => sheetIndexOf(s, n), 'opts.after.sheet')
  const { row, col } = parseAddress(checkAddress(o.cell, 'opts.after.cell'))
  const at = finder.positionOf(refs, i, row, col)
  if (!at) throw new InputError('opts.after', `expected a cell of the ranges searched, got ${s.sheetNames[i]}!${formatAddress(row, col)}`)
  return at
}

export function* findQuery(s: EngineState, revision: number, clock: Clock, given: RangeRef[], query: string, eo?: FindCellsOptions): Steps<ReadCellsResult> {
  const o = checkBase<FindCellsOptions>(eo, revision, FIND_CALL_OPTIONS)
  const finder = finderOf(s, checkString(query, 'query'), o, clock.now)
  const { refs, fields, rows, page, edit } = queryOf(s, revision, clock, given, o, [...finder.key, o.after ?? null])
  const start = o.after === undefined ? null : startOf(s, refs, finder, o.after)
  const cells: Cell[] = []
  for (const cell of start ? finder.around(refs, rows, start, page.from) : finder.found(refs, rows, page.from)) {
    if (!cell) {
      yield
      continue
    }
    const [i, r, c, position] = cell
    if (page.wanted(position)) cells.push([i, r, c, formatAddress(r, c)])
    if (page.done) break
  }
  page.items.push(...(yield* projectedAll(s, cells, fields, edit)))
  return { ...page.result(), cells: page.items }
}

export function* dependentsQuery(s: EngineState, revision: number, clock: Clock, given: RangeRef[], eo?: DependentsOptions): Steps<DependentsResult> {
  const o = checkBase<DependentsOptions>(eo, revision, [...PAGE, 'maxReferences']),
    refs = refsOf(s, given)
  const max = o.maxReferences ?? 64
  checkNumber(max, 'opts.maxReferences', 1, 1000, true)
  const page = new Page<DependentCell>(revision, clock, JSON.stringify(['dependents', refs, max]), o)
  if (page.from && page.from.length !== 2) throw new InputError('opts.cursor', 'expected a cursor returned by this query')
  // the formulas the graph's index of references names, in the book's order of sheets, then by rows: a cursor keeps the sheet and place of the last
  const ids = new Set<number>()
  for (const x of refs) readersOf(s, x, ids)
  const found: Formula[] = []
  for (const id of ids) {
    const f = s.formulas[id]
    if (f && s.input.sheets[f.s]?.cells[formatAddress(f.r, f.c)]?.formula !== undefined && s.at.get(placeOf(f.s, f.r, f.c)) === id) found.push(f)
  }
  found.sort(byPlace)
  const [fs, fm] = page.from ?? [-1, -1]
  for (const f of found) {
    if (paced(1)) yield
    const m = f.r * MAX_COLS + f.c
    if (f.s < fs || (f.s === fs && m <= fm)) continue
    const matches = referencesOf(s, f).filter((g) => refs.some((x) => x.s === g.s && meets(x, g)))
    if (!matches.length) continue
    if (page.wanted([f.s, m])) page.items.push({ sheet: s.sheetNames[f.s], cell: formatAddress(f.r, f.c), references: matches.slice(0, max).map((g) => ({ sheet: s.sheetNames[g.s], range: rectText(g) })), ...(matches.length > max ? { referencesTruncated: true } : {}) })
    if (page.done) break
  }
  return { ...page.result(), dependents: page.items }
}

export function* summaryQuery(s: EngineState, revision: number, clock: Clock, given: RangeRef[], eo?: SummaryOptions): Steps<RangeSummary> {
  const o = checkBase<SummaryOptions>(eo, revision, ['rows'])
  if (o.rows !== undefined) checkLineNumbers(o.rows, 'opts.rows', MAX_ROWS)
  const out: RangeSummary = { revision, clock: { ...clock }, cells: 0, numbers: 0, errors: 0, sum: 0, min: null, max: null },
    sum = new CompensatedSum()
  const add = (v: CellValue) => {
    if (v === null || v === '') return
    out.cells++
    if (typeof v === 'number') {
      out.numbers++
      sum.add(v)
      out.min = out.min === null ? v : Math.min(out.min, v)
      out.max = out.max === null ? v : Math.max(out.max, v)
    } else if (typeof v === 'object') out.errors++
  }
  const refs = refsOf(s, given)
  const rows = o.rows?.map((n) => n - 1) ?? null
  for (const [n, x] of refs.entries()) {
    const earlier = earlierOf(refs, n)
    if (spillsIn(s, x).length) {
      // spilled cells among the range's: read in readCells' order, this range's cells alone
      for (const at of readOrder(s, refs, rows, [n, -1], false)) {
        if (!at) yield
        else if (at[3] !== n) break
        else add(cellValue(s, at[0], at[1], at[2]))
      }
      continue
    }
    // no spill: the values of the index's places by rows, read from the columns; a place whose cell is gone holds none
    const { at } = yield* textsOf(s, x.s)
    for (let j = firstIn(at, x, rows, 0); j < at.length; j = nextIn(at, x, rows, 0, j + 1)) {
      if (paced(1)) yield
      const r = rowOf(at[j])
      const c = colOf(at[j])
      if (!earlier(r, c)) add(cellValue(s, x.s, r, c))
    }
  }
  out.sum = Number.isFinite(sum.value) ? sum.value : { error: 'NUM' }
  return out
}
