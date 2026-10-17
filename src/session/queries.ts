/** Bounded queries executed in the session: no whole-sheet value copy; a search keeps the texts it reads with the state (`search.ts`). */
import type { CellReadField, CellView, Clock, DependentCell, DependentsOptions, DependentsResult, FindCellsOptions, PageOptions, QueryPage, RangeRef, RangeSummary, SummaryOptions, ReadCell, ReadCellsOptions, ReadCellsResult } from '../contract.js'
import { key, toCellValue, valueAt, type EngineState } from '../engine.js'
import { CalcError, paced, type Steps } from '../exec/budget.js'
import { referencesOf } from '../graph.js'
import { array, checkAddress, checkBoolean, checkNumber, checkSheet, checkString, InputError, object, onlyKeys } from '../input.js'
import { issueDetail } from '../issues.js'
import { formatAddress, MAX_COLS, MAX_ROWS, parseAddress } from '../model.js'
import { cellAt, outValue } from '../result.js'
import { rectText, meets } from '../structure/area.js'
import { sheetKey } from '../syntax.js'
import { CompensatedSum } from '../functions/statistical/sum.js'
import { checkBase } from './answer.js'
import { inputIn } from '../looks/sheet.js'
import { cellValue } from './looks.js'
import { checkLineNumbers, checkRange, sessionEditTexts } from './viewport.js'
import { FIND_CALL_OPTIONS, READ_OPTIONS } from '../strict-args.js'
import { finderOf, type Finder, type Ref } from './find.js'
import { readOrder } from './read-order.js'

const PAGE = ['maxCells', 'cursor', 'count'] as const
const FIELDS: CellReadField[] = ['value', 'text', 'formula', 'editText', 'style', 'numberFormat', 'note', 'link', 'issue']
/** A query's ranges checked: rectangles of the book's sheets. */
export const refsOf = (s: EngineState, refs: RangeRef[]) =>
  array(refs, 'ranges', 'an array of sheet/range references').map((v, i): Ref => {
    const p = `ranges[${i}]`,
      o = object(v, p, 'a sheet and range')
    onlyKeys(o, p, ['sheet', 'range'])
    return { ...checkRange(o.range, `${p}.range`), s: checkSheet(o.sheet, (n) => s.sheetIdx.get(sheetKey(n)), `${p}.sheet`) }
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
    if (this.items.length < this.max) {
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
    return { revision: this.revision, clock: { ...this.clock }, truncated: this.more, ...(this.more ? { cursor: JSON.stringify([this.revision, this.offset + this.items.length, this.signature, total, ...this.last]) } : {}), ...(this.count ? { total } : {}) }
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
      const id = s.at.get(key(i, r, c))
      const text = id === undefined ? input?.formula : s.formulas[id]!.text
      if (text !== undefined) out.formula = text
    } else if (f === 'numberFormat' && input?.numberFormat !== undefined) out.numberFormat = input.numberFormat
  }
  return out
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
  const shown = fields.includes('text')
  for (const cell of readOrder(s, refs, rows, page.from)) {
    if (!cell) {
      yield
      continue
    }
    const [i, r, c, n, addr] = cell
    if (page.wanted([n, r * MAX_COLS + c])) page.items.push(projected(s, i, r, c, addr, fields, edit, shown ? yield* cellAt(s, i, r, c, true, true) : null))
    if (page.done) break
  }
  return { ...page.result(), cells: page.items }
}

/** The position in Find's order of the cell a search begins after (`FindCellsOptions.after`), checked: a cell of the ranges. */
function startOf(s: EngineState, refs: Ref[], finder: Finder, after: unknown): number[] {
  const o = object(after, 'opts.after', 'a cell such as { sheet: "Sheet1", cell: "D7" }')
  onlyKeys(o, 'opts.after', ['sheet', 'cell'])
  const i = checkSheet(o.sheet, (n) => s.sheetIdx.get(sheetKey(n)), 'opts.after.sheet')
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
  for (const cell of start ? finder.around(refs, rows, start, page.from) : finder.found(refs, rows, page.from)) {
    if (!cell) {
      yield
      continue
    }
    const [i, r, c, position] = cell
    if (page.wanted(position)) page.items.push(projected(s, i, r, c, formatAddress(r, c), fields, edit, fields.includes('text') ? yield* cellAt(s, i, r, c, true, true) : null))
    if (page.done) break
  }
  return { ...page.result(), cells: page.items }
}

export function* dependentsQuery(s: EngineState, revision: number, clock: Clock, given: RangeRef[], eo?: DependentsOptions): Steps<DependentsResult> {
  const o = checkBase<DependentsOptions>(eo, revision, [...PAGE, 'maxReferences']),
    refs = refsOf(s, given)
  const max = o.maxReferences ?? 64
  checkNumber(max, 'opts.maxReferences', 1, 1000, true)
  const page = new Page<DependentCell>(revision, clock, JSON.stringify(['dependents', refs, max]), o)
  // the cells before the cursor are passed over, by their count
  const from = page.from?.[0] ?? 0
  let at = 0
  scan: for (let i = 0; i < s.input.sheets.length; i++)
    for (const addr in s.input.sheets[i].cells) {
      if (paced(1)) yield
      if (++at <= from || s.input.sheets[i].cells[addr].formula === undefined) continue
      const { row, col } = parseAddress(addr),
        id = s.at.get(key(i, row, col)),
        f = id === undefined ? undefined : s.formulas[id]
      if (!f) continue
      const matches = referencesOf(s, f).filter((g) => refs.some((x) => x.s === g.s && meets(x, g)))
      if (!matches.length) continue
      if (page.wanted([at])) page.items.push({ sheet: s.sheetNames[f.s], cell: formatAddress(f.r, f.c), references: matches.slice(0, max).map((g) => ({ sheet: s.sheetNames[g.s], range: rectText(g) })), ...(matches.length > max ? { referencesTruncated: true } : {}) })
      if (page.done) break scan
    }
  return { ...page.result(), dependents: page.items }
}

export function* summaryQuery(s: EngineState, revision: number, clock: Clock, given: RangeRef[], eo?: SummaryOptions): Steps<RangeSummary> {
  const o = checkBase<SummaryOptions>(eo, revision, ['rows'])
  if (o.rows !== undefined) checkLineNumbers(o.rows, 'opts.rows', MAX_ROWS)
  const out: RangeSummary = { revision, clock: { ...clock }, cells: 0, numbers: 0, errors: 0, sum: 0, min: null, max: null },
    sum = new CompensatedSum()
  for (const at of readOrder(s, refsOf(s, given), o.rows?.map((n) => n - 1) ?? null, null, false)) {
    if (!at) {
      yield
      continue
    }
    const v = outValue(toCellValue(valueAt(s, at[0], at[1], at[2])))
    if (v === null || v === '') continue
    out.cells++
    if (typeof v === 'number') {
      out.numbers++
      sum.add(v)
      out.min = out.min === null ? v : Math.min(out.min, v)
      out.max = out.max === null ? v : Math.max(out.max, v)
    } else if (typeof v === 'object') out.errors++
  }
  out.sum = Number.isFinite(sum.value) ? sum.value : { error: 'NUM' }
  return out
}
