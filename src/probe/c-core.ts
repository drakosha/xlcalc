/**
 * Вариант C: AST нашего разбора компилируется в код JS (одна `new Function` на книгу, только над
 * сгенерированным кодом: литералы через `JSON.stringify`, вызовы только функций рантайма из белого
 * списка), ячейки считаются в топологическом порядке графа из AST, диапазоны это срезы колонок.
 *
 * Без импорта node:*: тот же код собирается в бандл для замера в QuickJS.
 */
import { CErr, E, makeRuntime, Rng, type RuntimeContext, type SheetData, type V } from './c-runtime'
import { formatAddress, inputScalar, parseAddress, type BookInput, type CellValue, type Clock } from './model'
import type { Ast } from './parser'
import type { CellEdit } from './engine'
import { invalidError, prepareBook, type InvalidFormula, type PreparedBook } from './prep'
import { nowSerial, WHITELIST } from './functions'

/** Функции, которые рантайм C умеет (прототип: синтетика и список семантики). */
export const C_FUNCTIONS: ReadonlySet<string> = new Set(['SUM', 'SUMIFS', 'COUNTIFS', 'IF', 'IFERROR', 'XLOOKUP', 'ROUND', 'EOMONTH', 'DATE', 'TEXT', 'TRUE', 'FALSE', 'TODAY', 'NOW', 'RAND', 'RANDBETWEEN'])
const RANGE_FUNCTIONS = new Set(['SUM', 'SUMIFS', 'COUNTIFS', 'XLOOKUP'])

interface Formula {
  s: number
  r: number
  c: number
  refs: Rng[]
}

export interface CState {
  book: BookInput
  opts: CompileOptions
  prep: PreparedBook
  ctx: RuntimeContext
  fns: (() => V | Rng)[]
  formulas: Formula[]
  /** Позиция формулы в топологическом порядке; -1 у ячеек цикла. */
  topo: number[]
  order: number[]
  succ: number[][]
  /** Слушатели: лист → колонка → [r1, r2, формула]. */
  listeners: Map<number, Map<number, [number, number, number][]>>
  /** Формулы, которые C не умеет (белый список есть, реализации в прототипе нет). */
  unsupported: { sheet: string; addr: string; fn: string }[]
  invalid: InvalidFormula[]
  ms: { parse: number; build: number; evaluate: number }
  generatedChars: number
}

export interface CompileOptions {
  clock: Clock
  ifs: 'fj' | 'idx'
}

/** Миллисекунды: в QuickJS нет `performance`. */
const clockMs = () => (typeof performance !== 'undefined' ? performance.now() : Date.now())

function emptySheet(name: string): SheetData {
  return { name, cols: [], rows: 0 }
}

function put(sh: SheetData, r: number, c: number, v: V) {
  ;(sh.cols[c] ??= [])[r] = v
  if (r + 1 > sh.rows) sh.rows = r + 1
}

/** Код JS для AST; диапазоны складываются в `G`, имена разворачиваются в свои ссылки. */
function codegen(ast: Ast, sheetIdx: (s: string) => number, here: number, G: Rng[], nameRef: (n: string) => Ast | null, unsupported: string[]): string {
  const range = (n: Extract<Ast, { k: 'ref' }>) => {
    const s = n.sheet === null ? here : sheetIdx(n.sheet)
    G.push(new Rng(s, n.ref.r1, n.ref.c1, n.ref.r2, n.ref.c2))
    return `G[${G.length - 1}]`
  }
  const deref = (n: Ast): Ast => {
    if (n.k === 'name') return nameRef(n.name) ?? n
    if (n.k === 'paren') return deref(n.a)
    return n
  }
  const asRange = (n: Ast): string => {
    const d = deref(n)
    return d.k === 'ref' ? range(d) : go(n)
  }
  const go = (n: Ast): string => {
    switch (n.k) {
      case 'num':
      case 'str':
      case 'bool':
        return JSON.stringify(n.v)
      case 'err':
        return `E.${n.code}`
      case 'ref': {
        if (n.ref.kind === 'cell') return `cell(${n.sheet === null ? here : sheetIdx(n.sheet)},${n.ref.r1},${n.ref.c1})`
        return `scalar(${range(n)})`
      }
      case 'name': {
        const d = nameRef(n.name)
        return d ? go(d) : 'E.NAME'
      }
      case 'paren':
        return `(${go(n.a)})`
      case 'un':
        return n.op === '-' ? `neg(${go(n.a)})` : go(n.a)
      case 'pct':
        return `pct(${go(n.a)})`
      case 'bin': {
        const a = go(n.a)
        const b = go(n.b)
        switch (n.op) {
          case '+':
            return `add(${a},${b})`
          case '-':
            return `sub(${a},${b})`
          case '*':
            return `mul(${a},${b})`
          case '/':
            return `div(${a},${b})`
          case '^':
            return `pow(${a},${b})`
          case '&':
            return `cat(${a},${b})`
          default:
            return `cmp(${JSON.stringify(n.op)},${a},${b})`
        }
      }
      case 'fn': {
        const args = n.args
        if (n.name === 'IF') {
          if (args.length < 2 || args.length > 3) return 'E.VALUE'
          return `IF(${go(args[0])},()=>${go(args[1])},()=>${args[2] ? go(args[2]) : 'false'})`
        }
        if (n.name === 'IFERROR') {
          if (args.length !== 2) return 'E.VALUE'
          return `IFERR(()=>${go(args[0])},()=>${go(args[1])})`
        }
        if (n.name === 'RAND') return `rand(${here},R,C)`
        if (n.name === 'RANDBETWEEN') return args.length === 2 ? `randbetween(${here},R,C,${go(args[0])},${go(args[1])})` : 'E.VALUE'
        if (!C_FUNCTIONS.has(n.name)) {
          unsupported.push(n.name)
          return 'E.NAME'
        }
        const rangeArgs = RANGE_FUNCTIONS.has(n.name)
        return `F.${n.name}(${args.map((a) => (rangeArgs ? asRange(a) : go(a))).join(',')})`
      }
    }
  }
  return go(ast)
}

/** Ссылки формулы (после разворота имён) для графа. */
function collectRefs(ast: Ast, sheetIdx: (s: string) => number, here: number, nameRef: (n: string) => Ast | null, out: Rng[]) {
  const visit = (n: Ast) => {
    switch (n.k) {
      case 'ref': {
        const s = n.sheet === null ? here : sheetIdx(n.sheet)
        out.push(new Rng(s, n.ref.r1, n.ref.c1, n.ref.r2, n.ref.c2))
        break
      }
      case 'name': {
        const d = nameRef(n.name)
        if (d) visit(d)
        break
      }
      case 'fn':
        n.args.forEach(visit)
        break
      case 'un':
      case 'pct':
      case 'paren':
        visit(n.a)
        break
      case 'bin':
        visit(n.a)
        visit(n.b)
        break
    }
  }
  visit(ast)
}

export function compileBook(book: BookInput, opts: CompileOptions): CState {
  const t0 = clockMs()
  const prep = prepareBook(book, WHITELIST)
  const t1 = clockMs()

  const sheetIdxMap = new Map(prep.sheets.map((s, i) => [s, i]))
  const sheetIdx = (s: string) => sheetIdxMap.get(prep.sheetOf(s)) ?? -1
  const sheets = prep.sheets.map(emptySheet)
  for (const [i, s] of book.sheets.entries()) {
    for (const [addr, cell] of Object.entries(s.cells)) {
      if (cell.formula !== undefined) continue
      const { row, col } = parseAddress(addr)
      const v = inputScalar(cell.value)
      if (v !== null) put(sheets[i], row, col, v)
    }
  }
  for (const f of prep.invalid) {
    const { row, col } = parseAddress(f.addr)
    put(sheets[sheetIdx(f.sheet)], row, col, E[invalidError(f.code)])
  }

  const nameRef = (here: number) => (name: string): Ast | null => {
    const up = name.toUpperCase()
    const local = prep.names.find((n) => n.name.toUpperCase() === up && n.scope !== null && sheetIdx(n.scope) === here)
    const global = prep.names.find((n) => n.name.toUpperCase() === up && n.scope === null)
    return (local ?? global)?.ast ?? null
  }

  const G: Rng[] = []
  const bodies: string[] = []
  const formulas: Formula[] = []
  const unsupported: CState['unsupported'] = []
  for (const f of prep.formulas) {
    const s = sheetIdx(f.sheet)
    const un: string[] = []
    const code = codegen(f.ast, sheetIdx, s, G, nameRef(s), un)
    for (const fn of un) unsupported.push({ sheet: f.sheet, addr: f.addr, fn })
    bodies.push(`function(){const R=${f.row},C=${f.col};return ${code}}`)
    const refs: Rng[] = []
    collectRefs(f.ast, sheetIdx, s, nameRef(s), refs)
    formulas.push({ s, r: f.row, c: f.col, refs })
  }
  const src = `"use strict";const {cell,scalar,add,sub,mul,div,pow,neg,pct,cat,cmp,F,rand,randbetween,E,IF,IFERR}=rt;const G=rt.G;return [${bodies.join(',\n')}]`

  const ctx: RuntimeContext = { sheets, now: nowSerial(opts.clock), seed: opts.clock.seed, ifs: opts.ifs, flat: new Map(), index: new Map() }
  const base = rtOf(ctx)
  const IF = (c: V, a: () => V, b: () => V): V => {
    const x = base.toBool(base.scalar(c))
    if (x instanceof CErr) return x
    return x ? base.scalar(a()) : base.scalar(b())
  }
  const IFERR = (a: () => V, b: () => V): V => {
    const x = base.scalar(a())
    return x instanceof CErr ? base.scalar(b()) : x
  }
  const rt = { ...base, IF, IFERR, G }
  const fns = new Function('rt', src)(rt) as (() => V | Rng)[]

  // ---------- граф ----------
  const at = new Map<number, Map<number, { r: number; id: number }[]>>()
  formulas.forEach((f, id) => {
    let bySheet = at.get(f.s)
    if (!bySheet) at.set(f.s, (bySheet = new Map()))
    let col = bySheet.get(f.c)
    if (!col) bySheet.set(f.c, (col = []))
    col.push({ r: f.r, id })
  })
  for (const bySheet of at.values()) for (const col of bySheet.values()) col.sort((a, b) => a.r - b.r)
  const lowerBound = (col: { r: number }[], r: number) => {
    let lo = 0
    let hi = col.length
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (col[mid].r < r) lo = mid + 1
      else hi = mid
    }
    return lo
  }
  const succ: number[][] = formulas.map(() => [])
  const indeg = new Array<number>(formulas.length).fill(0)
  const listeners: CState['listeners'] = new Map()
  const usedCols = sheets.map((sh, s) => Math.max(sh.cols.length, ...[...(at.get(s)?.keys() ?? [])].map((c) => c + 1)))
  formulas.forEach((f, id) => {
    for (const g of f.refs) {
      if (g.s < 0) continue
      let ls = listeners.get(g.s)
      if (!ls) listeners.set(g.s, (ls = new Map()))
      const maxCol = Math.min(g.c2, usedCols[g.s] - 1)
      for (let c = g.c1; c <= maxCol; c++) {
        let lc = ls.get(c)
        if (!lc) ls.set(c, (lc = []))
        lc.push([g.r1, g.r2, id])
        const col = at.get(g.s)?.get(c)
        if (!col) continue
        for (let i = lowerBound(col, g.r1); i < col.length && col[i].r <= g.r2; i++) {
          succ[col[i].id].push(id)
          indeg[id]++
        }
      }
    }
  })
  const order: number[] = []
  const queue: number[] = []
  indeg.forEach((d, id) => d === 0 && queue.push(id))
  while (queue.length) {
    const id = queue.pop()!
    order.push(id)
    for (const n of succ[id]) if (--indeg[n] === 0) queue.push(n)
  }
  const topo = new Array<number>(formulas.length).fill(-1)
  order.forEach((id, i) => (topo[id] = i))
  const t2 = clockMs()

  const state: CState = {
    book,
    opts,
    prep,
    ctx,
    fns,
    formulas,
    topo,
    order,
    succ,
    listeners,
    unsupported,
    invalid: prep.invalid,
    ms: { parse: t1 - t0, build: t2 - t1, evaluate: 0 },
    generatedChars: src.length,
  }
  const t3 = clockMs()
  evaluate(state, order)
  // ячейки цикла и всё, что от них зависит
  topo.forEach((p, id) => {
    if (p < 0) put(sheets[formulas[id].s], formulas[id].r, formulas[id].c, E.CYCLE)
  })
  state.ms.evaluate = clockMs() - t3
  return state
}

function normalize(x: V | Rng, scalar: (x: V | Rng) => V): V {
  const v = x instanceof Rng ? scalar(x) : x
  if (typeof v === 'number' && !Number.isFinite(v)) return E.NUM
  // формула, которая ссылается на пустую ячейку, даёт 0, как в Excel
  return v === undefined || v === null ? 0 : v
}

const runtimeCache = new WeakMap<RuntimeContext, ReturnType<typeof makeRuntime>>()
function rtOf(ctx: RuntimeContext) {
  let rt = runtimeCache.get(ctx)
  if (!rt) runtimeCache.set(ctx, (rt = makeRuntime(ctx)))
  return rt
}

/** Пересчёт формул в данном порядке; число изменившихся значений. */
function evaluate(state: CState, ids: number[]): number {
  const { ctx, fns, formulas } = state
  const rt = rtOf(ctx)
  let changed = 0
  for (const id of ids) {
    const f = formulas[id]
    const sh = ctx.sheets[f.s]
    const old = sh.cols[f.c]?.[f.r] ?? null
    const v = normalize(fns[id](), rt.scalar)
    if (!same(old, v)) {
      changed++
      put(sh, f.r, f.c, v)
      invalidate(ctx, f.s, f.r, f.c)
    }
  }
  return changed
}

function same(a: V, b: V): boolean {
  if (a instanceof CErr || b instanceof CErr) return a instanceof CErr && b instanceof CErr && a.code === b.code
  return Object.is(a, b) || a === b
}

/** Сбросить кэши больших диапазонов и индексов, которые задевает ячейка. */
function invalidate(ctx: RuntimeContext, s: number, r: number, c: number) {
  for (const cache of [ctx.flat, ctx.index] as Map<string, unknown>[]) {
    if (cache.size === 0) continue
    for (const key of cache.keys()) {
      const [ks, r1, c1, r2, c2] = key.split(':').map(Number)
      if (ks === s && r >= r1 && r <= r2 && c >= c1 && c <= c2) cache.delete(key)
    }
  }
}

/** Правка значений атомарно, один пересчёт зависимых; правка формулы пересобирает книгу. */
export function setCells(state: CState, edits: CellEdit[]): { state: CState; changed: number } {
  const sheetIdx = (s: string) => state.prep.sheets.indexOf(state.prep.sheetOf(s))
  const isFormula = new Set(state.formulas.map((f) => `${f.s}:${f.r}:${f.c}`))
  if (edits.some((e) => e.input.formula !== undefined || isFormula.has(`${sheetIdx(e.sheet)}:${parseAddress(e.addr).row}:${parseAddress(e.addr).col}`))) {
    // прототип: структурная правка пересобирает книгу целиком
    const book = applyToBook(state.book, edits)
    const next = compileBook(book, state.opts)
    return { state: next, changed: -1 }
  }
  const dirty = new Set<number>()
  let changed = 0
  for (const e of edits) {
    const s = sheetIdx(e.sheet)
    const { row, col } = parseAddress(e.addr)
    const sh = state.ctx.sheets[s]
    const v = inputScalar(e.input.value)
    const old = sh.cols[col]?.[row] ?? null
    if (same(old, v)) continue
    changed++
    put(sh, row, col, v)
    invalidate(state.ctx, s, row, col)
    for (const [r1, r2, id] of state.listeners.get(s)?.get(col) ?? []) if (row >= r1 && row <= r2) dirty.add(id)
  }
  state.book = applyToBook(state.book, edits)
  // замыкание по графу и пересчёт в топологическом порядке
  const stack = [...dirty]
  while (stack.length) {
    const id = stack.pop()!
    for (const n of state.succ[id])
      if (!dirty.has(n)) {
        dirty.add(n)
        stack.push(n)
      }
  }
  const ids = [...dirty].filter((id) => state.topo[id] >= 0).sort((a, b) => state.topo[a] - state.topo[b])
  changed += evaluate(state, ids)
  return { state, changed }
}

/** Вход книги с правками. Сессия владеет своей копией (`ownBook`), поэтому правит её на месте. */
function applyToBook(book: BookInput, edits: CellEdit[]): BookInput {
  for (const e of edits) {
    const sh = book.sheets.find((s) => s.name === e.sheet)
    if (sh) sh.cells[e.addr] = e.input
  }
  return book
}

export function ownBook(book: BookInput): BookInput {
  return { sheets: book.sheets.map((s) => ({ name: s.name, cells: { ...s.cells } })), names: book.names }
}

export function readValues(state: CState): Map<string, Map<string, CellValue>> {
  const out = new Map<string, Map<string, CellValue>>()
  for (const sh of state.ctx.sheets) {
    const cells = new Map<string, CellValue>()
    sh.cols.forEach((col, c) => {
      if (!col) return
      col.forEach((v, r) => {
        if (v === null || v === undefined) return
        cells.set(formatAddress(r, c), v instanceof CErr ? { error: v.code } : v)
      })
    })
    out.set(sh.name, cells)
  }
  return out
}
