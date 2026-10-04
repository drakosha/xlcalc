/**
 * Движок: формулы листа компилируются в JS, ячейки считаются по графу ссылок.
 *
 * Одна `new Function` на книгу, каждая формула своей функцией (правка формулы в сессии компилирует
 * только её). Код генерирует только пакет: литералы через `JSON.stringify`, вызываются только функции
 * рантайма из белого списка, листы и диапазоны переданы числами. Полный расчёт идёт по
 * топологическому порядку; правка значения пересчитывает только зависимые ячейки.
 *
 * Без импорта node:* (хэши считает `index.ts`), чтобы движок можно было исполнить и вне Node.
 */
import type { BookInput, CellEdit, CellInput, CellValue, CheckCode, ErrorCode, Externals, NameInput, NameIssue } from './contract'
import { FUNCTIONS } from './functions'
import { compareAddress, formatAddress, indexToCol, inputScalar, MAX_COLS, MAX_ROWS, parseAddress } from './model'
import { CErr, E, isErr, makeRuntime, Rng, type Runtime, type RuntimeContext, type SheetData, type V } from './runtime'
import { canonical, FormulaError, parse, quoteSheet, sheetKey, validate, validateName, type Ast } from './syntax'

interface Formula {
  id: number
  s: number
  r: number
  c: number
  /** Канонический текст. */
  text: string
  refs: Rng[]
  /** Прямой вызов волатильной функции. */
  volatile: boolean
  source: string
  fn: () => V | Rng
}

interface Invalid {
  code: CheckCode
  message: string
  at: number
  error: ErrorCode
}

interface PreparedName {
  name: string
  /** Индекс листа области видимости; -1 у глобального. */
  scope: number
  ast: Extract<Ast, { k: 'ref' }>
  area: boolean
}

export interface EngineOptions {
  /** Момент часов, серийный номер (UTC). */
  now: number
  seed: number
  externals?: Externals
}

export interface EngineState {
  input: BookInput
  sheetNames: string[]
  sheetIdx: Map<string, number>
  names: PreparedName[]
  nameIssues: NameIssue[]
  externalFunctions: string[]
  ctx: RuntimeContext
  rt: Runtime
  /** id → формула; после удаления дыра. */
  formulas: (Formula | undefined)[]
  /** `s:r:c` → id формулы. */
  at: Map<string, number>
  invalid: Map<string, Invalid>
  succ: number[][]
  /** Позиция в топологическом порядке; -1 у ячеек цикла и зависящих от них. */
  topo: number[]
  order: number[]
  /** Слушатели узких диапазонов: лист → колонка → [r1, r2, id]. */
  listeners: Map<number, Map<number, [number, number, number][]>>
  /** Слушатели широких диапазонов (целые строки и больше 64 колонок): лист → [r1, r2, c1, c2, id]. */
  wide: Map<number, [number, number, number, number, number][]>
  /** Формулы, зависящие от волатильных функций (прямо или через другие). */
  volatile: Set<number>
}

const key = (s: number, r: number, c: number) => `${s}:${r}:${c}`
const WIDE = 64

function emptySheet(name: string): SheetData {
  return { name, cols: [], rows: 0 }
}

function put(sh: SheetData, r: number, c: number, v: V) {
  ;(sh.cols[c] ??= [])[r] = v
  if (r + 1 > sh.rows) sh.rows = r + 1
}

export function invalidError(code: CheckCode): ErrorCode {
  if (code === 'unknown_function' || code === 'unknown_name') return 'NAME'
  if (code === 'unknown_sheet' || code === 'external_ref') return 'REF'
  return 'ERROR'
}

/** Текст ссылки для `explain` и причин: `'Лист'!B3`, `'Лист'!D:D`, `'Лист'!A1:C9`. */
export function rngText(sheetNames: string[], g: Rng): string {
  const sheet = quoteSheet(sheetNames[g.s])
  if (g.r1 === 0 && g.r2 === MAX_ROWS - 1) return `${sheet}!${indexToCol(g.c1)}:${indexToCol(g.c2)}`
  if (g.c1 === 0 && g.c2 === MAX_COLS - 1) return `${sheet}!${g.r1 + 1}:${g.r2 + 1}`
  const a = formatAddress(g.r1, g.c1)
  return g.r1 === g.r2 && g.c1 === g.c2 ? `${sheet}!${a}` : `${sheet}!${a}:${formatAddress(g.r2, g.c2)}`
}

// ---------- подготовка ----------

function prepareNames(state: Pick<EngineState, 'sheetIdx' | 'sheetNames'>, names: NameInput[]): { names: PreparedName[]; issues: NameIssue[] } {
  const out: PreparedName[] = []
  const issues: NameIssue[] = []
  const seen = new Set<string>()
  for (const n of names) {
    const issue = (reason: string) => issues.push({ name: n.name, ...(n.scope ? { scope: n.scope } : {}), reason })
    const ok = validateName(n.name)
    if (!ok.ok) {
      issue(ok.message)
      continue
    }
    const scope = n.scope === undefined ? -1 : (state.sheetIdx.get(sheetKey(n.scope)) ?? -2)
    if (scope === -2) {
      issue(`unknown scope sheet ${n.scope}`)
      continue
    }
    const dup = `${scope}:${n.name.toUpperCase()}`
    if (seen.has(dup)) {
      issue('duplicate name')
      continue
    }
    seen.add(dup)
    let ast: Ast
    try {
      ast = parse(n.ref)
    } catch (e) {
      if (!(e instanceof FormulaError)) throw e
      issue(`bad reference: ${e.message}`)
      continue
    }
    if (ast.k !== 'ref' || ast.sheet === null) {
      issue('a name must reference a cell or a range with a sheet')
      continue
    }
    if (!state.sheetIdx.has(sheetKey(ast.sheet))) {
      issue(`unknown sheet ${ast.sheet}`)
      continue
    }
    const ref = ast.ref
    out.push({ name: n.name, scope, ast, area: ref.kind !== 'cell' && !(ref.r1 === ref.r2 && ref.c1 === ref.c2) })
  }
  return { names: out, issues }
}

function nameFor(state: EngineState, here: number, name: string): PreparedName | undefined {
  const up = name.toUpperCase()
  return state.names.find((n) => n.scope === here && n.name.toUpperCase() === up) ?? state.names.find((n) => n.scope === -1 && n.name.toUpperCase() === up)
}

/** Код JS одной формулы (самодостаточное выражение-фабрика) и её ссылки. */
function codegen(state: EngineState, ast: Ast, s: number, r: number, c: number): { source: string; refs: Rng[]; volatile: boolean } {
  const refs: Rng[] = []
  const hoisted: Rng[] = []
  let volatile = false
  const sheetOf = (name: string | null) => (name === null ? s : state.sheetIdx.get(sheetKey(name))!)
  const deref = (n: Ast): Ast => (n.k === 'name' ? (nameFor(state, s, n.name)?.ast ?? n) : n)
  // rangeOk: место, где функция ждёт диапазон (туда идёт объект диапазона, в том числе из одной ячейки)
  const go = (n: Ast, rangeOk = false): string => {
    switch (n.k) {
      case 'num':
      case 'str':
      case 'bool':
        return JSON.stringify(n.v)
      case 'err':
        return `E.${n.code}`
      case 'ref': {
        const g = new Rng(sheetOf(n.sheet), n.ref.r1, n.ref.c1, n.ref.r2, n.ref.c2)
        refs.push(g)
        if (n.ref.kind === 'cell' && !rangeOk) return `cell(${g.s},${g.r1},${g.c1})`
        hoisted.push(g)
        const h = `g${hoisted.length - 1}`
        return rangeOk ? h : `scalar(${h})`
      }
      case 'name': {
        const d = deref(n)
        return d.k === 'name' ? 'E.NAME' : go(d, rangeOk)
      }
      case 'paren':
        return rangeOk ? go(n.a, true) : `(${go(n.a)})`
      case 'un':
        return n.op === '-' ? `neg(${go(n.a)})` : go(n.a)
      case 'pct':
        return `pct(${go(n.a)})`
      case 'bin': {
        const a = go(n.a)
        const b = go(n.b)
        const op = ({ '+': 'add', '-': 'sub', '*': 'mul', '/': 'div', '^': 'pow', '&': 'cat' } as Record<string, string>)[n.op]
        return op ? `${op}(${a},${b})` : `cmp(${JSON.stringify(n.op)},${a},${b})`
      }
      case 'fn': {
        const spec = FUNCTIONS[n.name]
        if (!spec) return `ext(${JSON.stringify(canonical(n))})`
        if (spec.volatile) volatile = true
        const args = n.args
        if (n.name === 'IF') return `IF(${go(args[0])},()=>${go(args[1])},()=>${args[2] ? go(args[2]) : 'false'})`
        if (n.name === 'IFERROR') return `IFERR(()=>${go(args[0])},()=>${go(args[1])})`
        if (n.name === 'RAND') return `rand(${s},${r},${c})`
        if (n.name === 'RANDBETWEEN') return `randbetween(${s},${r},${c},${go(args[0])},${go(args[1])})`
        return `F.${n.name}(${args.map((a, i) => go(a, spec.range(i))).join(',')})`
      }
    }
  }
  const expr = go(ast)
  const decl = hoisted.length ? `const ${hoisted.map((g, i) => `g${i}=R(${g.s},${g.r1},${g.c1},${g.r2},${g.c2})`).join(',')};` : ''
  return { source: `(()=>{${decl}return ()=>${expr}})()`, refs, volatile }
}

const PRELUDE = '"use strict";const {cell,scalar,add,sub,mul,div,pow,neg,pct,cat,cmp,F,IF,IFERR,rand,randbetween,ext,E,R}=rt;'

function compileSources(rt: Runtime, sources: string[]): (() => V | Rng)[] {
  const R = (s: number, r1: number, c1: number, r2: number, c2: number) => new Rng(s, r1, c1, r2, c2)
  return new Function('rt', `${PRELUDE}return [${sources.join(',\n')}]`)({ ...rt, R }) as (() => V | Rng)[]
}

type Prepared = { ok: true; text: string; ast: Ast } | { ok: false; invalid: Invalid }

function prepare(state: EngineState, s: number, text: string): Prepared {
  const visible = state.names.filter((n) => n.scope === -1 || n.scope === s).map((n) => n.name)
  const r = validate(text, { sheets: state.sheetNames, names: visible, externalFunctions: state.externalFunctions }, (name) => nameFor(state, s, name)?.area)
  if (!r.ok) return { ok: false, invalid: { code: r.code, message: r.message, at: r.at, error: invalidError(r.code) } }
  return { ok: true, text: canonical(r.ast, (sh) => state.sheetNames[state.sheetIdx.get(sheetKey(sh))!]), ast: r.ast }
}

// ---------- граф ----------

function rebuildGraph(state: EngineState) {
  const { formulas } = state
  const at = new Map<number, Map<number, { r: number; id: number }[]>>()
  for (const f of formulas) {
    if (!f) continue
    let bySheet = at.get(f.s)
    if (!bySheet) at.set(f.s, (bySheet = new Map()))
    let col = bySheet.get(f.c)
    if (!col) bySheet.set(f.c, (col = []))
    col.push({ r: f.r, id: f.id })
  }
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
  const listeners: EngineState['listeners'] = new Map()
  const wide: EngineState['wide'] = new Map()
  for (const f of formulas) {
    if (!f) continue
    for (const g of f.refs) {
      if (g.c2 - g.c1 >= WIDE) {
        let w = wide.get(g.s)
        if (!w) wide.set(g.s, (w = []))
        w.push([g.r1, g.r2, g.c1, g.c2, f.id])
      } else {
        let ls = listeners.get(g.s)
        if (!ls) listeners.set(g.s, (ls = new Map()))
        for (let c = g.c1; c <= g.c2; c++) {
          let lc = ls.get(c)
          if (!lc) ls.set(c, (lc = []))
          lc.push([g.r1, g.r2, f.id])
        }
      }
      const bySheet = at.get(g.s)
      if (!bySheet) continue
      for (const [c, col] of bySheet) {
        if (c < g.c1 || c > g.c2) continue
        for (let i = lowerBound(col, g.r1); i < col.length && col[i].r <= g.r2; i++) {
          succ[col[i].id].push(f.id)
          indeg[f.id]++
        }
      }
    }
  }
  const order: number[] = []
  const queue: number[] = []
  formulas.forEach((f, id) => f && indeg[id] === 0 && queue.push(id))
  // порядок очереди по адресу не нужен: значения от порядка вычисления не зависят
  while (queue.length) {
    const id = queue.pop()!
    order.push(id)
    for (const n of succ[id]) if (--indeg[n] === 0) queue.push(n)
  }
  const topo = new Array<number>(formulas.length).fill(-1)
  order.forEach((id, i) => (topo[id] = i))
  state.succ = succ
  state.topo = topo
  state.order = order
  state.listeners = listeners
  state.wide = wide
  const vol = new Set<number>()
  const stack = formulas.filter((f): f is Formula => !!f?.volatile).map((f) => f.id)
  while (stack.length) {
    const id = stack.pop()!
    if (vol.has(id)) continue
    vol.add(id)
    stack.push(...succ[id])
  }
  state.volatile = vol
}

/** Формулы, читающие ячейку (по ссылкам, независимо от того, есть ли в ней что-то). */
function listenersOf(state: EngineState, s: number, r: number, c: number, out: Set<number>) {
  for (const [r1, r2, id] of state.listeners.get(s)?.get(c) ?? []) if (r >= r1 && r <= r2) out.add(id)
  for (const [r1, r2, c1, c2, id] of state.wide.get(s) ?? []) if (r >= r1 && r <= r2 && c >= c1 && c <= c2) out.add(id)
}

// ---------- вычисление ----------

function same(a: V, b: V): boolean {
  if (isErr(a) || isErr(b)) return isErr(a) && isErr(b) && a.code === b.code
  return Object.is(a, b) || a === b
}

/** Сбросить кэши больших диапазонов и индексов, которые задевает ячейка. */
function invalidate(ctx: RuntimeContext, s: number, r: number, c: number) {
  for (const cache of [ctx.flat, ctx.index] as Map<string, unknown>[]) {
    if (cache.size === 0) continue
    for (const k of cache.keys()) {
      const [ks, r1, c1, r2, c2] = k.split(':').map(Number)
      if (ks === s && r >= r1 && r <= r2 && c >= c1 && c <= c2) cache.delete(k)
    }
  }
}

type Tracker = Map<string, V>

function write(state: EngineState, s: number, r: number, c: number, v: V, track?: Tracker): boolean {
  const sh = state.ctx.sheets[s]
  const old = sh.cols[c]?.[r] ?? null
  if (same(old, v)) return false
  const k = key(s, r, c)
  if (track && !track.has(k)) track.set(k, old)
  put(sh, r, c, v)
  invalidate(state.ctx, s, r, c)
  return true
}

function evaluate(state: EngineState, ids: number[], track?: Tracker) {
  const { rt } = state
  for (const id of ids) {
    const f = state.formulas[id]!
    let v: V
    if (state.topo[id] < 0) v = E.CYCLE
    else {
      const x = f.fn()
      v = x instanceof Rng ? rt.scalar(x) : x
      // формула со ссылкой на пустую ячейку даёт 0, как в Excel
      if (v === null || v === undefined) v = 0
      else if (typeof v === 'number' && !Number.isFinite(v)) v = E.NUM
    }
    write(state, f.s, f.r, f.c, v, track)
  }
}

function evaluateAll(state: EngineState, track?: Tracker) {
  evaluate(state, state.order, track)
  const cyc = state.formulas.filter((f): f is Formula => !!f && state.topo[f.id] < 0).map((f) => f.id)
  evaluate(state, cyc, track)
}

// ---------- сборка ----------

export function createState(book: BookInput, opts: EngineOptions): EngineState {
  const sheetNames = book.sheets.map((s) => s.name)
  const sheetIdx = new Map(sheetNames.map((n, i) => [sheetKey(n), i]))
  if (sheetIdx.size !== sheetNames.length) throw new Error('duplicate sheet name')
  const input: BookInput = { sheets: book.sheets.map((s) => ({ name: s.name, cells: { ...s.cells } })), names: [...(book.names ?? [])] }
  const { names, issues } = prepareNames({ sheetIdx, sheetNames }, input.names!)
  const ctx: RuntimeContext = { sheets: sheetNames.map(emptySheet), now: opts.now, seed: opts.seed, externals: opts.externals?.values ?? {}, flat: new Map(), index: new Map() }
  const state: EngineState = {
    input,
    sheetNames,
    sheetIdx,
    names,
    nameIssues: issues,
    externalFunctions: (opts.externals?.functions ?? []).map((f) => f.toUpperCase()).filter((f) => !FUNCTIONS[f]),
    ctx,
    rt: makeRuntime(ctx),
    formulas: [],
    at: new Map(),
    invalid: new Map(),
    succ: [],
    topo: [],
    order: [],
    listeners: new Map(),
    wide: new Map(),
    volatile: new Set(),
  }
  const pending: Omit<Formula, 'fn'>[] = []
  input.sheets.forEach((sheet, s) => {
    for (const [addr, cellIn] of Object.entries(sheet.cells)) {
      const { row, col } = parseAddress(addr)
      if (cellIn.formula === undefined) {
        const v = inputScalar(cellIn.value)
        if (v !== null) put(ctx.sheets[s], row, col, v)
        continue
      }
      const p = prepare(state, s, cellIn.formula)
      if (!p.ok) {
        state.invalid.set(key(s, row, col), p.invalid)
        put(ctx.sheets[s], row, col, E[p.invalid.error])
        continue
      }
      const g = codegen(state, p.ast, s, row, col)
      pending.push({ id: pending.length, s, r: row, c: col, text: p.text, refs: g.refs, volatile: g.volatile, source: g.source })
    }
  })
  const fns = compileSources(state.rt, pending.map((p) => p.source))
  state.formulas = pending.map((p, i) => ({ ...p, fn: fns[i] }))
  for (const f of state.formulas) state.at.set(key(f!.s, f!.r, f!.c), f!.id)
  rebuildGraph(state)
  evaluateAll(state)
  return state
}

/**
 * Правки атомарно, один пересчёт зависимых; изменившиеся ячейки (`s:r:c`). `withVolatile`: часы
 * сменились (живая сессия), волатильные ячейки пересчитываются в том же проходе.
 */
export function applyEdits(state: EngineState, edits: CellEdit[], withVolatile = false): string[] {
  const track: Tracker = new Map()
  const dirty = new Set<number>()
  const fresh: Formula[] = []
  let structural = false
  for (const e of edits) {
    const s = state.sheetIdx.get(sheetKey(e.sheet))
    if (s === undefined) throw new Error(`unknown sheet ${e.sheet}`)
    const { row, col } = parseAddress(e.cell)
    const k = key(s, row, col)
    const cells = state.input.sheets[s].cells
    if (e.input === null) delete cells[e.cell]
    else cells[e.cell] = e.input
    const old = state.at.get(k)
    if (old !== undefined) {
      state.formulas[old] = undefined
      state.at.delete(k)
      structural = true
    }
    state.invalid.delete(k)
    listenersOf(state, s, row, col, dirty)
    const input: CellInput | null = e.input
    if (input?.formula !== undefined) {
      const p = prepare(state, s, input.formula)
      if (!p.ok) {
        state.invalid.set(k, p.invalid)
        write(state, s, row, col, E[p.invalid.error], track)
        continue
      }
      const g = codegen(state, p.ast, s, row, col)
      const f = { id: state.formulas.length, s, r: row, c: col, text: p.text, refs: g.refs, volatile: g.volatile, source: g.source, fn: () => null as V }
      state.formulas.push(f)
      state.at.set(k, f.id)
      fresh.push(f)
      structural = true
    } else write(state, s, row, col, input ? inputScalar(input.value) : null, track)
  }
  if (fresh.length) {
    const fns = compileSources(state.rt, fresh.map((f) => f.source))
    fresh.forEach((f, i) => (f.fn = fns[i]))
  }
  if (structural) rebuildGraph(state)
  for (const f of fresh) dirty.add(f.id)
  if (withVolatile) for (const id of state.volatile) dirty.add(id)
  // ячейки, бывшие в цикле, пересчитываются, если цикл разорван
  if (structural) for (const f of state.formulas) if (f && isErr(state.ctx.sheets[f.s].cols[f.c]?.[f.r]) && (state.ctx.sheets[f.s].cols[f.c]![f.r] as CErr).code === 'CYCLE') dirty.add(f.id)
  recompute(state, dirty, track)
  return changed(state, track)
}

/** Пересчёт множества формул и всего, что от них зависит, в топологическом порядке. */
function recompute(state: EngineState, dirty: Set<number>, track: Tracker) {
  const stack = [...dirty]
  while (stack.length) {
    const id = stack.pop()!
    for (const n of state.succ[id] ?? [])
      if (!dirty.has(n)) {
        dirty.add(n)
        stack.push(n)
      }
  }
  const live = [...dirty].filter((id) => state.formulas[id])
  const ordered = live.filter((id) => state.topo[id] >= 0).sort((a, b) => state.topo[a] - state.topo[b])
  evaluate(state, ordered, track)
  evaluate(state, live.filter((id) => state.topo[id] < 0), track)
}

function changed(state: EngineState, track: Tracker): string[] {
  const out: string[] = []
  for (const [k, old] of track) {
    const [s, r, c] = k.split(':').map(Number)
    if (!same(old, state.ctx.sheets[s].cols[c]?.[r] ?? null)) out.push(k)
  }
  return out
}

/** Новые часы: пересчёт волатильных ячеек и зависящих от них. */
export function applyClock(state: EngineState, now: number, seed: number): string[] {
  state.ctx.now = now
  state.ctx.seed = seed
  const track: Tracker = new Map()
  recompute(state, new Set(state.volatile), track)
  return changed(state, track)
}

/** Значение ячейки (пустая это null). */
export function valueAt(state: EngineState, s: number, r: number, c: number): V {
  return state.ctx.sheets[s].cols[c]?.[r] ?? null
}

export function toCellValue(v: V): CellValue {
  return isErr(v) ? { error: v.code } : v
}

// ---------- причины ошибок ----------

const GENERIC: Record<ErrorCode, string> = {
  DIV0: 'division by zero',
  NA: 'value not found',
  NAME: 'unknown function or name',
  NUM: 'number out of range',
  REF: 'invalid reference',
  VALUE: 'wrong type of value',
  CYCLE: 'circular reference',
  SPILL: 'array result does not fit',
  ERROR: 'error',
}

export interface Issue {
  s: number
  r: number
  c: number
  error: ErrorCode
  reason: string
  check?: CheckCode
  at?: number
}

/** Ячейки с ошибкой и причина: проверка формулы, цикл, ошибка от другой ячейки или своя. */
export function issues(state: EngineState): Issue[] {
  const out: Issue[] = []
  for (const [k, inv] of state.invalid) {
    const [s, r, c] = k.split(':').map(Number)
    out.push({ s, r, c, error: inv.error, reason: inv.message, check: inv.code, at: inv.at })
  }
  for (const f of state.formulas) {
    if (!f) continue
    const v = valueAt(state, f.s, f.r, f.c)
    if (!isErr(v)) continue
    let reason = GENERIC[v.code]
    if (v.code !== 'CYCLE') {
      const from = errorSource(state, f, v.code)
      if (from) reason = `${GENERIC[v.code]}, from ${from}`
      else if (v.code === 'NA' && /ext\(/.test(f.source)) reason = 'value not found (or no value for an external call)'
    }
    out.push({ s: f.s, r: f.r, c: f.c, error: v.code, reason })
  }
  return out.sort((a, b) => a.s - b.s || a.r - b.r || a.c - b.c)
}

function errorSource(state: EngineState, f: Formula, code: ErrorCode): string | null {
  for (const g of f.refs) {
    const flat = state.rt.flat(g)
    const b = state.rt.bounds(g)
    for (let i = 0; i < flat.length; i++) {
      const v = flat[i]
      if (isErr(v) && v.code === code) return `${quoteSheet(state.sheetNames[g.s])}!${formatAddress(b.r1 + Math.floor(i / b.w), b.c1 + (i % b.w))}`
    }
  }
  return null
}

// ---------- отладка и explain ----------

/** Сгенерированный код в порядке вычисления, для чтения глазами (не контракт). */
export function debugSource(state: EngineState): string {
  const lines = [PRELUDE]
  const ids = [...state.order, ...state.formulas.filter((f) => f && state.topo[f.id] < 0).map((f) => f!.id)]
  for (const id of ids) {
    const f = state.formulas[id]!
    lines.push(`// ${quoteSheet(state.sheetNames[f.s])}!${formatAddress(f.r, f.c)} = ${f.text}${state.topo[id] < 0 ? '  (цикл)' : ''}`, f.source)
  }
  return lines.join('\n')
}

export interface ExplainTree {
  ref: string
  value?: CellValue
  formula?: string
  cells?: number
  children: ExplainTree[]
}

export function explainCell(state: EngineState, s: number, r: number, c: number, depth: number, maxNodes: number, expand: boolean): { root: ExplainTree; nodes: number; truncated: boolean } {
  let nodes = 0
  let truncated = false
  const visit = (s: number, r: number, c: number, d: number): ExplainTree => {
    nodes++
    const id = state.at.get(key(s, r, c))
    const f = id === undefined ? undefined : state.formulas[id]
    const node: ExplainTree = { ref: `${quoteSheet(state.sheetNames[s])}!${formatAddress(r, c)}`, value: toCellValue(valueAt(state, s, r, c)), children: [] }
    const inv = state.invalid.get(key(s, r, c))
    if (inv) node.formula = state.input.sheets[s].cells[formatAddress(r, c)]?.formula
    if (!f) return node
    node.formula = f.text
    if (d <= 0) {
      if (f.refs.length) truncated = true
      return node
    }
    for (const g of f.refs) {
      if (nodes >= maxNodes) {
        truncated = true
        break
      }
      if (g.r1 === g.r2 && g.c1 === g.c2) {
        node.children.push(visit(g.s, g.r1, g.c1, d - 1))
        continue
      }
      const b = state.rt.bounds(g)
      const flat = state.rt.flat(g)
      const filled: number[] = []
      flat.forEach((v, i) => v !== null && filled.push(i))
      nodes++
      const rn: ExplainTree = { ref: rngText(state.sheetNames, g), cells: filled.length, children: [] }
      node.children.push(rn)
      if (!expand) continue
      for (const i of filled) {
        if (nodes >= maxNodes) {
          truncated = true
          break
        }
        rn.children.push(visit(g.s, b.r1 + Math.floor(i / b.w), b.c1 + (i % b.w), d - 2))
      }
    }
    return node
  }
  const root = visit(s, r, c, depth)
  return { root, nodes, truncated }
}

export function explainText(root: ExplainTree): string {
  const fmt = (v: CellValue | undefined) => (v === undefined ? '' : v !== null && typeof v === 'object' ? `#${v.error}` : JSON.stringify(v))
  const lines: string[] = []
  const go = (n: ExplainTree, indent: string) => {
    const head = n.cells !== undefined ? `${n.ref} (${n.cells} cells)` : `${n.ref} = ${fmt(n.value)}`
    lines.push(`${indent}${head}${n.formula ? `  ← ${n.formula}` : ''}`)
    for (const ch of n.children) go(ch, indent + '  ')
  }
  go(root, '')
  return lines.join('\n')
}

/** Ячейки книги по адресу (для выдачи и хэша). */
export function sortedCells(state: EngineState, s: number): string[] {
  const addrs = new Set(Object.keys(state.input.sheets[s].cells))
  return [...addrs].sort(compareAddress)
}
