/**
 * Вариант H: минимальный адаптер HyperFormula 3.4.0 на замороженной конфигурации шага 2.
 * Особенности 3.4.0 закрывает наш канонизатор (`toHfText`) и запись значений; см. `hf-facts.test.ts`.
 */
import { CellError, DetailedCellError, EmptyValue, ErrorType, FunctionArgumentType, FunctionPlugin, HyperFormula, SimpleRangeValue, type ConfigParams, type RawCellContent } from 'hyperformula'
import { buildIndex, parseCrit, select, type CritColumn, type CritIndex, type CritValue } from './criteria'
import { inputHash, valuesHash, type BookValues } from './canon'
import type { CellEdit, ProbeEngine, ProbeResult, ProbeSession } from './engine'
import { nowSerial, randAt, randBetween, VOLATILE, WHITELIST } from './functions'
import { formatAddress, inputScalar, parseAddress, type BookInput, type CellInput, type CellValue, type Clock, type ErrorCode } from './model'
import { toHfText, validateFormula, walk } from './parser'
import { invalidError, prepareBook, type PreparedBook } from './prep'

interface VolatileContext {
  now: number
  seed: number
  sheetNames: string[]
  /** Своя SUMIFS/COUNTIFS с индексом (вариант H со своей SUMIFS), иначе встроенная HF. */
  ownIfs: boolean
  /** Плоские значения и индексы диапазонов; сбрасываются при каждой правке. */
  flat: Map<string, CritValue[]>
  index: Map<string, CritIndex>
}

/**
 * Замороженная конфигурация (причины в `docs/probes/WB1-engine.md`): английские имена и разделители,
 * даты 1900 с ложным 29.02, строки только ISO-даты, регистр не важен, «й» ≠ «и», шаблоны `*`/`?`/`~`
 * без регулярных выражений, массивы растекаются, точность выдачи 14 знаков.
 */
export const HF_CONFIG: Partial<ConfigParams> = {
  licenseKey: 'gpl-v3',
  language: 'enGB',
  functionArgSeparator: ',',
  decimalSeparator: '.',
  thousandSeparator: '',
  nullDate: { year: 1899, month: 12, day: 31 },
  leapYear1900: true,
  dateFormats: ['YYYY-MM-DD'],
  timeFormats: ['hh:mm', 'hh:mm:ss.sss'],
  caseSensitive: false,
  accentSensitive: true,
  localeLang: 'ru-RU',
  useWildcards: true,
  useRegularExpressions: false,
  matchWholeCell: true,
  evaluateNullToZero: true,
  useArrayArithmetic: true,
  ignoreWhiteSpace: 'any',
  smartRounding: true,
  precisionRounding: 14,
  maxRows: 120_000,
  maxColumns: 16_384,
  undoLimit: 0,
}

export const HF_CONFIG_VERSION = `hyperformula@${HyperFormula.version};${JSON.stringify(HF_CONFIG)}`

/** Волатильные функции от момента и зерна из `config.context` вместо часов и Math.random. */
class ClockPlugin extends FunctionPlugin {
  private get clock(): VolatileContext {
    return this.config.context as VolatileContext
  }
  private sheetName(sheet: number): string {
    return this.clock.sheetNames[sheet] ?? String(sheet)
  }
  today(ast: any, state: any) {
    return this.runFunction(ast.args, state, this.metadata('TODAY'), () => Math.floor(this.clock.now))
  }
  now(ast: any, state: any) {
    return this.runFunction(ast.args, state, this.metadata('NOW'), () => this.clock.now)
  }
  rand(ast: any, state: any) {
    const a = state.formulaAddress
    return this.runFunction(ast.args, state, this.metadata('RAND'), () => randAt(this.clock.seed, this.sheetName(a.sheet), a.row, a.col))
  }
  randbetween(ast: any, state: any) {
    const a = state.formulaAddress
    return this.runFunction(ast.args, state, this.metadata('RANDBETWEEN'), (lo: number, hi: number) =>
      randBetween(this.clock.seed, this.sheetName(a.sheet), a.row, a.col, lo, hi),
    )
  }
}
;(ClockPlugin as unknown as { implementedFunctions: unknown }).implementedFunctions = {
  TODAY: { method: 'today', parameters: [], isVolatile: true },
  NOW: { method: 'now', parameters: [], isVolatile: true },
  RAND: { method: 'rand', parameters: [], isVolatile: true },
  RANDBETWEEN: {
    method: 'randbetween',
    parameters: [{ argumentType: FunctionArgumentType.NUMBER }, { argumentType: FunctionArgumentType.NUMBER }],
    isVolatile: true,
  },
}
for (const fn of VOLATILE) HyperFormula.registerFunction(fn, ClockPlugin as never)

/**
 * SUMIFS/COUNTIFS с индексом по критериям (общим с C) через FunctionPlugin. Наследует встроенный
 * плагин: регистрация функций в HF глобальна на процесс, поэтому без `ownIfs` зовётся встроенная.
 */
const BuiltinIfs = HyperFormula.getFunctionPlugin('SUMIFS') as unknown as new (...a: any[]) => any
class IndexIfsPlugin extends BuiltinIfs {
  private get ctx(): VolatileContext | undefined {
    return (this as any).config.context as VolatileContext | undefined
  }
  private flatOf(rv: SimpleRangeValue): { key: string; flat: CritValue[] } {
    const r = rv.range
    const key = r ? `${r.start.sheet}:${r.start.row}:${r.start.col}:${r.end.row}:${r.end.col}` : ''
    const hit = key ? this.ctx!.flat.get(key) : undefined
    if (hit) return { key, flat: hit }
    const flat = rv.valuesFromTopLeftCorner().map(hfScalar)
    if (key && flat.length >= 256) this.ctx!.flat.set(key, flat)
    return { key, flat }
  }
  private pick(pairs: [SimpleRangeValue, unknown][], h: number, w: number): number[] | CellError {
    const cols: CritColumn[] = []
    for (const [rv, v] of pairs) {
      if (rv.height() !== h || rv.width() !== w) return new CellError(ErrorType.VALUE)
      const { key, flat } = this.flatOf(rv)
      const s = hfScalar(v)
      if (s !== null && typeof s === 'object') return s as CellError
      cols.push({
        flat,
        crit: parseCrit(s),
        index: () => {
          let ix = key ? this.ctx!.index.get(key) : undefined
          if (!ix) {
            ix = buildIndex(flat)
            if (key) this.ctx!.index.set(key, ix)
          }
          return ix
        },
      })
    }
    return select(cols, h * w)
  }
  sumifs(ast: any, state: any) {
    if (!this.ctx?.ownIfs) return super.sumifs(ast, state)
    return (this as any).runFunction(ast.args, state, (this as any).metadata('SUMIFS'), (values: SimpleRangeValue, ...args: unknown[]) => {
      const pairs: [SimpleRangeValue, unknown][] = []
      for (let i = 0; i < args.length; i += 2) pairs.push([args[i] as SimpleRangeValue, args[i + 1]])
      const sel = this.pick(pairs, values.height(), values.width())
      if (sel instanceof CellError) return sel
      const { flat } = this.flatOf(values)
      let acc = 0
      for (const p of sel) {
        const v = flat[p]
        if (typeof v === 'number') acc += v
        else if (v !== null && typeof v === 'object') return v
      }
      return acc
    })
  }
  countifs(ast: any, state: any) {
    if (!this.ctx?.ownIfs) return super.countifs(ast, state)
    return (this as any).runFunction(ast.args, state, (this as any).metadata('COUNTIFS'), (...args: unknown[]) => {
      const pairs: [SimpleRangeValue, unknown][] = []
      for (let i = 0; i < args.length; i += 2) pairs.push([args[i] as SimpleRangeValue, args[i + 1]])
      const sel = this.pick(pairs, pairs[0][0].height(), pairs[0][0].width())
      return sel instanceof CellError ? sel : sel.length
    })
  }
}
;(IndexIfsPlugin as unknown as { implementedFunctions: unknown }).implementedFunctions = {
  SUMIFS: (BuiltinIfs as any).implementedFunctions.SUMIFS,
  COUNTIFS: (BuiltinIfs as any).implementedFunctions.COUNTIFS,
}
HyperFormula.registerFunction('SUMIFS', IndexIfsPlugin as never)
HyperFormula.registerFunction('COUNTIFS', IndexIfsPlugin as never)

/** Внутреннее значение HF в значение критерия: пусто → null, RichNumber → число, ошибка как есть. */
function hfScalar(v: unknown): CritValue {
  if (v === EmptyValue || v === undefined || v === null) return null
  if (typeof v === 'object' && !(v instanceof CellError) && 'val' in (v as object)) return (v as { val: number }).val
  return v as CritValue
}

const HF_ERROR: Record<string, ErrorCode> = {
  DIV_BY_ZERO: 'DIV0',
  NA: 'NA',
  NAME: 'NAME',
  NUM: 'NUM',
  REF: 'REF',
  VALUE: 'VALUE',
  CYCLE: 'CYCLE',
  SPILL: 'SPILL',
  ERROR: 'ERROR',
  LIC: 'ERROR',
}

const ERROR_FORMULA: Record<ErrorCode, string> = {
  DIV0: '=#DIV/0!',
  NA: '=#N/A',
  NAME: '=#NAME?',
  NUM: '=#NUM!',
  REF: '=#REF!',
  VALUE: '=#VALUE!',
  CYCLE: '=#ERROR!',
  SPILL: '=#ERROR!',
  ERROR: '=#ERROR!',
}

/** Значение для HF: строка всегда с `'` (иначе HF угадывает тип), дата серийным номером. */
function rawValue(cell: CellInput): RawCellContent {
  const v = inputScalar(cell.value)
  return typeof v === 'string' ? `'${v}` : v
}

function fromHf(v: unknown): CellValue {
  if (v instanceof DetailedCellError) return { error: HF_ERROR[v.type] ?? 'ERROR' }
  if (v === null || v === undefined) return null
  return v as CellValue
}

interface Built {
  hf: HyperFormula
  prep: PreparedBook
  context: VolatileContext
  ms: { parse: number; build: number }
}

function build(book: BookInput, clock: Clock, ownIfs: boolean): Built {
  const t0 = performance.now()
  const prep = prepareBook(book)
  const grid: Record<string, RawCellContent[][]> = {}
  for (const s of book.sheets) {
    const rows: RawCellContent[][] = []
    for (const [addr, cell] of Object.entries(s.cells)) {
      if (cell.formula !== undefined) continue
      const { row, col } = parseAddress(addr)
      ;(rows[row] ??= [])[col] = rawValue(cell)
    }
    grid[s.name] = rows
  }
  for (const f of prep.formulas) (grid[f.sheet][f.row] ??= [])[f.col] = toHfText(f.ast, prep.sheetOf)
  for (const f of prep.invalid) {
    const { row, col } = parseAddress(f.addr)
    ;(grid[f.sheet][row] ??= [])[col] = ERROR_FORMULA[invalidError(f.code)]
  }
  // дыры разреженных массивов HF понимает как пусто, но надёжнее явный null
  for (const rows of Object.values(grid)) {
    for (let r = 0; r < rows.length; r++) {
      const row = (rows[r] ??= [])
      for (let c = 0; c < row.length; c++) if (row[c] === undefined) row[c] = null
    }
  }
  const named = prep.names.map((n) => ({
    name: n.name,
    expression: toHfText(n.ast, prep.sheetOf),
    scope: n.scope === null ? undefined : prep.sheets.indexOf(n.scope),
  }))
  const context: VolatileContext = { now: nowSerial(clock), seed: clock.seed, sheetNames: Object.keys(grid), ownIfs, flat: new Map(), index: new Map() }
  const t1 = performance.now()
  const hf = HyperFormula.buildFromSheets(grid, { ...HF_CONFIG, context }, named)
  const t2 = performance.now()
  return { hf, prep, context, ms: { parse: t1 - t0, build: t2 - t1 } }
}

function readValues(hf: HyperFormula): BookValues {
  const out: BookValues = new Map()
  for (const name of hf.getSheetNames()) {
    const id = hf.getSheetId(name)!
    const cells = new Map<string, CellValue>()
    const grid = hf.getSheetValues(id)
    for (let r = 0; r < grid.length; r++) {
      const row = grid[r]
      for (let c = 0; c < row.length; c++) {
        const v = fromHf(row[c])
        if (v !== null) cells.set(formatAddress(r, c), v)
      }
    }
    out.set(name, cells)
  }
  return out
}

function volatileCells(prep: PreparedBook): string[] {
  const vol = new Set<string>(VOLATILE)
  const out: string[] = []
  for (const f of prep.formulas) {
    let hit = false
    walk(f.ast, (n) => {
      if (n.k === 'fn' && vol.has(n.name)) hit = true
    })
    if (hit) out.push(`${f.sheet}!${f.addr}`)
  }
  return out
}

/** Вариант H; `ownIfs` это третий вариант: H со своей SUMIFS/COUNTIFS (индекс общий с C). */
export function makeHfEngine(ownIfs: boolean): ProbeEngine {
  const configVersion = ownIfs ? `${HF_CONFIG_VERSION};ifs=idx` : HF_CONFIG_VERSION
  return {
    name: ownIfs ? 'H-idx' : 'H',
    configVersion,
    compute(book, { clock }): ProbeResult {
      const t0 = performance.now()
      const b = build(book, clock, ownIfs)
      const t1 = performance.now()
      const values = readValues(b.hf)
      const t2 = performance.now()
      b.hf.destroy()
      return {
        values,
        invalid: b.prep.invalid,
        hash: valuesHash(values, configVersion),
        inputHash: inputHash(book),
        volatileCells: volatileCells(b.prep),
        ms: { parse: b.ms.parse, build: b.ms.build, evaluate: t2 - t1, total: t2 - t0 },
      }
    },
    open(book, { clock }): ProbeSession {
      const { hf, prep, context } = build(book, clock, ownIfs)
      return {
        setCells(edits: CellEdit[]) {
          // кэш своей SUMIFS грубо сбрасывается целиком: индексы строятся заново лениво
          context.flat.clear()
          context.index.clear()
          const changes = hf.batch(() => {
            for (const e of edits) {
              const sheet = hf.getSheetId(prep.sheetOf(e.sheet))!
              const { row, col } = parseAddress(e.addr)
              hf.setCellContents({ sheet, row, col }, e.input.formula !== undefined ? formulaContent(e.input.formula, prep) : rawValue(e.input))
            }
          })
          return changes.length
        },
        values: () => readValues(hf),
        hash: () => valuesHash(readValues(hf), configVersion),
        close: () => hf.destroy(),
      }
    },
  }
}

export const hfEngine = makeHfEngine(false)

function formulaContent(text: string, prep: PreparedBook): RawCellContent {
  const r = validateFormula(text, { sheets: prep.sheets, names: prep.names.map((n) => n.name), functions: WHITELIST })
  return r.ok ? toHfText(r.ast, prep.sheetOf) : ERROR_FORMULA[invalidError(r.code)]
}
