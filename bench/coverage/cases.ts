/**
 * Случаи замера из таблицы сигнатур: типичные вызовы, ручные краевые и автоматические краевые
 * (подстановка в аргумент значений по его виду). Отбор детерминирован (сид фиксирован), на функцию
 * до 15 случаев, кроме функций с большим числом типичных вызовов.
 */
import * as fj from '@formulajs/formulajs'
import { mulberry32 } from '../../src/synth.js'
import { DATA_SHEET, REGIONS } from './data.js'
import EXCEL_FUNCTIONS from './excel-functions.json' with { type: 'json' }
import { SIGNATURES, XLFN, type Arg, type Pick, type Sig } from './signatures.js'

export const SEED = 20261005
export const MAX_CASES = 15
/** Потолок «большого» для аргументов-размеров и счётчиков (`Sig.cap`). */
export const CAP = 1000

export interface Case {
  id: string
  fn: string
  category: string
  /** typical: типичный вызов; edge: краевой. */
  kind: 'typical' | 'edge'
  tag: string
  args: Arg[]
  pick?: Pick
  formula: string
}

/** Волатильные: исключены из замера. */
export const VOLATILE = ['NOW', 'RAND', 'RANDBETWEEN', 'TODAY']

type Pool = [string, Arg | Arg[]][]

const NUM: Pool = [
  ['текст-число', '3'],
  ['логическое', true],
  ['пустая ячейка', '@empty'],
  ['ошибка на входе', '@div0'],
  ['текст', 'abc'],
  ['отрицательное', -2.5],
  ['ноль', 0],
  ['дробное', 0.75],
  ['большое', 1e10],
  ['текст-число в ячейке', '@textnum'],
  ['дата в ячейке', '@date'],
  ['число текстом с точкой', '1.5'],
]
const INT: Pool = NUM.map(([t, v]) => (t === 'дробное' ? [t, 2.7] : [t, v]))
const PROB: Pool = [
  ['текст-число', '0.5'],
  ['логическое', true],
  ['пустая ячейка', '@empty'],
  ['ошибка на входе', '@div0'],
  ['текст', 'abc'],
  ['отрицательное', -0.1],
  ['ноль', 0],
  ['больше 1', 1.5],
  ['число текстом с точкой', '0.25'],
]
const DATE: Pool = [
  ['дата текстом ISO', '2025-01-31'],
  ['дата текстом m/d/yyyy', '1/2/2025'],
  ['високосный 29.02.2024', 45351],
  ['ложный 29.02.1900', 60],
  ['28.02.1900', 59],
  ['01.03.1900', 61],
  ['ноль', 0],
  ['отрицательное', -1],
  ['дата со временем', '@datetime'],
  ['пустая ячейка', '@empty'],
  ['ошибка на входе', '@div0'],
  ['текст', 'abc'],
  ['логическое', true],
]
const TIME: Pool = [
  ['время текстом', '18:30'],
  ['ноль', 0],
  ['дата со временем', '@datetime'],
  ['отрицательное', -0.5],
  ['больше суток', 1.25],
  ['пустая ячейка', '@empty'],
  ['ошибка на входе', '@div0'],
  ['текст', 'abc'],
  ['логическое', true],
]
const TEXT: Pool = [
  ['кириллица', 'Привет Мир'],
  ['регистр', 'пРИВЕТ мИР'],
  ['число', 12.5],
  ['логическое', true],
  ['пустая ячейка', '@empty'],
  ['пустая строка', ''],
  ['ошибка на входе', '@na'],
  ['пробелы', '  a  b  '],
  ['кириллица в ячейке', '@cyr'],
]
const BOOL: Pool = [
  ['число 0', 0],
  ['число 2', 2],
  ['текст TRUE', 'TRUE'],
  ['текст', 'abc'],
  ['пустая ячейка', '@empty'],
  ['ошибка на входе', '@div0'],
  ['логическое в ячейке', '@true'],
  ['ЛОЖЬ', false],
]
const ANY: Pool = [
  ['число', 2.5],
  ['текст', 'abc'],
  ['пустая ячейка', '@empty'],
  ['логическое', true],
  ['#N/A', '@na'],
  ['#DIV/0!', '@div0'],
  ['пустая строка в ячейке', '@emptystr'],
  ['текст-число в ячейке', '@textnum'],
  ['дата в ячейке', '@date'],
]
/** Хвост переменной длины (SUM и родня): весь хвост заменяется списком. */
const TAIL: Pool = [
  ['диапазон: смесь', ['@mixed']],
  ['диапазон с ошибкой', ['@witherr']],
  ['пустой диапазон', ['@blank']],
  ['текст-число аргументом', ['2', 1]],
  ['логическое аргументом', [true, 1]],
  ['текст аргументом', ['abc', 1]],
  ['ссылка на текст-число', ['@textnum', 1]],
  ['ссылка на логическое', ['@true', 1]],
  ['пустая ячейка аргументом', ['@empty', 1]],
  ['отрицательные и дробные', [-2.5, 0.75, 0]],
]
/** Один диапазон на месте массива (LARGE, PERCENTILE, CORREL): только вместо колонки из 8 чисел. */
const RANGE8: Pool = [
  ['диапазон: смесь', '@mixed'],
  ['диапазон с ошибкой', '@witherr'],
  ['пустой диапазон', '@blank'],
  ['разный размер', '@nums1short'],
]
const DATES: Pool = [
  ['праздники: смесь', '@mixed'],
  ['праздники: пустой диапазон', '@blank'],
  ['праздники с ошибкой', '@witherr'],
]

const POOLS: Record<string, Pool> = { n: NUM, i: INT, p: PROB, d: DATE, t: TIME, s: TEXT, b: BOOL, a: ANY }

export interface Kind {
  k: string
  optional: boolean
}

export function parseKinds(sig: Sig): Kind[] {
  return sig.kinds
    .split(' ')
    .filter(Boolean)
    .map((t) => ({ k: t[0], optional: t.endsWith('?') }))
}

/** Вид аргумента на позиции i с учётом повторяющегося хвоста. */
export function kindAt(sig: Sig, kinds: Kind[], i: number): string {
  if (i < kinds.length) return kinds[i].k
  const rep = sig.rep ?? 0
  if (!rep) throw new Error(`${sig.name}: too many args`)
  const start = kinds.length - rep
  return kinds[start + ((i - start) % rep)].k
}

/** Число аргументов по Excel: обязательные виды до необязательных; хвост до 255. */
export function arity(sig: Sig): { min: number; max: number } {
  const kinds = parseKinds(sig)
  return { min: kinds.filter((x) => !x.optional).length, max: sig.rep ? 255 : kinds.length }
}

const norm = (a: Arg): Arg => (typeof a === 'number' ? Number(a.toPrecision(15)) : a)

export function renderArg(a: Arg): string {
  if (typeof a === 'boolean') return a ? 'TRUE' : 'FALSE'
  if (typeof a === 'number') return String(a).replace('e', 'E')
  if (a.startsWith('@')) {
    const r = REGIONS[a.slice(1)]
    if (!r) throw new Error(`unknown region ${a}`)
    return `${DATA_SHEET}!${r.ref}`
  }
  return `"${a.replace(/"/g, '""')}"`
}

export function renderFormula(fn: string, args: Arg[], pick?: Pick): string {
  const call = `${XLFN[fn] ?? fn}(${args.map(renderArg).join(',')})`
  if (pick === undefined) return call
  if (pick === 'rows') return `ROWS(${call})`
  if (pick === 'cols') return `COLUMNS(${call})`
  return `INDEX(${call},${pick[0]},${pick[1]})`
}

/** Автоматические краевые варианты по позициям аргументов базового вызова. */
function autoEdges(sig: Sig, base: Arg[], rnd: () => number): [string, Arg[]][][] {
  const kinds = parseKinds(sig)
  const out: [string, Arg[]][][] = []
  const tailStart = sig.rep ? kinds.length - sig.rep : Infinity
  for (let i = 0; i < base.length; i++) {
    const k = kindAt(sig, kinds, i)
    const list: [string, Arg[]][] = []
    if ((k === 'N' || k === 'V') && sig.rep === 1 && i === tailStart) {
      // хвост целиком
      for (const [tag, v] of TAIL) list.push([tag, [...base.slice(0, i), ...(v as Arg[])]])
      out.push(list)
      break
    }
    let pool: Pool = POOLS[k] ?? []
    if (k === 'N' || k === 'M') pool = base[i] === '@nums1' || base[i] === '@nums2' ? RANGE8 : []
    if (k === 'D') pool = base[i] === '@hol' ? DATES : []
    for (const [tag, v] of pool) {
      if (v === base[i] || sig.skip?.includes(tag)) continue
      if (sig.cap && tag === 'дата в ячейке') continue
      const big = sig.cap && tag === 'большое'
      const args = [...base]
      args[i] = big ? CAP : (v as Arg)
      list.push([big ? `большое (${CAP})` : tag, args])
    }
    if (i > 0) shuffle(list, rnd)
    out.push(list)
  }
  return out
}

function shuffle<T>(xs: T[], rnd: () => number) {
  for (let i = xs.length - 1; i > 0; i--) {
    const j = Math.floor(rnd() * (i + 1))
    ;[xs[i], xs[j]] = [xs[j], xs[i]]
  }
}

/** Отбор: первый аргумент по два варианта за круг, остальные по одному (порядок перемешан сидом). */
function pickEdges(groups: [string, Arg[]][][], budget: number): [string, Arg[]][] {
  const out: [string, Arg[]][] = []
  const idx = groups.map(() => 0)
  while (out.length < budget && groups.some((g, i) => idx[i] < g.length)) {
    groups.forEach((g, i) => {
      const take = i === 0 ? 2 : 1
      for (let t = 0; t < take && idx[i] < g.length && out.length < budget; t++) out.push(g[idx[i]++])
    })
  }
  return out
}

export interface CaseSet {
  cases: Case[]
  /** Имена Excel, которые formulajs отдаёт функцией (замеряемый набор до исключений). */
  inFormulajs: string[]
  excluded: { fn: string; why: string }[]
  /** Имена formulajs, которых нет в Excel: синонимы и имена с точкой не по Excel. */
  notInExcel: { name: string; same: string[] }[]
  /** Функции Excel, которых нет в formulajs (или в formulajs это не функция). */
  missing: { fn: string; category: string; note?: string }[]
}

const getFj = (path: string): unknown => path.split('.').reduce<any>((o, k) => (o == null ? undefined : o[k]), fj)

export function buildCases(): CaseSet {
  const excel = new Map(EXCEL_FUNCTIONS.map((x) => [x.name, x]))
  // все вызываемые пути formulajs
  const paths = new Map<string, Function>()
  const walk = (o: any, prefix: string, depth: number) => {
    for (const k of Object.keys(o)) {
      if (!prefix && k === 'utils') continue
      const v = o[k]
      const p = prefix ? `${prefix}.${k}` : k
      if (typeof v === 'function') paths.set(p, v)
      if (v && (typeof v === 'function' || typeof v === 'object') && depth < 3) walk(v, p, depth + 1)
    }
  }
  walk(fj, '', 0)
  const inFormulajs = [...paths.keys()].filter((p) => excel.has(p)).sort()
  const byFn = new Map<Function, string[]>()
  for (const [p, f] of paths) byFn.set(f, [...(byFn.get(f) ?? []), p])
  const notInExcel = [...paths.keys()]
    .filter((p) => !excel.has(p))
    .sort()
    .map((name) => ({ name, same: byFn.get(paths.get(name)!)!.filter((q) => q !== name && excel.has(q)) }))
  const missing = EXCEL_FUNCTIONS.filter((x) => !paths.has(x.name)).map((x) => {
    const v = getFj(x.name)
    return { fn: x.name, category: x.category, note: v !== undefined ? 'в formulajs это не функция, а объект с подфункциями' : undefined }
  })
  const excluded = VOLATILE.map((fn) => ({ fn, why: 'волатильная' }))

  const sigs = new Map(SIGNATURES.map((s) => [s.name, s]))
  const want = inFormulajs.filter((n) => !VOLATILE.includes(n))
  const noSig = want.filter((n) => !sigs.has(n))
  const extraSig = [...sigs.keys()].filter((n) => !want.includes(n))
  if (noSig.length || extraSig.length) throw new Error(`signatures out of sync: no signature ${noSig.join(',')}; not in set ${extraSig.join(',')}`)
  if (SIGNATURES.length !== sigs.size) throw new Error('duplicate signature')

  const rnd = mulberry32(SEED)
  const cases: Case[] = []
  for (const name of want) {
    const sig = sigs.get(name)!
    const { min, max } = arity(sig)
    const extra = sig.extra ?? []
    const seen = new Set([...sig.calls, ...extra.map((x) => x[1])].map((a) => JSON.stringify(a.map(norm))))
    const budget = Math.max(0, MAX_CASES - sig.calls.length - extra.length)
    const edges = sig.noAuto || !sig.calls.length ? [] : autoEdges(sig, sig.calls[0], rnd).map((g) => g.filter(([, a]) => !seen.has(JSON.stringify(a.map(norm)))))
    const auto = pickEdges(edges, budget)
    const all: ['typical' | 'edge', string, Arg[], Pick?][] = [
      ...sig.calls.map((a) => ['typical', 'обычный', a] as ['typical', string, Arg[]]),
      ...[...extra, ...auto].map(([t, a, p]) => ['edge', t, a, p] as ['edge', string, Arg[], Pick?]),
    ]
    let k = 0
    for (const [kind, tag, raw, own] of all) {
      const args = raw.map(norm)
      if (args.length < min || args.length > max) throw new Error(`${name}: ${args.length} args, Excel allows ${min}..${max}`)
      const picks: (Pick | undefined)[] = sig.picks ? (kind === 'typical' ? sig.picks : [own ?? sig.picks[0]]) : [undefined]
      for (const pick of picks) {
        const id = `${name}#${++k}`
        cases.push({ id, fn: name, category: excel.get(name)!.category, kind, tag: pick && kind === 'typical' && pick !== sig.picks![0] ? `${tag}, ${pickLabel(pick)}` : tag, args, pick, formula: renderFormula(name, args, pick) })
      }
    }
  }
  return { cases, inFormulajs, excluded, notInExcel, missing }
}

function pickLabel(p: Pick): string {
  return p === 'rows' ? 'число строк' : p === 'cols' ? 'число колонок' : `элемент ${p[0]},${p[1]}`
}
