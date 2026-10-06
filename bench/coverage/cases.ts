/**
 * The measurement's cases from the signature table: typical calls, hand-written edge cases and automatic edge
 * cases (values put into an argument by its kind). The choice is deterministic (a fixed seed), up to 15 cases
 * per function except for functions with many typical calls.
 */
import * as fj from '@formulajs/formulajs'
import { mulberry32 } from '../../src/synth.js'
import { DATA_SHEET, REGIONS } from './data.js'
import EXCEL_FUNCTIONS from './excel-functions.json' with { type: 'json' }
import { ARRAY_SIZE, SIGNATURES, SIZE_LIMIT, XLFN, type Arg, type Pick, type Sig } from './signatures.js'

export const SEED = 20261005
export const MAX_CASES = 15
/** The ceiling of "large" for size and count arguments (`Sig.cap`). */
export const CAP = 1000

export interface Case {
  id: string
  fn: string
  category: string
  /** typical: a typical call; edge: an edge case. */
  kind: 'typical' | 'edge'
  tag: string
  args: Arg[]
  pick?: Pick
  formula: string
}

/** Volatile: left out of the measurement. */
export const VOLATILE = ['NOW', 'RAND', 'RANDBETWEEN', 'TODAY']

type Pool = [string, Arg | Arg[]][]

const NUM: Pool = [
  ['numeric text', '3'],
  ['boolean', true],
  ['blank cell', '@empty'],
  ['error input', '@div0'],
  ['text', 'abc'],
  ['negative', -2.5],
  ['zero', 0],
  ['fractional', 0.75],
  ['large', 1e10],
  ['numeric text in a cell', '@textnum'],
  ['date in a cell', '@date'],
  ['number text with a point', '1.5'],
]
const INT: Pool = NUM.map(([t, v]) => (t === 'fractional' ? [t, 2.7] : [t, v]))
const PROB: Pool = [
  ['numeric text', '0.5'],
  ['boolean', true],
  ['blank cell', '@empty'],
  ['error input', '@div0'],
  ['text', 'abc'],
  ['negative', -0.1],
  ['zero', 0],
  ['above 1', 1.5],
  ['number text with a point', '0.25'],
]
const DATE: Pool = [
  ['ISO date text', '2025-01-31'],
  ['m/d/yyyy date text', '1/2/2025'],
  ['leap day 2024-02-29', 45351],
  ['false leap day 1900-02-29', 60],
  ['1900-02-28', 59],
  ['1900-03-01', 61],
  ['zero', 0],
  ['negative', -1],
  ['date with a time', '@datetime'],
  ['blank cell', '@empty'],
  ['error input', '@div0'],
  ['text', 'abc'],
  ['boolean', true],
]
const TIME: Pool = [
  ['time text', '18:30'],
  ['zero', 0],
  ['date with a time', '@datetime'],
  ['negative', -0.5],
  ['over a day', 1.25],
  ['blank cell', '@empty'],
  ['error input', '@div0'],
  ['text', 'abc'],
  ['boolean', true],
]
const TEXT: Pool = [
  ['Cyrillic', 'Привет Мир'],
  ['case', 'пРИВЕТ мИР'],
  ['number', 12.5],
  ['boolean', true],
  ['blank cell', '@empty'],
  ['empty string', ''],
  ['error input', '@na'],
  ['spaces', '  a  b  '],
  ['Cyrillic in a cell', '@cyr'],
]
const BOOL: Pool = [
  ['number 0', 0],
  ['number 2', 2],
  ['text TRUE', 'TRUE'],
  ['text', 'abc'],
  ['blank cell', '@empty'],
  ['error input', '@div0'],
  ['boolean in a cell', '@true'],
  ['FALSE', false],
]
const ANY: Pool = [
  ['number', 2.5],
  ['text', 'abc'],
  ['blank cell', '@empty'],
  ['boolean', true],
  ['#N/A', '@na'],
  ['#DIV/0!', '@div0'],
  ['empty string in a cell', '@emptystr'],
  ['numeric text in a cell', '@textnum'],
  ['date in a cell', '@date'],
]
/** A tail of variable length (SUM and its kin): the whole tail is replaced by a list. */
const TAIL: Pool = [
  ['range: mixed', ['@mixed']],
  ['range with an error', ['@witherr']],
  ['blank range', ['@blank']],
  ['numeric text argument', ['2', 1]],
  ['boolean argument', [true, 1]],
  ['text argument', ['abc', 1]],
  ['reference to numeric text', ['@textnum', 1]],
  ['reference to a boolean', ['@true', 1]],
  ['blank cell argument', ['@empty', 1]],
  ['negative and fractional', [-2.5, 0.75, 0]],
]
/** One range in place of an array (LARGE, PERCENTILE, CORREL): only instead of a column of 8 numbers. */
const RANGE8: Pool = [
  ['range: mixed', '@mixed'],
  ['range with an error', '@witherr'],
  ['blank range', '@blank'],
  ['different size', '@nums1short'],
]
/** An array size (ARRAY_SIZE): small number literals only. */
const SIZE: Pool = [
  ['negative', -2],
  ['zero', 0],
  ['fractional', 2.7],
  [`large (${SIZE_LIMIT})`, SIZE_LIMIT],
]
const DATES: Pool = [
  ['holidays: mixed', '@mixed'],
  ['holidays: blank range', '@blank'],
  ['holidays with an error', '@witherr'],
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

/** The argument kind at position i, the repeating tail counted. */
export function kindAt(sig: Sig, kinds: Kind[], i: number): string {
  if (i < kinds.length) return kinds[i].k
  const rep = sig.rep ?? 0
  if (!rep) throw new Error(`${sig.name}: too many args`)
  const start = kinds.length - rep
  return kinds[start + ((i - start) % rep)].k
}

/** Excel's argument counts: the required kinds before the optional ones; a tail up to 255. */
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

/** Automatic edge variants by argument position of the base call. */
function autoEdges(sig: Sig, base: Arg[], rnd: () => number): [string, Arg[]][][] {
  const kinds = parseKinds(sig)
  const out: [string, Arg[]][][] = []
  const tailStart = sig.rep ? kinds.length - sig.rep : Infinity
  for (let i = 0; i < base.length; i++) {
    const k = kindAt(sig, kinds, i)
    const list: [string, Arg[]][] = []
    if ((k === 'N' || k === 'V') && sig.rep === 1 && i === tailStart) {
      // the whole tail
      for (const [tag, v] of TAIL) list.push([tag, [...base.slice(0, i), ...(v as Arg[])]])
      out.push(list)
      break
    }
    let pool: Pool = ARRAY_SIZE[sig.name]?.includes(i) ? SIZE : (POOLS[k] ?? [])
    if (k === 'N' || k === 'M') pool = base[i] === '@nums1' || base[i] === '@nums2' ? RANGE8 : []
    if (k === 'D') pool = base[i] === '@hol' ? DATES : []
    for (const [tag, v] of pool) {
      if (v === base[i]) continue
      if (sig.cap && (tag === 'date in a cell' || tag === 'number text with a point')) continue
      const big = sig.cap && tag === 'large'
      const args = [...base]
      args[i] = big ? CAP : (v as Arg)
      list.push([big ? `large (${CAP})` : tag, args])
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

/** The choice: two variants of the first argument per round, one of each other (the order shuffled by the seed). */
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
  /** Excel names formulajs has as functions (the measured set before exclusions). */
  inFormulajs: string[]
  excluded: { fn: string; why: string }[]
  /** formulajs names Excel lacks: synonyms and dotted names not as in Excel. */
  notInExcel: { name: string; same: string[] }[]
  /** Excel functions formulajs lacks (or has as something other than a function). */
  missing: { fn: string; category: string; note?: string }[]
}

const getFj = (path: string): unknown => path.split('.').reduce<any>((o, k) => (o == null ? undefined : o[k]), fj)

export function buildCases(): CaseSet {
  const excel = new Map(EXCEL_FUNCTIONS.map((x) => [x.name, x]))
  // every callable path of formulajs
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
    return { fn: x.name, category: x.category, note: v !== undefined ? 'in formulajs an object of subfunctions, not a function' : undefined }
  })
  const excluded = VOLATILE.map((fn) => ({ fn, why: 'volatile' }))

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
      ...sig.calls.map((a) => ['typical', 'typical', a] as ['typical', string, Arg[]]),
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
  const bad = sizeGuardViolations(cases)
  if (bad.length) throw new Error(`size guard: ${bad.join('; ')}`)
  return { cases, inFormulajs, excluded, notInExcel, missing }
}

/** Cases with something other than a small number literal in an array-size position (see ARRAY_SIZE). */
export function sizeGuardViolations(cases: Case[]): string[] {
  const out: string[] = []
  for (const c of cases)
    for (const i of ARRAY_SIZE[c.fn] ?? []) {
      if (i >= c.args.length) continue
      const a = c.args[i]
      if (typeof a !== 'number' || !Number.isFinite(a) || Math.abs(a) > SIZE_LIMIT) out.push(`${c.id} ${c.formula}`)
    }
  return out
}

function pickLabel(p: Pick): string {
  return p === 'rows' ? 'row count' : p === 'cols' ? 'column count' : `element ${p[0]},${p[1]}`
}
