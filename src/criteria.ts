/**
 * Критерии SUMIF, SUMIFS, COUNTIFS и индекс по ним: строки без
 * учёта регистра, но с учётом «й»/«и» и «ё»/«е»; шаблоны `*`, `?`, экранирование `~`; числа и ISO-даты
 * в тексте критерия сравниваются как числа; `""` это пустые, `"<>"` непустые.
 */
import { dateToSerial } from './model.js'

/** Значение ячейки для критерия; ошибки любым объектом. */
export type CritValue = number | string | boolean | null | object

export type Crit =
  | { op: string; kind: 'num'; n: number }
  | { op: string; kind: 'bool'; b: boolean }
  | { op: string; kind: 'str'; s: string }
  | { op: string; kind: 'wild'; re: RegExp }
  | { op: string; kind: 'blank' }

/**
 * Порядок строк для `<`/`>`: без учёта регистра, с учётом «й»/«и». В QuickJS нет Intl, там порядок
 * по кодам символов строчных букв (замер в QuickJS справочный).
 */
export const collator: { compare(a: string, b: string): number } =
  typeof Intl !== 'undefined' ? new Intl.Collator('ru-RU', { sensitivity: 'accent' }) : { compare: (a, b) => ((a = a.toLowerCase()) < (b = b.toLowerCase()) ? -1 : a > b ? 1 : 0) }
const NUM_RE = /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/
const cache = new Map<string, Crit>()

function wildRe(s: string): RegExp | null {
  if (!/[*?~]/.test(s)) return null
  let re = ''
  for (let i = 0; i < s.length; i++) {
    const ch = s[i]
    if (ch === '~' && i + 1 < s.length && '*?~'.includes(s[i + 1])) re += '\\' + s[++i]
    else if (ch === '*') re += '[\\s\\S]*'
    else if (ch === '?') re += '[\\s\\S]'
    else re += ch.replace(/[.+^${}()|[\]\\]/g, '\\$&')
  }
  return new RegExp(`^${re}$`, 'iu')
}

/** Критерий из значения (скаляр после вычисления аргумента). */
export function parseCrit(v: number | string | boolean | null): Crit {
  const key = typeof v + ':' + String(v)
  const hit = cache.get(key)
  if (hit) return hit
  let c: Crit
  if (typeof v === 'number') c = { op: '=', kind: 'num', n: v }
  else if (typeof v === 'boolean') c = { op: '=', kind: 'bool', b: v }
  else if (v === null) c = { op: '=', kind: 'num', n: 0 }
  else {
    const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(v)!
    const op = m[1] ?? '='
    const rest = m[2]
    const t = rest.trim()
    if (rest === '') c = { op, kind: 'blank' }
    else if (NUM_RE.test(t)) c = { op, kind: 'num', n: Number(t) }
    else if (DATE_RE.test(t)) c = { op, kind: 'num', n: dateToSerial(t) }
    else if (/^(TRUE|FALSE)$/i.test(rest)) c = { op, kind: 'bool', b: rest.toUpperCase() === 'TRUE' }
    else {
      const re = op === '=' || op === '<>' ? wildRe(rest) : null
      c = re ? { op, kind: 'wild', re } : { op, kind: 'str', s: rest.toLowerCase() }
    }
  }
  cache.set(key, c)
  return c
}

const ord = (x: number, y: number) => (x < y ? -1 : x > y ? 1 : 0)
const opOk = (op: string, c: number) => (op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0)

export function matches(v: CritValue, c: Crit): boolean {
  switch (c.kind) {
    case 'num':
      if (typeof v !== 'number') return c.op === '<>'
      return opOk(c.op, ord(v, c.n))
    case 'bool':
      if (typeof v !== 'boolean') return c.op === '<>'
      return opOk(c.op, ord(Number(v), Number(c.b)))
    case 'blank':
      if (c.op === '=') return v === null || v === ''
      if (c.op === '<>') return !(v === null || v === '')
      return false
    case 'wild':
      if (typeof v !== 'string') return c.op === '<>'
      return c.re.test(v) === (c.op === '=')
    case 'str':
      if (typeof v !== 'string') return c.op === '<>'
      if (c.op === '=' || c.op === '<>') return (v.toLowerCase() === c.s) === (c.op === '=')
      return opOk(c.op, collator.compare(v.toLowerCase(), c.s))
  }
}

/** Ключ равенства в индексе: строки без учёта регистра, числа, логические. */
function eqKey(v: CritValue): string | null {
  if (typeof v === 'string') return 's' + v.toLowerCase()
  if (typeof v === 'number') return 'n' + v
  if (typeof v === 'boolean') return 'b' + v
  return null
}

function critKey(c: Crit): string | null {
  if (c.op !== '=') return null
  if (c.kind === 'str') return 's' + c.s
  if (c.kind === 'num') return 'n' + c.n
  if (c.kind === 'bool') return 'b' + c.b
  return null
}

export type CritIndex = Map<string, number[]>

export function buildIndex(flat: readonly CritValue[]): CritIndex {
  const ix: CritIndex = new Map()
  for (let i = 0; i < flat.length; i++) {
    const k = eqKey(flat[i])
    if (k === null) continue
    const list = ix.get(k)
    if (list) list.push(i)
    else ix.set(k, [i])
  }
  return ix
}

export interface CritColumn {
  flat: readonly CritValue[]
  crit: Crit
  /** Индекс этого диапазона (строится лениво и кэшируется вызывающим). */
  index: () => CritIndex
}

/**
 * Позиции, где выполнены все критерии. Первый критерий-равенство идёт через индекс, остальные
 * проверяются по кандидатам; без равенства полный проход.
 */
export function select(cols: CritColumn[], n: number): number[] {
  let cand: number[] | null = null
  let used = -1
  for (let i = 0; i < cols.length; i++) {
    const k = critKey(cols[i].crit)
    if (k === null) continue
    cand = cols[i].index().get(k) ?? []
    used = i
    break
  }
  const out: number[] = []
  const check = (p: number) => {
    for (let i = 0; i < cols.length; i++) if (i !== used && !matches(cols[i].flat[p], cols[i].crit)) return
    out.push(p)
  }
  if (cand) for (const p of cand) check(p)
  else for (let p = 0; p < n; p++) check(p)
  return out
}
