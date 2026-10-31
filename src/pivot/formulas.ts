/**
 * Formulas of calculated fields and items, calculated by the engine: a formula's names (fields by name, `'Unit price'`
 * quoted; items as `Region[East]`, `Product['Long name']`, `East` or `'East'`) read into the arguments of a template,
 * and templates over their arguments calculated in batches, each a row of a sheet of their own.
 */
import type { CellInput, CellValue } from '../contract.js'
import { createState } from '../engine.js'
import { formatAddress } from '../model.js'
import { clockOptions } from '../result.js'
import { stateValues } from '../structure/values.js'
import { ERROR_TEXT } from '../syntax/errors.js'
import { eachToken } from '../syntax/links.js'

/** A formula read: its text with the k-th name it names as `__pivot<k>_`, and what each names. */
export interface Template<T> {
  text: string
  args: T[]
}

const MARK = /__pivot(\d+)_/g
const WORD = /^[\p{L}_\\][\p{L}\p{N}._]*/u

/** A name quoted by `'` from `at` (`''` an apostrophe in it): the name and the position after it. */
function quoted(s: string, at: number): [string, number] {
  let out = ''
  let i = at + 1
  for (; i < s.length; i++) {
    if (s[i] !== "'") out += s[i]
    else if (s[i + 1] === "'") (out += "'"), i++
    else break
  }
  return [out, i + 1]
}

/**
 * Reads a calculated field's or item's formula (`=` before it or not): `resolve` gives what a name stands for (`field`
 * the name before `[item]`, null for a name alone), or undefined for none, which is left for the engine (`#NAME?`).
 */
export function template<T>(formula: string, resolve: (field: string | null, name: string) => T | undefined): Template<T> {
  const s = formula.replace(/^\s*=/, '')
  const args: T[] = []
  const mark = (x: T) => `__pivot${args.push(x) - 1}_`
  let text = ''
  let i = 0
  while (i < s.length) {
    const ch = s[i]
    if (ch === '"') {
      let j = i + 1
      while (j < s.length && !(s[j] === '"' && s[j + 1] !== '"')) j += s[j] === '"' ? 2 : 1
      text += s.slice(i, j + 1)
      i = j + 1
      continue
    }
    let name: string | undefined
    let next = i
    if (ch === "'") [name, next] = quoted(s, i)
    else if (!/[\p{L}\p{N}._]/u.test(s[i - 1] ?? '')) {
      const w = WORD.exec(s.slice(i))
      if (w && s[i + w[0].length] === '[') [name, next] = [w[0], i + w[0].length]
    }
    if (name === undefined) {
      text += ch
      i++
      continue
    }
    if (s[next] === '[') {
      // an item of a field: Region[East], Product['Long name']
      const close = s[next + 1] === "'" ? quoted(s, next + 1) : [s.slice(next + 1, s.indexOf(']', next)), s.indexOf(']', next)]
      const [item, after] = close as [string, number]
      const x = resolve(name, item)
      text += x === undefined ? '#REF!' : mark(x)
      i = after + (s[after] === ']' ? 1 : 0)
      continue
    }
    const x = resolve(null, name)
    text += x === undefined ? '#NAME?' : mark(x)
    i = next
  }
  // names alone, which the tokenizer reads as names
  text = eachToken(text, (tok, t) => {
    if (tok.t !== 'ident' && tok.t !== 'ref') return undefined
    const x = resolve(null, t)
    return x === undefined ? undefined : mark(x)
  })
  return { text, args }
}

/** A value as a cell of the sheet of a batch. */
const input = (v: CellValue): CellInput | undefined => (v === null ? undefined : typeof v === 'object' ? { formula: ERROR_TEXT[v.error] } : { value: v })

const CLOCK = clockOptions({ now: '2000-01-01', seed: 0 })

/**
 * Templates over their arguments' values calculated by the engine in one go: each a row of a sheet, its arguments in
 * the columns after the first, its formula in the first; a blank argument an empty cell.
 */
export function calculate(rows: readonly { text: string; values: readonly CellValue[] }[]): CellValue[] {
  if (!rows.length) return []
  const cells: Record<string, CellInput> = {}
  rows.forEach(({ text, values }, r) => {
    values.forEach((v, k) => {
      const c = input(v)
      if (c) cells[formatAddress(r, k + 1)] = c
    })
    cells[formatAddress(r, 0)] = { formula: text.replace(MARK, (_, k: string) => formatAddress(r, Number(k) + 1)) }
  })
  const got = stateValues(createState({ sheets: [{ name: 'C', cells }] }, CLOCK))('C', { r1: 0, r2: rows.length - 1, c1: 0, c2: 0 })
  const out: CellValue[] = rows.map(() => 0)
  for (const [r, , v] of got) out[r] = v
  return out
}
