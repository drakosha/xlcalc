/**
 * Поячеечное сравнение значений двух движков (не по хэшу: в хэш входит конфигурация движка).
 * Числа с допуском |a−b| ≤ 1e-9·max(1,|a|), строки, логические и ошибки точно.
 */
import type { BookValues } from './canon'
import type { BookInput, CellValue } from './model'

export interface CellDiff {
  sheet: string
  addr: string
  kind: string
  a: CellValue | undefined
  b: CellValue | undefined
}

export interface Comparison {
  cells: number
  equal: number
  diffs: CellDiff[]
  /** Вид формулы → [всего, расходится]. */
  byKind: Record<string, [number, number]>
}

export function sameCell(a: CellValue | undefined, b: CellValue | undefined): boolean {
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a))
  if (a !== null && b !== null && typeof a === 'object' && typeof b === 'object') return a.error === b.error
  return (a ?? null) === (b ?? null)
}

/** Вид ячейки: функция верхнего уровня формулы, «ссылка», «выражение» или «значение». */
export function cellKind(formula: string | undefined): string {
  if (formula === undefined) return 'значение'
  const m = /^\s*([A-Za-z][A-Za-z0-9.]*)\s*\(/.exec(formula)
  if (m) return m[1].toUpperCase()
  if (/^\s*('[^']*(?:''[^']*)*'|[\p{L}_][\p{L}\p{N}_.]*)?!?\$?[A-Za-z]{1,3}\$?\d+\s*$/u.test(formula)) return 'ссылка'
  return 'выражение'
}

export function compareValues(book: BookInput, a: BookValues, b: BookValues): Comparison {
  const res: Comparison = { cells: 0, equal: 0, diffs: [], byKind: {} }
  const formulaOf = new Map<string, string | undefined>()
  for (const s of book.sheets) for (const [addr, c] of Object.entries(s.cells)) formulaOf.set(`${s.name}!${addr}`, c.formula)
  const keys = new Set<string>()
  for (const v of [a, b]) for (const [sheet, cells] of v) for (const addr of cells.keys()) keys.add(`${sheet}\u0000${addr}`)
  for (const key of keys) {
    const [sheet, addr] = key.split('\u0000')
    const va = a.get(sheet)?.get(addr)
    const vb = b.get(sheet)?.get(addr)
    const kind = cellKind(formulaOf.get(`${sheet}!${addr}`))
    const k = (res.byKind[kind] ??= [0, 0])
    k[0]++
    res.cells++
    if (sameCell(va, vb)) res.equal++
    else {
      k[1]++
      res.diffs.push({ sheet, addr, kind, a: va, b: vb })
    }
  }
  return res
}
