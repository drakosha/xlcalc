/** Прогон списка семантики по всем вариантам: H, C (своя SUMIFS), C с SUMIFS formulajs, голая formulajs, оракул. */
import { cEngine } from '../../src/probe/c-engine'
import { hfEngine } from '../../src/probe/hf-engine'
import type { CellValue } from '../../src/probe/model'
import { oracleValues } from './oracle'
import { CASES, semanticsBook, type SemCase } from './semantics-cases'

export interface SemRow {
  c: SemCase
  H: CellValue
  C: CellValue
  Cfj: CellValue
  /** Голая formulajs: значение, `Date ISO` для дат, `нет функции`. */
  fjRaw: CellValue | string
  oracle: CellValue | undefined
}

export function sameValue(a: unknown, b: unknown): boolean {
  if (typeof a === 'number' && typeof b === 'number') return Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a))
  return JSON.stringify(a) === JSON.stringify(b)
}

function fjValue(c: SemCase): CellValue | string {
  if (!c.fj) return 'нет функции'
  let v: unknown
  try {
    v = c.fj()
  } catch (e) {
    return `исключение ${(e as Error).message}`
  }
  if (v instanceof Error) {
    const code = { '#DIV/0!': 'DIV0', '#N/A': 'NA', '#NAME?': 'NAME', '#NUM!': 'NUM', '#VALUE!': 'VALUE' }[v.message]
    return code ? ({ error: code } as CellValue) : `ошибка ${v.message}`
  }
  if (v instanceof Date) return `Date ${v.toISOString()}`
  return v as CellValue
}

export async function runSemantics(withOracle = true): Promise<SemRow[]> {
  const book = semanticsBook(CASES)
  const clock = { now: '2026-10-04T12:00:00Z', seed: 1 }
  const h = hfEngine.compute(book, { clock }).values.get('Cases')!
  const c = cEngine('idx').compute(book, { clock }).values.get('Cases')!
  const cf = cEngine('fj').compute(book, { clock }).values.get('Cases')!
  const oracle = withOracle ? await oracleValues(book, 'Cases') : null
  return CASES.map((cs, i) => ({
    c: cs,
    H: h.get(`A${i + 1}`) ?? null,
    C: c.get(`A${i + 1}`) ?? null,
    Cfj: cf.get(`A${i + 1}`) ?? null,
    fjRaw: fjValue(cs),
    oracle: oracle ? (oracle.get(`A${i + 1}`) ?? null) : undefined,
  }))
}
