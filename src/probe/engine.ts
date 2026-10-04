/** Общий интерфейс прототипов движка для замера (не контракт шага 2). */
import type { BookValues } from './canon'
import type { BookInput, CellInput, CellValue, Clock, ErrorCode } from './model'
import type { InvalidFormula } from './prep'

export interface ProbeTimings {
  /** Наш разбор и проверка формул. */
  parse: number
  /** Сборка движка (HF: buildFromSheets; C: компиляция и граф). */
  build: number
  /** Вычисление (у HF входит в build, здесь 0). */
  evaluate: number
  total: number
}

export interface ProbeResult {
  values: BookValues
  invalid: InvalidFormula[]
  hash: string
  inputHash: string
  /** Ячейки с волатильными вызовами (прямыми). */
  volatileCells: string[]
  ms: ProbeTimings
}

export interface CellEdit {
  sheet: string
  addr: string
  input: CellInput
}

export interface ProbeSession {
  /** Атомарно, один пересчёт; число изменившихся значений. */
  setCells(edits: CellEdit[]): number
  values(): BookValues
  hash(): string
  close(): void
}

export interface ProbeEngine {
  name: string
  configVersion: string
  compute(book: BookInput, opts: { clock: Clock }): ProbeResult
  open(book: BookInput, opts: { clock: Clock }): ProbeSession
}

export const errorValue = (code: ErrorCode): CellValue => ({ error: code })
