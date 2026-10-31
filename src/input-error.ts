/** The error of input of the wrong shape, in a module of its own: the modules without the engine (`a1.ts`) throw it too. */
import type { PivotRefusalReason } from './contract.js'

/** Input of the wrong shape. `path` points at the offending part, for example `book.sheets[1].cells["B0"]`. */
export class InputError extends TypeError {
  override readonly name: string = 'InputError'
  constructor(
    readonly path: string,
    problem: string,
  ) {
    super(`${path}: ${problem}`)
  }
}

/**
 * An edit that throws over a PivotTable's cells (`setCells` of several cells, `replaceCells`, `clearCells`,
 * `applyChanges`, `validateEntry`), an `InputError` that names Excel's alert (`PivotRefusalReason`) and the cell.
 */
export class PivotInputError extends InputError {
  override readonly name = 'PivotInputError'
  constructor(
    path: string,
    problem: string,
    readonly reason: PivotRefusalReason,
    readonly sheet: string,
    readonly cell: string,
  ) {
    super(path, problem)
  }
}

/** A value as a message names it: `the string "x"`, `number 3`, `an object`. */
export function what(v: unknown): string {
  if (v === undefined || v === null) return String(v)
  if (Array.isArray(v)) return 'an array'
  if (typeof v === 'object') return 'an object'
  if (typeof v === 'function') return 'a function'
  if (typeof v === 'string') return `the string ${JSON.stringify(v)}`
  return `${typeof v} ${String(v)}`
}
