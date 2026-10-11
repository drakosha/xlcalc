/** The error of input of the wrong shape, in a module of its own: the modules without the engine (`a1.ts`) throw it too. */

/** Input of the wrong shape. `path` points at the offending part, for example `book.sheets[1].cells["B0"]`. */
export class InputError extends TypeError {
  override readonly name = 'InputError'
  constructor(
    readonly path: string,
    problem: string,
  ) {
    super(`${path}: ${problem}`)
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
