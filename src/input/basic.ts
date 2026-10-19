/**
 * The checks of input's shape that the API's boundaries share: a plain object, an array, a string, a boolean, a number
 * within bounds, one of a few values, and the fields an object may have. Each refuses with `InputError` naming the
 * place; the module needs nothing of the engine.
 */
import { InputError, what } from '../input-error.js'
import { optionHint } from '../strict-args.js'

/** A plain object: not null, not an array. */
export const isObject = (v: unknown): v is Record<string, unknown> => typeof v === 'object' && v !== null && !Array.isArray(v)

export function object(v: unknown, path: string, expected: string): Record<string, unknown> {
  if (!isObject(v)) throw new InputError(path, `expected ${expected}, got ${what(v)}`)
  return v
}

export function array(v: unknown, path: string, expected: string): unknown[] {
  if (!Array.isArray(v)) throw new InputError(path, `expected ${expected}, got ${what(v)}`)
  return v
}

export function checkString(v: unknown, path: string): string {
  if (typeof v !== 'string') throw new InputError(path, `expected a string, got ${what(v)}`)
  return v
}

export function checkBoolean(v: unknown, path: string) {
  if (typeof v !== 'boolean') throw new InputError(path, `expected a boolean, got ${what(v)}`)
}

export function checkOneOf(v: unknown, path: string, values: readonly (string | number)[]) {
  if (!values.includes(v as string)) throw new InputError(path, `expected one of ${values.map((x) => JSON.stringify(x)).join(', ')}, got ${what(v)}`)
}

export function checkNumber(v: unknown, path: string, min: number, max: number, integer = false) {
  if (typeof v !== 'number' || !Number.isFinite(v) || v < min || v > max || (integer && !Number.isInteger(v)))
    throw new InputError(path, `expected ${integer ? 'an integer' : 'a number'} from ${min} to ${max}, got ${what(v)}`)
}

/** A whole number from `min` to `max`, its refusal showing the value as JSON (`checkNumber` names it as `what` does): a structure operation's. */
export function integer(v: unknown, path: string, min: number, max: number): number {
  if (typeof v !== 'number' || !Number.isInteger(v) || v < min || v > max) throw new InputError(path, `expected an integer from ${min} to ${max}, got ${JSON.stringify(v)}`)
  return v
}

/** Fields by their old names, refused since they were renamed, and the name of each now. */
const RENAMED: Record<string, string> = { format: 'numberFormat' }

/** The fields of `o` are of `allowed`: an unknown one is refused, with the name now of a field by its old name (`format`). */
export function onlyKeys(o: Record<string, unknown>, path: string, allowed: readonly string[], hint = '') {
  for (const k of Object.keys(o))
    if (!allowed.includes(k)) {
      const now = Object.hasOwn(RENAMED, k) && allowed.includes(RENAMED[k]) ? ` (${k} is the old name of ${RENAMED[k]})` : ''
      throw new InputError(path, `unknown field ${JSON.stringify(k)} (expected ${allowed.join(', ')})${hint}${now}`)
    }
}

/** An options object's fields, as `onlyKeys`: an unknown one that another call takes says which call (`optionHint`). */
export const onlyOptions = (o: Record<string, unknown>, path: string, allowed: readonly string[], hint = '') => onlyKeys(o, path, allowed, hint + optionHint(o, allowed))
