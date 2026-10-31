/**
 * What a call does not take is refused, not ignored: an options object or a callback past a call's own arguments
 * throws `InputError` naming the call that takes those options (`getCells(sheet, range, { formulaLocal: true })`:
 * formulaLocal is an option of getViewport), and so does an unknown field of an options object, by its check. An array,
 * a number or undefined past them passes, as `map` and `forEach` give a callback its index and array. The functions and
 * methods of the API are given out through `strict` and `strictMethods`; the module needs nothing of the engine.
 */
import { InputError, what } from './input-error.js'

/** The calls that type text into cells. */
const TYPING = 'setCellText, setCellTexts, fillCellText and their previews'
const VIEWPORT = 'getViewport'
const FIND = 'findCells and replaceInCells'
const COMPUTE = 'compute and createSession'
const CSV = 'readCsv and writeCsv'

/** The options of the calls that type text beside `baseRevision` (`CellTextOptions`), and of an entry into a selection (`FillTextOptions`). */
export const TEXT_OPTIONS = ['widen', 'refuseInvalid', 'inPlace'] as const
export const FILL_OPTIONS = ['widen', 'refuseInvalid', 'inPlace', 'active'] as const
/** The options of `readCells` beside `baseRevision` (`ReadCellsOptions`): a page's, `fields` and `rows`. */
export const READ_OPTIONS = ['maxCells', 'cursor', 'count', 'fields', 'rows'] as const
/** The options of `findCells` beside `readCells`' (`FindCellsOptions`). */
export const FIND_OPTIONS = ['in', 'caseSensitive', 'wildcards', 'wholeCell', 'byColumns', 'skipHidden'] as const
/** All the options of `findCells` beside `baseRevision`: `readCells`', Find's, and the cell it begins after. */
export const FIND_CALL_OPTIONS = [...READ_OPTIONS, ...FIND_OPTIONS, 'after'] as const
/** The options of `replaceInCells` beside `baseRevision` (`ReplaceOptions`): Find's but what it reads and the order, and `rows`. */
export const REPLACE_OPTIONS = ['caseSensitive', 'wildcards', 'wholeCell', 'skipHidden', 'rows'] as const

/** The API's options by name and the calls that take them, for a refusal to say where a stray option goes. */
const TAKEN_BY: Record<string, string> = {
  formulaLocal: VIEWPORT,
  diagnostics: VIEWPORT,
  maxCells: VIEWPORT,
  rows: 'getViewport, readCells, findCells, replaceInCells and summarizeRanges',
  columns: VIEWPORT,
  editText: VIEWPORT,
  appearance: 'viewportOf',
  depth: 'explain',
  maxNodes: 'explain',
  expand: 'explain',
  errorDepth: 'explain',
  widen: TYPING,
  refuseInvalid: TYPING,
  inPlace: TYPING,
  active: 'fillCellText and previewFillCellText',
  caseSensitive: FIND,
  wildcards: FIND,
  wholeCell: FIND,
  byColumns: 'findCells',
  skipHidden: FIND,
  after: 'findCells',
  baseRevision: "a session's calls that change its workbook",
  table: 'toLocal and toInvariant',
  clock: COMPUTE,
  externals: COMPUTE,
  execution: COMPUTE,
  display: COMPUTE,
  cached: COMPUTE,
  unknownFunctions: COMPUTE,
  limits: 'compute, createSession, the calculators and readXlsx',
  live: 'createSession',
  undo: 'createSession',
  values: 'writeXlsx',
  mode: 'writeXlsx',
  delimiter: CSV,
  encoding: CSV,
  formulas: 'readCsv',
  codesAsText: 'readCsv',
  bom: 'writeCsv',
}

/** Where the fields of an options object that `allowed` does not hold go, as a refusal's hint: ` (formulaLocal is an option of getViewport)`. */
export function optionHint(o: object, allowed: readonly string[] = []): string {
  const hints = Object.keys(o).flatMap((k) => (allowed.includes(k) || !Object.hasOwn(TAKEN_BY, k) ? [] : [`${k} is an option of ${TAKEN_BY[k]}`]))
  return hints.length ? ` (${hints.join('; ')})` : ''
}

/** An argument that is options or a callback: an object other than an array, or a function. */
const optionLike = (v: unknown) => (typeof v === 'object' && v !== null && !Array.isArray(v)) || typeof v === 'function'

const isAsync = (f: object) => (f as { [Symbol.toStringTag]?: string })[Symbol.toStringTag] === 'AsyncFunction'

/**
 * `f` as the API gives it out, by its name: an options object or a callback past its first `max` arguments (its own by
 * default) is refused, by a rejected promise for a call that answers one (`async`, an async function's by default).
 */
export function strict<F extends (...args: never[]) => unknown>(f: F, name = f.name, max = f.length, async = isAsync(f)): F {
  const checked = function (this: unknown, ...args: unknown[]) {
    for (let i = max; i < args.length; i++) {
      const v = args[i]
      if (!optionLike(v)) continue
      const e = new InputError(name, `takes ${max} argument${max === 1 ? '' : 's'}, got ${what(v)} as argument ${i + 1}${typeof v === 'object' ? optionHint(v as object) : ''}`)
      if (async) return Promise.reject(e)
      throw e
    }
    return (f as unknown as (...a: unknown[]) => unknown).apply(this, args)
  }
  return Object.defineProperties(checked, { name: { value: name }, length: { value: max } }) as unknown as F
}

/**
 * `f`, by its name and number of parameters, calling `setup` before each call (what an entry loads for the call: the
 * declared arguments, the groups of functions); an async function stays one, so that `strict` refuses its options by a
 * rejected promise.
 */
export function before<F extends (...args: never[]) => unknown>(f: F, setup: () => void): F {
  const call = function (this: unknown, ...args: never[]) {
    setup()
    return f.apply(this, args)
  }
  const g = isAsync(f)
    ? async function (this: unknown, ...args: never[]) {
        return call.apply(this, args)
      }
    : call
  return Object.defineProperties(g, { name: { value: f.name }, length: { value: f.length } }) as unknown as F
}

/**
 * An API object (a session, a calculator) with each of its methods `strict`, by its own name and parameters; getters
 * stay. `async`: its methods answer promises, and a refusal is a rejected one.
 */
export function strictMethods<T extends object>(api: T, async = false): T {
  const o = api as Record<string, unknown>
  for (const k of Object.keys(o)) {
    const f = Object.getOwnPropertyDescriptor(o, k)!.value as unknown
    if (typeof f === 'function') o[k] = strict(f as (...a: never[]) => unknown, k, f.length, async || isAsync(f))
  }
  return api
}
