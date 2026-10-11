/**
 * A workbook formula engine (the package root). The contract (types only) is also available as the
 * `./contract` subpath; the implementation compiles formulas to closures (or to JS text) and computes over
 * the reference graph.
 * The `./unstable/*` subpaths (internals, the synthetic book) are outside semver.
 */
import type { BookInput, CellDisplay, CellValue, ComputeOptions, ComputeResult, DisplayOptions, FormulaEngine, NameCheck, NameInput, NameIssue, NameOptions, NamesContext, ValidateContext, ValidateResult } from './contract.js'
import { createSession } from './session.js'
import { editStructure } from './structure/index.js'
import { epochOf } from './dates.js'
import { createState, debugSource } from './engine.js'
import { limitsBudget, run, withBudget } from './exec/budget.js'
import { computeSteps } from './compute.js'
import { currentClock } from './env.js'
import { checkBook, checkCellValue, checkComputeOptions, checkDisplayOptions, checkNameOptions, checkNames, checkNamesContext, checkString, checkValidateContext } from './input.js'
import { localeBooleans } from './locale/registry.js'
import { displayRules } from './locale/rules.js'
import { columnPoints, room } from './numfmt.js'
import { clockSerial } from './result.js'
import { FUNCTIONS } from './functions.js'
import { prepareNames } from './names.js'
import { sheetKey, validate, validateName as checkName } from './syntax.js'
import { parseEntry } from './entry/index.js'
import { entryWidth, fitColumns } from './layout/fit.js'
import { fromLegacy as legacyFormula, type LegacyContext as LegacyFormulaContext } from './legacy/convert.js'

export * from './contract.js'
export { FUNCTION_NAMES } from './functions.js'
export { CONFIG_VERSION } from './config.js'
export { createSession } from './session.js'
export { InputError } from './input.js'
export { CalcError } from './exec/budget.js'
export { createInlineCalc } from './async/inline.js'
export { enUS, toInvariant, toLocal, validateLocalFormula, type Locale, type Translation } from './locale/index.js'
export { registerLocale } from './locale/load.js'
export { LOCALES } from './locale/registry.js'
export { editStructure } from './structure/index.js'
export { copyFormula, type CopyContext } from './structure/copy.js'
export { parseEntry } from './entry/index.js'
export { entryText, needsApostrophe } from './entry/text.js'
export { entryWidth, fitColumns } from './layout/fit.js'
export * from './a1.js'
export { canonValue, errorText, sameValue, validateSheetName } from './helpers.js'
export { scanReferences } from './scan.js'
export { builtinFormatCode } from './builtin-format.js'
export { dateToSerial, serialToDate, type DateSystemOptions } from './date-serial.js'

/**
 * Computes the whole book. Without `opts.clock` the clock is the current moment and a random seed; the
 * result records it in `clock`, and passing that clock again reproduces the result. Input of the wrong
 * shape throws `InputError`; invalid formulas do not throw, they give error values with reasons in `diagnostics`.
 */
export function compute(book: BookInput, opts?: ComputeOptions): ComputeResult {
  const o = checkComputeOptions(opts)
  return withBudget(limitsBudget(o.limits), () => run(computeSteps(book, o)))
}

export function validateFormula(text: string, ctx?: ValidateContext): ValidateResult {
  const r = validate(checkString(text, 'text'), checkValidateContext(ctx))
  if (!r.ok) return r
  const { ast: _ast, ...rest } = r
  return rest
}

export function validateName(name: string, opts?: NameOptions): NameCheck {
  return checkName(checkString(name, 'name'), localeBooleans(checkNameOptions(opts).locale))
}

/** The issues of a book's names, by the same check as a book's (`prepareNames`), without the book. */
export function validateNames(names: NameInput[], ctx: NamesContext): NameIssue[] {
  const c = checkNamesContext(ctx)
  const sheetIdx = new Map(c.sheets.map((name, i) => [sheetKey(name), i]))
  const externalFunctions = (c.externalFunctions ?? []).map((f) => f.toUpperCase()).filter((f) => !FUNCTIONS[f])
  return prepareNames(sheetIdx, checkNames(names, 'names'), externalFunctions, c.tables ?? [], localeBooleans(c.locale)).issues
}

/** The `eval` mode's generated code of the book in evaluation order, for reading while debugging; not part of the contract. */
export function debugGeneratedCode(book: BookInput, opts?: ComputeOptions): string {
  const o = checkComputeOptions(opts)
  book = checkBook(book)
  const clock = o.clock ?? currentClock()
  return debugSource(createState(book, { now: clockSerial(clock), seed: clock.seed, externals: o.externals, execution: 'eval' }))
}

/**
 * A value as a cell shows it without a book (a cell being edited, say): by `numberFormat` (General without one), in
 * `locale` and `dateSystem`, in a column of `width` (no limit without one: a value is not narrowed to a column it was
 * not given), as `CellResult.text` and `color` show it.
 */
export function displayValue(value: CellValue, opts?: DisplayOptions): CellDisplay {
  const o = checkDisplayOptions(opts)
  return displayRules(o.locale, epochOf(o.dateSystem)).display(checkCellValue(value, 'value'), o.numberFormat, o.width === undefined ? Infinity : room(columnPoints(o.width)))
}

/** @deprecated `fromLegacy` reads formulas of files: `import { fromLegacy } from 'xlcalc/xlsx'`. */
export const fromLegacy: typeof legacyFormula = legacyFormula
/** @deprecated `import type { LegacyContext } from 'xlcalc/xlsx'`. */
export type LegacyContext = LegacyFormulaContext

/** @deprecated The name is `createSession`. */
export const openSession: typeof createSession = createSession
/** @deprecated The name is `editStructure`. */
export const restructure: typeof editStructure = editStructure

/** @deprecated The package's named exports are these functions: `import { compute, createSession } from 'xlcalc'`. */
export const engine: FormulaEngine = { compute, validateFormula, validateName, validateNames, createSession, editStructure, openSession, restructure, displayValue, parseEntry, fitColumns, entryWidth }
