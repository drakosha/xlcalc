/**
 * A workbook formula engine (the package root). The contract (types only) is also available as the
 * `./contract` subpath; the implementation compiles formulas to closures (or to JS text) and computes over
 * the reference graph.
 * The `./unstable/*` subpaths (internals, the synthetic book) are outside semver.
 */
import type { BookInput, CellDisplay, CellValue, ComputeOptions, ComputeResult, DisplayOptions, NameCheck, NameInput, NameIssue, NameOptions, NamesContext, ValidateContext, ValidateResult } from './contract.js'
import * as sessions from './session.js'
import * as saved from './session/saved.js'
import * as structure from './structure/index.js'
import * as copy from './structure/copy.js'
import * as inline from './async/inline.js'
import * as locales from './locale/index.js'
import * as load from './locale/load.js'
import * as entry from './entry/index.js'
import * as entryTexts from './entry/text.js'
import * as fit from './layout/fit.js'
import * as heights from './layout/heights.js'
import * as helpers from './helpers.js'
import * as styleTable from './style-table.js'
import * as builtin from './builtin-format.js'
import * as serials from './date-serial.js'
import { strict } from './strict-args.js'
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

export * from './contract.js'
export { FUNCTION_NAMES } from './functions.js'
export { CONFIG_VERSION } from './config.js'
export { InputError } from './input.js'
export { CalcError } from './exec/budget.js'
export { enUS, type Locale, type Translation } from './locale/index.js'
export { LOCALES } from './locale/registry.js'
export type { CopyContext } from './structure/copy.js'
export * from './a1.js'
export { escapeColumn, numberFormatCategory, numberFormatChoices, parseClipboard, scanReferences } from './text.js'
export type { DateSystemOptions } from './date-serial.js'

// the functions of the modules, each refusing options it does not take (`strict-args.ts`); a bundle keeps those it uses
export const createSession = /* @__PURE__ */ strict(sessions.createSession)
export const createInlineCalc = /* @__PURE__ */ strict(inline.createInlineCalc)
export const viewportOf = /* @__PURE__ */ strict(saved.viewportOf)
export const toLocal = /* @__PURE__ */ strict(locales.toLocal)
export const toInvariant = /* @__PURE__ */ strict(locales.toInvariant)
export const validateLocalFormula = /* @__PURE__ */ strict(locales.validateLocalFormula)
export const registerLocale = /* @__PURE__ */ strict(load.registerLocale)
// a locale not loaded is refused with the module to import, where registerLocale is at hand
load.rootLoaded()
export const editStructure = /* @__PURE__ */ strict(structure.editStructure)
export const copyFormula = /* @__PURE__ */ strict(copy.copyFormula)
export const parseEntry = /* @__PURE__ */ strict(entry.parseEntry)
export const entryText = /* @__PURE__ */ strict(entryTexts.entryText)
export const needsApostrophe = /* @__PURE__ */ strict(entryTexts.needsApostrophe)
export const entryWidth = /* @__PURE__ */ strict(fit.entryWidth)
export const fitColumns = /* @__PURE__ */ strict(fit.fitColumns)
export const fitRows = /* @__PURE__ */ strict(heights.fitRows)
export const canonValue = /* @__PURE__ */ strict(helpers.canonValue)
export const errorText = /* @__PURE__ */ strict(helpers.errorText)
export const sameValue = /* @__PURE__ */ strict(helpers.sameValue)
export const validateSheetName = /* @__PURE__ */ strict(helpers.validateSheetName)
export const builtinFormatCode = /* @__PURE__ */ strict(builtin.builtinFormatCode)
export const dateToSerial = /* @__PURE__ */ strict(serials.dateToSerial)
export const serialToDate = /* @__PURE__ */ strict(serials.serialToDate)
export const internStyle = /* @__PURE__ */ strict(styleTable.internStyle)

/**
 * Computes the whole book. Without `opts.clock` the clock is the current moment and a random seed; the
 * result records it in `clock`, and passing that clock again reproduces the result. Input of the wrong
 * shape throws `InputError`; invalid formulas do not throw, they give error values with reasons in `diagnostics`.
 */
export const compute = /* @__PURE__ */ strict(function compute(book: BookInput, opts?: ComputeOptions): ComputeResult {
  const o = checkComputeOptions(opts)
  return withBudget(limitsBudget(o.limits), () => run(computeSteps(book, o)))
})

export const validateFormula = /* @__PURE__ */ strict(function validateFormula(text: string, ctx?: ValidateContext): ValidateResult {
  const r = validate(checkString(text, 'text'), checkValidateContext(ctx))
  if (!r.ok) return r
  const { ast: _ast, ...rest } = r
  return rest
})

/**
 * Excel's rules for a name (`NameInput.name`) in a book of `opts.locale` (en-US without one; a locale's words for TRUE
 * and FALSE are refused too, ИСТИНА and ЛОЖЬ in ru-RU).
 */
export const validateName = /* @__PURE__ */ strict(function validateName(name: string, opts?: NameOptions): NameCheck {
  return checkName(checkString(name, 'name'), localeBooleans(checkNameOptions(opts).locale))
})

/** The issues of a book's names, by the same check as a book's (`prepareNames`), without the book. */
export const validateNames = /* @__PURE__ */ strict(function validateNames(names: NameInput[], ctx: NamesContext): NameIssue[] {
  const c = checkNamesContext(ctx)
  const sheetIdx = new Map(c.sheets.map((name, i) => [sheetKey(name), i]))
  const externalFunctions = (c.externalFunctions ?? []).map((f) => f.toUpperCase()).filter((f) => !FUNCTIONS[f])
  return prepareNames(sheetIdx, checkNames(names, 'names'), externalFunctions, c.tables ?? [], localeBooleans(c.locale), c.externalBooks).issues
})

/** The `eval` mode's generated code of the book in evaluation order, for reading while debugging; not part of the contract. */
export const debugGeneratedCode = /* @__PURE__ */ strict(function debugGeneratedCode(book: BookInput, opts?: ComputeOptions): string {
  const o = checkComputeOptions(opts)
  book = checkBook(book)
  const clock = o.clock ?? currentClock()
  return debugSource(createState(book, { now: clockSerial(clock), seed: clock.seed, externals: o.externals, execution: 'eval' }))
})

/**
 * A value as a cell shows it without a book (a cell being edited, say): by `numberFormat` (General without one), in
 * `locale` and `dateSystem`, in a column of `width` (no limit without one: a value is not narrowed to a column it was
 * not given), as `CellResult.text` and `color` show it.
 */
export const displayValue = /* @__PURE__ */ strict(function displayValue(value: CellValue, opts?: DisplayOptions): CellDisplay {
  const o = checkDisplayOptions(opts)
  return displayRules(o.locale, epochOf(o.dateSystem)).display(checkCellValue(value, 'value'), o.numberFormat, o.width === undefined ? Infinity : room(columnPoints(o.width)))
})
