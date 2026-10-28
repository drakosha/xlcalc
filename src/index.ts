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
import * as live from './live.js'
import * as locales from './locale/index.js'
import * as load from './locale/load.js'
import * as packaged from './locale/package.js'
import * as entry from './entry/index.js'
import * as entryTexts from './entry/text.js'
import * as fit from './layout/fit.js'
import * as heights from './layout/heights.js'
import * as helpers from './helpers.js'
import * as styleTable from './style-table.js'
import * as builtin from './builtin-format.js'
import * as serials from './date-serial.js'
import { strict } from './strict-args.js'
import { withDeclaredArities } from './declared-arity.js'
import * as signatures from './signatures.js'
import { epochOf } from './dates.js'
import { createState, debugSource } from './engine.js'
import { limitsBudget, run, withBudget } from './exec/budget.js'
import { computeSteps } from './compute.js'
import { currentClock } from './env.js'
import { checkBook, checkCellValue, checkComputeOptions, checkDisplayOptions, checkNameOptions, checkNames, checkNamesContext, checkValidateContext } from './input.js'
import { checkString } from './input/basic.js'
import { localeBooleans } from './locale/registry.js'
import { displayRules } from './locale/rules.js'
import { columnPoints, room } from './numfmt.js'
import { clockOptions } from './result.js'
import { externalNames } from './functions.js'
import { prepareNames } from './names.js'
import { sheetKey, validate, validateName as checkName } from './syntax.js'

export * from './contract.js'
export { FUNCTION_NAMES } from './functions.js'
export { FUNCTION_CATEGORIES } from './function-info.js'
export { CONFIG_VERSION } from './config.js'
export { InputError } from './input-error.js'
export { CalcError } from './exec/budget.js'
export { enUS } from './locale/index.js'
export { LOCALES } from './locale/registry.js'
export * from './a1.js'
export { callAt, compareValues, columnWidthFromChars, columnWidthFromPoints, columnWidthToChars, columnWidthToPoints, DEFAULT_ROW_HEIGHT, escapeColumn, findMatcher, findReplacer, MAX_OUTLINE, MAX_ROW_HEIGHT, nameAt, numberFormatCategory, numberFormatChoices, formatClipboard, parseClipboard, refAt, scanReferences, stepDecimals, valueComparator } from './text.js'

// the functions of the modules, each refusing options it does not take (`strict-args.ts`); a bundle keeps those it uses;
// those that check a formula's calls count by the functions' declared arguments in a bundle without the engine (`declared-arity.ts`)
export const createSession = /* @__PURE__ */ strict(sessions.createSession)
export const createInlineCalc = /* @__PURE__ */ strict(inline.createInlineCalc)
export const liveBook = /* @__PURE__ */ strict(live.liveBook)
export const viewportOf = /* @__PURE__ */ strict(saved.viewportOf)
// in the root alone: its names would take `xlcalc/text` past its 55 KB (`text-bundle.test.ts`)
export const functionSignature = /* @__PURE__ */ strict(signatures.functionSignature)
export const toLocal = /* @__PURE__ */ strict(locales.toLocal)
export const toInvariant = /* @__PURE__ */ strict(locales.toInvariant)
export const validateLocalFormula = /* @__PURE__ */ strict(/* @__PURE__ */ withDeclaredArities(locales.validateLocalFormula))
export const registerLocale = /* @__PURE__ */ strict(load.registerLocale)
export const loadLocale = /* @__PURE__ */ strict(packaged.loadLocale)
// a locale not loaded is refused with the module to import, where registerLocale is at hand
load.rootLoaded()
export const editStructure = /* @__PURE__ */ strict(structure.editStructure)
export const copyFormula = /* @__PURE__ */ strict(copy.copyFormula)
export const parseEntry = /* @__PURE__ */ strict(/* @__PURE__ */ withDeclaredArities(entry.parseEntry))
export const entryText = /* @__PURE__ */ strict(/* @__PURE__ */ withDeclaredArities(entryTexts.entryText))
export const needsApostrophe = /* @__PURE__ */ strict(/* @__PURE__ */ withDeclaredArities(entryTexts.needsApostrophe))
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
 * Calculates the entire workbook. Without `opts.clock`, the clock is the current time and a random seed; the
 * result records it in `clock`, and passing that clock again reproduces the result. Input of the wrong
 * shape throws `InputError`; invalid formulas do not throw, they produce error values with reasons in `diagnostics`.
 */
export const compute = /* @__PURE__ */ strict(function compute(book: BookInput, opts?: ComputeOptions): ComputeResult {
  const o = checkComputeOptions(opts)
  return withBudget(limitsBudget(o.limits), () => run(computeSteps(book, o)))
})

/**
 * Checks formula text as a cell stores it (without `=`, en-US syntax): returns its canonical text and the functions,
 * sheets, names, external workbooks and tables it refers to, or why and where it fails to parse (`ValidateResult`).
 * With `ctx`, the sheets, names and tables it refers to are checked against the workbook's.
 */
export const validateFormula = /* @__PURE__ */ strict(/* @__PURE__ */ withDeclaredArities(function validateFormula(text: string, ctx?: ValidateContext): ValidateResult {
  const r = validate(checkString(text, 'text'), checkValidateContext(ctx))
  if (!r.ok) return r
  const { ast: _ast, ...rest } = r
  return rest
}))

/**
 * Checks a name (`NameInput.name`) against Excel's rules in a workbook of `opts.locale` (en-US without one; a locale's
 * words for TRUE and FALSE are also rejected, ИСТИНА and ЛОЖЬ in ru-RU).
 */
export const validateName = /* @__PURE__ */ strict(function validateName(name: string, opts?: NameOptions): NameCheck {
  return checkName(checkString(name, 'name'), localeBooleans(checkNameOptions(opts).locale))
})

/** The issues of a workbook's names, found by the same check a workbook gets (`prepareNames`), without the workbook. */
export const validateNames = /* @__PURE__ */ strict(function validateNames(names: NameInput[], ctx: NamesContext): NameIssue[] {
  const c = checkNamesContext(ctx)
  const sheetIdx = new Map(c.sheets.map((name, i) => [sheetKey(name), i]))
  return prepareNames(sheetIdx, checkNames(names, 'names'), externalNames(c.externalFunctions), c.tables ?? [], localeBooleans(c.locale), c.externalBooks).issues
})

/** The code that `eval` mode generates for the workbook, in evaluation order, for reading while debugging; not part of the contract. */
export const debugGeneratedCode = /* @__PURE__ */ strict(function debugGeneratedCode(book: BookInput, opts?: ComputeOptions): string {
  const o = checkComputeOptions(opts)
  book = checkBook(book)
  const clock = o.clock ?? currentClock()
  return debugSource(createState(book, { ...clockOptions(clock), externals: o.externals, execution: 'eval' }))
})

/**
 * A value as a cell would display it, without a workbook (for example, a cell being edited): in `numberFormat` (General
 * without one), in `locale` and `dateSystem`, in a column of `width` (no limit without one: a value is not shortened to
 * fit a column it was not given), as `CellResult.text` and `color` display it.
 */
export const displayValue = /* @__PURE__ */ strict(function displayValue(value: CellValue, opts?: DisplayOptions): CellDisplay {
  const o = checkDisplayOptions(opts)
  return displayRules(o.locale, epochOf(o.dateSystem)).display(checkCellValue(value, 'value'), o.numberFormat, o.width === undefined ? Infinity : room(columnPoints(o.width)))
})
