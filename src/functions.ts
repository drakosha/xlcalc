/**
 * The function whitelist, assembled from the category files in `src/functions/`. A function gets into its
 * category's registry only when all of its cases in the Excel acceptance test (`test/excel.test.ts`) match
 * Excel en-US, or with its own semantics tests where the measurement has no cases.
 */
import { specOf, type FnDef, type FunctionSpec } from './args.js'
import { functions as database } from './functions/database.js'
import { functions as date } from './functions/date.js'
import { functions as engineering } from './functions/engineering.js'
import { functions as financial } from './functions/financial.js'
import { functions as information } from './functions/information.js'
import { functions as logical } from './functions/logical.js'
import { functions as lookup } from './functions/lookup.js'
import { functions as math } from './functions/math.js'
import { functions as statistical } from './functions/statistical.js'
import { NOT_OMITTED, OMITTED } from './functions/omitted.js'
import { functions as text } from './functions/text.js'

export type { FunctionSpec } from './args.js'

const CATEGORIES = { database, date, engineering, financial, information, logical, lookup, math, statistical, text }

/**
 * The former Analysis ToolPak's functions, with NETWORKDAYS.INTL and WORKDAY.INTL. Excel 365 takes an argument
 * left out of them as not given, a required one as #N/A (the `omit` set of `bench/coverage/sets/`); and refuses a
 * range of several cells in their one-value arguments (#VALUE!; an array is lifted), as the `lift` cases of
 * `bench/coverage/sets/arrays.ts` measured, except the ones in `LIFTING` and WORKDAY.INTL's weekends.
 */
const ANALYSIS = new Set([
  'ACCRINT', 'ACCRINTM', 'AMORDEGRC', 'AMORLINC', 'BESSELI', 'BESSELJ', 'BESSELK', 'BESSELY', 'BIN2DEC', 'BIN2HEX', 'BIN2OCT', 'COMPLEX',
  'CONVERT', 'COUPDAYBS', 'COUPDAYSNC', 'COUPNCD', 'COUPNUM', 'COUPPCD', 'CUMIPMT', 'CUMPRINC', 'DEC2BIN', 'DEC2HEX', 'DEC2OCT', 'DELTA',
  'DISC', 'DOLLARDE', 'DOLLARFR', 'DURATION', 'EDATE', 'EFFECT', 'EOMONTH', 'ERF', 'ERF.PRECISE', 'ERFC', 'ERFC.PRECISE', 'FACTDOUBLE',
  'FVSCHEDULE', 'GCD', 'GESTEP', 'HEX2BIN', 'HEX2DEC', 'HEX2OCT', 'IMABS', 'IMAGINARY', 'IMARGUMENT', 'IMCONJUGATE', 'IMCOS', 'IMCOSH',
  'IMCOT', 'IMCSC', 'IMCSCH', 'IMDIV', 'IMEXP', 'IMLN', 'IMLOG10', 'IMLOG2', 'IMPOWER', 'IMPRODUCT', 'IMREAL', 'IMSEC', 'IMSECH', 'IMSIN',
  'IMSINH', 'IMSQRT', 'IMSUB', 'IMSUM', 'IMTAN', 'INTRATE', 'ISEVEN', 'ISODD', 'LCM', 'MDURATION', 'MROUND', 'MULTINOMIAL', 'NETWORKDAYS',
  'NETWORKDAYS.INTL', 'NOMINAL', 'OCT2BIN', 'OCT2DEC', 'OCT2HEX', 'ODDFPRICE', 'ODDFYIELD', 'ODDLPRICE', 'ODDLYIELD', 'PRICE', 'PRICEDISC',
  'PRICEMAT', 'QUOTIENT', 'RANDBETWEEN', 'RECEIVED', 'SERIESSUM', 'SQRTPI', 'TBILLEQ', 'TBILLPRICE', 'TBILLYIELD', 'WEEKNUM', 'WORKDAY',
  'WORKDAY.INTL', 'XIRR', 'XNPV', 'YEARFRAC', 'YIELD', 'YIELDDISC', 'YIELDMAT',
])

/** Analysis ToolPak functions that lift a range like an array: those over lists of numbers (RANDBETWEEN, volatile, is not measured). */
const LIFTING = ['GCD', 'IMPRODUCT', 'IMSUM', 'LCM', 'MULTINOMIAL', 'RANDBETWEEN', 'XIRR']

/** The database functions refuse a range of several cells as their field, as the `lift` cases measured. */
const DATABASE = ['DAVERAGE', 'DCOUNT', 'DCOUNTA', 'DGET', 'DMAX', 'DMIN', 'DPRODUCT', 'DSTDEV', 'DSTDEVP', 'DSUM', 'DVAR', 'DVARP']

const REFUSE_RANGES: Record<string, FnDef['ranges']> = {
  ...Object.fromEntries([...[...ANALYSIS].filter((n) => !LIFTING.includes(n)), ...DATABASE].map((n) => [n, 'error' as const])),
  'WORKDAY.INTL': [0, 1],
}

/** A definition with the tables of this file and of `functions/omitted.ts`. */
function measured(name: string, def: FnDef): FnDef {
  const omit = ANALYSIS.has(name) ? { ...OMITTED[name], analysis: true } : OMITTED[name]
  return { ...def, ...(name in REFUSE_RANGES && { ranges: REFUSE_RANGES[name] }), ...(omit && { omit }), ...(name in NOT_OMITTED && { notOmitted: NOT_OMITTED[name] }) }
}

function assemble(): Record<string, FnDef> {
  const out: Record<string, FnDef> = {}
  for (const [category, defs] of Object.entries(CATEGORIES))
    for (const [name, def] of Object.entries(defs)) {
      if (out[name]) throw new Error(`function ${name} is defined twice (again in ${category})`)
      out[name] = measured(name, def)
    }
  return out
}

/** Implementations and argument kinds by function name. */
export const DEFS: Readonly<Record<string, FnDef>> = assemble()

/** Argument counts and range positions by function name, for the parser and the code generator. */
export const FUNCTIONS: Readonly<Record<string, FunctionSpec>> = Object.fromEntries(Object.entries(DEFS).map(([name, def]) => [name, specOf(name, def)]))

/**
 * Whitelisted functions that are syntax rather than calls: LET's names are bound by the parser
 * (`syntax/parse.ts`) and its values computed by the walk (`exec/walk.ts`).
 */
export const SYNTAX: readonly string[] = ['LET']

/** A whitelisted function, a call or syntax. */
export const isFunction = (name: string): boolean => name in FUNCTIONS || SYNTAX.includes(name)

export const FUNCTION_NAMES: readonly string[] = [...Object.keys(FUNCTIONS), ...SYNTAX].sort()
