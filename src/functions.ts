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
import { functions as text } from './functions/text.js'

export type { FunctionSpec } from './args.js'

const CATEGORIES = { database, date, engineering, financial, information, logical, lookup, math, statistical, text }

/**
 * Functions whose one-value arguments answer #VALUE! for a range of several cells instead of lifting it (an
 * array is lifted): Excel 365's Analysis ToolPak functions and the field of the database functions, as the
 * `lift` cases of `bench/coverage/sets/arrays.ts` measured; WORKDAY.INTL lifts a range of weekends.
 */
const REFUSE_RANGES: Record<string, FnDef['ranges']> = {
  ...Object.fromEntries(
    [
      'ACCRINT', 'ACCRINTM', 'AMORDEGRC', 'AMORLINC', 'BESSELI', 'BESSELJ', 'BESSELK', 'BESSELY', 'BIN2DEC', 'BIN2HEX', 'BIN2OCT', 'COMPLEX',
      'CONVERT', 'COUPDAYBS', 'COUPDAYSNC', 'COUPNCD', 'COUPNUM', 'COUPPCD', 'CUMIPMT', 'CUMPRINC', 'DAVERAGE', 'DCOUNT', 'DCOUNTA', 'DEC2BIN',
      'DEC2HEX', 'DEC2OCT', 'DELTA', 'DGET', 'DISC', 'DMAX', 'DMIN', 'DOLLARDE', 'DOLLARFR', 'DPRODUCT', 'DSTDEV', 'DSTDEVP', 'DSUM', 'DURATION',
      'DVAR', 'DVARP', 'EDATE', 'EFFECT', 'EOMONTH', 'ERF', 'ERF.PRECISE', 'ERFC', 'ERFC.PRECISE', 'FACTDOUBLE', 'FVSCHEDULE', 'GESTEP',
      'HEX2BIN', 'HEX2DEC', 'HEX2OCT', 'IMABS', 'IMAGINARY', 'IMARGUMENT', 'IMCONJUGATE', 'IMCOS', 'IMCOSH', 'IMCOT', 'IMCSC', 'IMCSCH',
      'IMDIV', 'IMEXP', 'IMLN', 'IMLOG10', 'IMLOG2', 'IMPOWER', 'IMREAL', 'IMSEC', 'IMSECH', 'IMSIN', 'IMSINH', 'IMSQRT', 'IMSUB', 'IMTAN',
      'INTRATE', 'ISEVEN', 'ISODD', 'MDURATION', 'MROUND', 'NETWORKDAYS', 'NETWORKDAYS.INTL', 'NOMINAL', 'OCT2BIN', 'OCT2DEC', 'OCT2HEX',
      'ODDFPRICE', 'ODDFYIELD', 'ODDLPRICE', 'ODDLYIELD', 'PRICE', 'PRICEDISC', 'PRICEMAT', 'QUOTIENT', 'RECEIVED', 'SERIESSUM', 'SQRTPI',
      'TBILLEQ', 'TBILLPRICE', 'TBILLYIELD', 'WEEKNUM', 'WORKDAY', 'XNPV', 'YEARFRAC', 'YIELD', 'YIELDDISC', 'YIELDMAT',
    ].map((n) => [n, 'error' as const]),
  ),
  'WORKDAY.INTL': [0, 1],
}

function assemble(): Record<string, FnDef> {
  const out: Record<string, FnDef> = {}
  for (const [category, defs] of Object.entries(CATEGORIES))
    for (const [name, def] of Object.entries(defs)) {
      if (out[name]) throw new Error(`function ${name} is defined twice (again in ${category})`)
      out[name] = name in REFUSE_RANGES ? { ...def, ranges: REFUSE_RANGES[name] } : def
    }
  return out
}

/** Implementations and argument kinds by function name. */
export const DEFS: Readonly<Record<string, FnDef>> = assemble()

/** Argument counts and range positions by function name, for the parser and the code generator. */
export const FUNCTIONS: Readonly<Record<string, FunctionSpec>> = Object.fromEntries(Object.entries(DEFS).map(([name, def]) => [name, specOf(name, def)]))

export const FUNCTION_NAMES: readonly string[] = Object.keys(FUNCTIONS).sort()
