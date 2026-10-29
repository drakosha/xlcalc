/**
 * Arguments left out between commas (`f(a,,c)`), as Excel 16.113 takes them (the `omit` sets of
 * `bench/coverage/sets/`). The rule (`args.ts`, `specOf`): a blank, read by the argument's kind as 0, "" or FALSE
 * (`ROUND(1.5,)` is 2, `LEFT("abc",)` "", `VLOOKUP(4,t,2,)` an exact match); #VALUE! where the function takes an
 * array or a reference (kinds G, M, R: `CORREL(,A1:A8)`); in the former Analysis ToolPak's functions the argument
 * not given, a required one #N/A (`EDATE(,1)`). `OMITTED` lists where a function takes it otherwise,
 * `NOT_OMITTED` where Excel does not load a formula that leaves it out.
 */
import type { Omission } from '../args.js'
import type { Positions } from '../arity.js'
import { functions as database } from './database.js'

const DEFAULT = 'default' as const

/** The function's default, as if not given, at these positions. */
const defaults = (...at: number[]): Omission => ({ at: Object.fromEntries(at.map((i) => [i, DEFAULT])) })

export const OMITTED: Readonly<Record<string, Omission>> = {
  // the par value of $1,000, though before a required argument
  ACCRINT: defaults(4),
  ACCRINTM: defaults(3),
  ADDRESS: defaults(2, 3, 4),
  'BINOM.DIST.RANGE': defaults(3),
  'CEILING.MATH': defaults(1),
  'CEILING.PRECISE': defaults(1),
  // counted without a field: every record
  DCOUNT: defaults(1),
  DCOUNTA: defaults(1),
  DROP: defaults(1, 2),
  // the optional upper limit too
  ERF: { at: { 1: 'NA' } },
  // the padding is a blank, the array too
  EXPAND: { at: { 0: 'blank', 1: DEFAULT, 2: DEFAULT } },
  'F.TEST': { at: { 0: 'VALUE', 1: 'VALUE' } },
  FILTER: defaults(2),
  FIXED: defaults(1),
  'FLOOR.MATH': defaults(1),
  'FLOOR.PRECISE': defaults(1),
  FTEST: { at: { 0: 'VALUE', 1: 'VALUE' } },
  GROWTH: defaults(1, 2, 3),
  // a lookup in a blank finds nothing (#N/A)
  HLOOKUP: { at: { 1: 'blank' } },
  HSTACK: { rest: 'VALUE' },
  'ISO.CEILING': defaults(1),
  LINEST: defaults(1, 2),
  LOGEST: defaults(1, 2),
  MATCH: { at: { 1: 'blank' } },
  MIRR: { at: { 0: 'VALUE' } },
  PERCENTRANK: defaults(2),
  'PERCENTRANK.EXC': defaults(2),
  'PERCENTRANK.INC': defaults(2),
  PROB: defaults(3),
  // a value left out is not multiplied
  PRODUCT: { rest: DEFAULT },
  RANDARRAY: defaults(0, 1, 2, 3),
  SEQUENCE: defaults(0, 1, 2, 3),
  SORT: defaults(1, 2),
  // a blank by-array keeps the order
  SORTBY: { rest: ['blank', DEFAULT] },
  TAKE: defaults(1, 2),
  TEXTAFTER: defaults(2, 3, 4, 5),
  TEXTBEFORE: defaults(2, 3, 4, 5),
  // empty texts are ignored
  TEXTJOIN: { at: { 1: 'true' } },
  TEXTSPLIT: defaults(1, 2),
  TREND: defaults(1, 2, 3),
  UNIQUE: { at: { 0: 'blank' } },
  VLOOKUP: { at: { 1: 'blank' } },
  VSTACK: { rest: 'VALUE' },
  WRAPCOLS: defaults(2),
  WRAPROWS: defaults(2),
  XLOOKUP: defaults(3, 4, 5),
  XMATCH: defaults(2, 3),
}

/**
 * Positions where Excel does not load a formula with the argument left out (they take only a reference): the
 * ranges of the criteria functions, RANK's list, SUBTOTAL's references. AGGREGATE's references and the database
 * functions' ranges, also references only, load and are #VALUE!.
 */
export const NOT_OMITTED: Readonly<Record<string, Positions>> = {
  COUNTIF: { at: [0] },
  COUNTIFS: { from: 0, even: true },
  SUMIF: { at: [0, 2] },
  AVERAGEIF: { at: [0, 2] },
  SUMIFS: { at: [0], from: 1, odd: true },
  AVERAGEIFS: { at: [0], from: 1, odd: true },
  MAXIFS: { at: [0], from: 1, odd: true },
  MINIFS: { at: [0], from: 1, odd: true },
  RANK: { at: [1] },
  'RANK.EQ': { at: [1] },
  'RANK.AVG': { at: [1] },
  SUBTOTAL: { from: 1 },
}

/**
 * Positions that take a reference of an open book: another workbook's range there is `#VALUE!`, as Excel computes it
 * over a closed book (the criteria functions' ranges, COUNTBLANK's, a database function's range;
 * the sets `links/basics-held` and `links/functions-held`). RANK's list, SUBTOTAL's and AGGREGATE's references read
 * the stored values.
 */
export const OPEN_REFERENCE: Readonly<Record<string, Positions>> = {
  ...Object.fromEntries(Object.entries(NOT_OMITTED).filter(([name]) => !name.startsWith('RANK') && name !== 'SUBTOTAL')),
  COUNTBLANK: { at: [0] },
  ...Object.fromEntries(Object.keys(database).map((f) => [f, { at: [0] }])),
}
