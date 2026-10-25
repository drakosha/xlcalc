/**
 * The shared argument layer. Every function declares the kinds of its arguments (`FnDef.args`), and the
 * runtime coerces the arguments by kind before the call: a function never repeats the coercion, it names
 * the group it belongs to. An error in a coerced argument is the answer, without calling the function.
 *
 * The kinds follow the rules derived from the coverage measurement against Excel en-US
 * (`docs/coverage-formulajs.md`, "Proposal"):
 *
 * 1. Coercion by group, not one rule for all:
 *    - `n` number: a boolean is 1 or 0, a blank 0, text that reads as a number is that number, other text
 *      #VALUE!; `i` the same, truncated to an integer; `w` the same, read with Excel's tolerance of whole
 *      numbers (`whole` in values.ts: less than 2^-22 short of an integer is that integer; LEFT, INDEX,
 *      ROUND's digits, DATE and the others the `near-integer` set measured), while `i` truncates as MID, REPT,
 *      CHOOSE, ROUNDDOWN's digits and the statistical counts do; `x` the same as `n`, but a boolean is #VALUE! (functions
 *      of the former Analysis ToolPak: engineering and part of the financial ones);
 *    - `s` text (a number in the General format, TRUE/FALSE, by the book's locale); `b` boolean (text
 *      "TRUE"/"FALSE" too in en-US, "ИСТИНА"/"ЛОЖЬ" in ru-RU); `v` any value (an error stops the call); `a` any value, errors
 *      included (ISERROR, IFS);
 *    - lists of numbers from ranges and arguments: `N` the SUM rule (a range gives its numbers, a direct
 *      argument is coerced like `n`), `A` the MAXA rule (a range gives numbers, booleans as 1/0 and
 *      text as 0; a direct argument is coerced like `n`), `Y` the same for AVERAGEA, `Z` strict (text or a boolean in a range and a
 *      boolean argument are #VALUE!: GCD, LCM, MULTINOMIAL);
 *    - `V` values in order, errors included, each range as one segment, so a function can tell a range
 *      from a direct argument (COUNTA counts every direct argument); `G` a range as a grid with its
 *      referenced size, errors inside kept; `D` a reference kept as it is, anything else as `G` (the database
 *      functions read a reference's columns through the function context); `C` the same as `G`, but a reference to
 *      one cell is read first, its error the answer and a blank cell #VALUE!, as Excel reads a function's arrays
 *      before comparing them (SUMX2MY2, CORREL, CHISQ.TEST; an array of one is an array); `M` a numeric grid (an
 *      error inside is the answer); `R` the argument itself, a reference or a value (criteria functions, which work on the reference through
 *      the function context); `r` one value that may be a reference (the function gets a reference to one
 *      or more cells as it is; an array is lifted); `p` any value, reference or array, passed as it is
 *      unless the call is lifted, a one-cell reference as its value (IFERROR's alternative, FILTER's
 *      `if_empty`); `P` the same with a reference passed on as a reference, one cell too, for a function
 *      that answers it (CHOOSE, the values of IF, IFS and SWITCH, HSTACK: `SUM(IF(TRUE,A2))` skips text in
 *      A2, as Excel does); where the answer goes to a one-value position it is read as a value.
 * 2. Date and time text in a date argument (`d`, `t`, `h` a time of day) is read as Excel reads it in the
 *    book's locale (`Cx.rules`; en-US: `2025-01-31`, `1/31/2025`, `18:30`, see `dates.ts`); a date before 0 or after
 *    9999-12-31 is #NUM!. A function gets dates in the 1900 date system, whatever the book's (`FnDef.date` moves
 *    its date answer back). Text, numbers and booleans in every kind are coerced by the locale's rules too.
 * 3. Empty input: `FnDef.empty` names the error of a function whose lists (`N`, `A`, `Y`, `Z`) hold no
 *    numbers at all (AVERAGE #DIV/0!, MEDIAN #NUM!, MODE #N/A); a blank cell as a whole answer is null,
 *    which a cell shows as 0.
 * 4. Paired arrays (`pairs`): pairs where either value is not a number are dropped; arrays with different
 *    numbers of values are #N/A.
 * 5. The answer: a non-finite number is #NUM!, one below the smallest normal number is 0 (`fitNumber`);
 *    radix digits are upper-case and complex text is written as Excel writes it by the functions themselves
 *    (`functions/engineering/radix.ts`, `complex.ts`).
 * 6. Arrays (`docs/arrays-incell.md`): the list and grid kinds take an array as they take a range (`A`
 *    takes only an array's numbers, `Y` counts its text as 0); a one-value kind given an array or a range
 *    of several cells is lifted element by element (`src/arrays/call.ts`), except where `FnDef.ranges` says.
 */
import type { ErrorCode } from './contract.js'
import { Arr } from './arrays/value.js'
import { MAX_SERIAL } from './dates.js'
import { arityOf, type Positions } from './arity.js'
import type { Counts } from './locale/messages.js'
import type { Rules } from './locale/rules.js'
import { Areas, E, fitNumber, isErr, Rng, whole, type CErr, type V } from './values.js'
import { CalcError } from './exec/budget.js'
import type { Line } from './functions/lookup/search.js'

/** A range (or a single value) as a grid. */
export interface Grid {
  /** Rows and columns as referenced: a whole column has 1,048,576 rows. */
  rows: number
  cols: number
  /** The part that holds cells, from the top-left corner: h × w values row by row; the rest is blank. */
  h: number
  w: number
  v: readonly V[]
}

/** What the runtime passes a function: a value, a reference, an array, or a reference of several areas. */
export type Arg = V | Rng | Arr | Areas

/** A value from a `V` argument: a direct value, or the values of a range (row by row) as one segment. */
export type Segment = V | readonly V[]

export interface Bounds {
  r1: number
  c1: number
  r2: number
  c2: number
  h: number
  w: number
}

export interface Selection {
  sel: readonly number[]
  rest: number
  read(g: Rng): readonly V[]
}

/** Which hidden rows SUBTOTAL and AGGREGATE skip: all of them, those of a filtered sheet, none. */
export type HiddenRule = 'all' | 'filtered' | 'none'

/** What a function declared with `cx` gets as its first argument. */
export interface Cx {
  /** The clock moment, a serial number with a day fraction (UTC). */
  readonly now: number
  /** The rules of computation of the book's locale: text as numbers and dates, numbers as text, formats. */
  readonly rules: Rules
  /** Pseudo-random [0, 1) from the clock seed, a cell and an element of an array (RANDARRAY). */
  rand(s: number, r: number, c: number, k?: number): number
  /** Range values row by row (the occupied part); shared with caches, never to be modified. */
  flat(g: Rng): readonly V[]
  /**
   * Range values like `flat`, blank where SUBTOTAL and AGGREGATE skip: with `nested` the cells whose formula calls
   * either, with `hidden` the hidden rows (`all`), or those of a filtered sheet only (`filtered`).
   */
  counted(g: Rng, nested: boolean, hidden: HiddenRule): readonly V[]
  bounds(g: Rng): Bounds
  /** A value in a scalar position: the single cell of a range or array, or #VALUE!. */
  scalar(x: Arg): V
  grid(x: Arg): Grid
  /** A reference's first column or row, read without materializing its result columns. */
  lookup?(g: Rng, horizontal: boolean): Line
  /** A reference as an array of its referenced size; values and arrays as they are. */
  view(x: Arg): V | Arr
  /**
   * Criteria over ranges of one referenced size (else #VALUE!), on different sheets too: the positions
   * (row by row in the referenced width) where every (range, criterion) pair holds within the part any of
   * the ranges occupies, `rest` the matching positions past it (all blank: they count when every
   * criterion matches a blank), and `read` a range's values by those positions. `sum` is a range of
   * the same size read along (SUMIFS).
   */
  select(pairs: [V | Rng, V | Rng][], sum?: Rng): Selection | CErr
}

export interface FnDef {
  /** Argument kinds separated by spaces; `?` after a kind marks an optional argument (`'n i?'`). */
  args: string
  /** How many of the last kinds repeat (SUM: 1, SUMIFS: 2); the repeated arguments arrive as one array, lists spliced in. */
  rep?: number
  /** Excel's own limit on repeated arguments, below its 255 (SWITCH: 126 pairs and a default). */
  max?: number
  /** The implementation, given the coerced arguments (an omitted optional argument is undefined). */
  fn: (...a: any[]) => Arg
  /** The error when all list arguments (`N`, `A`, `Y`, `Z`) are empty (rule 3). */
  empty?: ErrorCode
  /** `fn` gets the function context (`Cx`) first. */
  cx?: boolean
  /**
   * The answer is a date, a serial number of the 1900 system the calendar counts in, moved to the book's date system:
   * `book` a date of that system (#NUM! before its first day), `calendar` any date the function reached.
   */
  date?: 'book' | 'calendar'
  /** `fn` gets the formula's cell (sheet, row, column) after the context: RAND. */
  cell?: boolean
  /** `fn` gets every argument as a function evaluated on demand, without coercion: IF, IFERROR. */
  lazy?: boolean
  volatile?: boolean
  /** The answer can be an array (`Arr`) or a reference of several cells. */
  array?: boolean
  /**
   * A range of several cells in a one-value argument is lifted like an array, except: `'error'` gives
   * #VALUE! (Excel 365's Analysis ToolPak functions, `EOMONTH(A1:A2,0)`), `'first'` takes its top-left cell
   * (N, T), a list names the arguments that give #VALUE! (WORKDAY.INTL). An array is lifted all the same.
   */
  ranges?: 'error' | 'first' | readonly number[]
  /** Lazy functions: the arguments whose array lifts the call (by default the first, IF's condition). */
  liftAt?: (i: number, n: number) => boolean
  /** Arguments taken whole that a lifted call reads by their top-left value, as a 1 × 1 array (VLOOKUP's index). */
  liftFirst?: readonly number[]
  /** An argument left out (`f(a,,c)`) where it is not a blank read by kind (`functions/omitted.ts`). */
  omit?: Omission
  /** Where Excel does not load a formula that leaves the argument out. */
  notOmitted?: Positions
  /**
   * A reference of several areas (a union, one across sheets): `true` its `R` arguments take it (INDEX's area number);
   * else the errors the function answers for one in an argument that takes a range (`areas.ts` by kind, measured).
   */
  areas?: true | (AreaErrors & { at?: Readonly<Record<number, AreaErrors>> })
  /**
   * Its `R` arguments are ranges of criteria (SUMIF): an array there, which only a name or LET can put, lifts the call
   * over its values, and each is #VALUE!, as in Excel.
   */
  criteria?: boolean
}

/** A left-out argument: not given (the function's default), a blank, TRUE, or an error. */
export type Omit = 'default' | 'blank' | 'true' | 'NA' | 'VALUE'

export interface Omission {
  /** By position. */
  at?: Readonly<Record<number, Omit>>
  /** Every repeated argument past the declared kinds, or one rule for each kind of the repeated group. */
  rest?: Omit | readonly Omit[]
  /** The former Analysis ToolPak's rule elsewhere: not given, a required one #N/A. */
  analysis?: boolean
}

const TAKEN: Record<Omit, V | undefined> = { default: undefined, blank: null, true: true, NA: E.NA, VALUE: E.VALUE }

/** What the parser and the code generator need to know about a function. */
/** How an argument takes an array: element by element, arrays only (kind `r`), passed unless lifted (`p`), whole. */
export type ArgMode = 'lift' | 'array' | 'pass' | 'whole'

export interface FunctionSpec {
  min: number
  max: number
  /** Whether a reference in this argument is passed as a reference (a single cell too). */
  range: (i: number) => boolean
  mode: (i: number) => ArgMode
  /** A range of several cells in a lifted argument: lifted, #VALUE!, or its top-left cell. */
  rangeAt: (i: number) => 'lift' | 'error' | 'first'
  array?: boolean
  /** Lazy functions: the arguments whose array lifts the call. */
  liftAt: (i: number, n: number) => boolean
  /** An argument of kind `P`: the function can answer a reference it was given. */
  refs: boolean
  /** Arguments taken whole that a lifted call reads by their top-left value. */
  liftFirst: readonly number[]
  /** Whether a call with `k` arguments is one Excel loads: within min…max, repeated pairs complete (SUMIFS 3, 5, …; SORTBY 2, 3, 5, …). */
  takes: (k: number) => boolean
  /** What the function gets for an argument left out at `i`: a value (a blank is null), undefined when it is not given. */
  omitted: (i: number) => V | undefined
  /** Whether Excel loads a formula that leaves out the argument at `i`. */
  omits: (i: number) => boolean
  /** The counts it takes, for a message (SUMIFS 3 to 255, odd). */
  counts: Counts
  volatile?: boolean
  lazy?: boolean
  cell?: boolean
  /** A range of criteria (`FnDef.criteria`): an array there makes the call an array of #VALUE!. */
  criteriaAt: (i: number) => boolean
}

const RANGE_KINDS = 'NAYZVGDCMRrP'
const LIST_KINDS = 'NAYZ'
const WHOLE_KINDS = 'NAYZVGDCMR'
/** Kinds of an array or a reference: an argument left out there is #VALUE!. */
const ARRAY_KINDS = 'GDCMR'

interface Kind {
  k: string
  optional: boolean
}

function parseKinds(args: string): Kind[] {
  return args
    .split(' ')
    .filter(Boolean)
    .map((t) => ({ k: t[0], optional: t.endsWith('?') }))
}

/** The kind of the argument at `i`: past the declared kinds the last `rep` of them repeat. */
const kindOf = (kinds: readonly string[], rep: number) => {
  const fixed = kinds.length - rep
  return (i: number) => (i < kinds.length ? kinds[i] : kinds[fixed + ((i - fixed) % rep)])
}

export function specOf(name: string, def: FnDef): FunctionSpec {
  const kinds = parseKinds(def.args)
  for (const k of kinds) if (!(k.k in COERCE)) throw new Error(`${name}: unknown argument kind ${k.k}`)
  const rep = def.rep ?? 0
  const fixed = kinds.length - rep
  const kindAt = kindOf(kinds.map((k) => k.k), rep)
  const { min, max, takes, counts, omits } = arityOf(def)
  const mode = (i: number): ArgMode => {
    const k = kindAt(i)
    return k === 'r' ? 'array' : k === 'p' || k === 'P' ? 'pass' : WHOLE_KINDS.includes(k) ? 'whole' : 'lift'
  }
  const { ranges, omit } = def
  const rangeAt = (i: number) => (ranges === undefined ? 'lift' : ranges === 'first' || ranges === 'error' ? ranges : ranges.includes(i) ? 'error' : 'lift')
  const omitted = (i: number): V | undefined => {
    const rest = rep && i >= fixed ? omit?.rest : undefined
    const own = omit?.at?.[i] ?? (typeof rest === 'object' ? rest[(i - fixed) % rep] : rest)
    return TAKEN[own ?? (omit?.analysis ? (i < min ? 'NA' : 'default') : ARRAY_KINDS.includes(kindAt(i)) ? 'VALUE' : 'blank')]
  }
  return {
    min,
    max,
    takes,
    counts,
    range: (i) => RANGE_KINDS.includes(kindAt(i)),
    mode,
    rangeAt,
    volatile: def.volatile,
    lazy: def.lazy,
    cell: def.cell,
    criteriaAt: (i) => def.criteria === true && kindAt(i) === 'R',
    array: def.array,
    liftAt: def.liftAt ?? ((i) => i === 0),
    refs: kinds.some((k) => k.k === 'P'),
    liftFirst: def.liftFirst ?? [],
    omitted,
    omits,
  }
}

// ---------- coercion by kind ----------

/** A date argument: a whole serial number, from a number or date text; outside 0…9999-12-31 #NUM!. */
export function toDate(v: V, rules: Rules): number | CErr {
  const n = toTime(v, rules)
  return isErr(n) ? n : Math.floor(n)
}

/**
 * A time or date-time argument: a serial number with a day fraction, from a number or text, in the book's date
 * system (outside its first day … 9999-12-31 #NUM!), given to the function in the 1900 system its calendar counts in.
 */
function toTime(v: V, rules: Rules): number | CErr {
  const n = timeOfDay(v, rules)
  return isErr(n) ? n : n + rules.epoch
}

/** A time of day argument (HOUR, MINUTE, SECOND): as `toTime`, in the book's date system, whose day fraction is the same. */
function timeOfDay(v: V, rules: Rules): number | CErr {
  const n = timeRead(v, rules)
  if (isErr(n)) return n
  return inCalendar(n, rules) ? n : E.NUM
}

/** A time or date-time argument as read, before its range is checked (WEEKDAY reads its type in between). */
export const timeRead = (v: V, rules: Rules): number | CErr => (typeof v === 'string' ? rules.dateText(v) : null) ?? rules.toNum(v)

/** Whether a serial number of the book's date system lies in its calendar, its first day … 9999-12-31. */
export const inCalendar = (n: number, rules: Rules) => n >= 0 && n < MAX_SERIAL + 1 - rules.epoch

/** A number where a boolean is #VALUE! (the former Analysis ToolPak's rule). */
export const noBool = (v: V, rules: Rules): number | CErr => (typeof v === 'boolean' ? E.VALUE : rules.toNum(v))

type ListOther = (v: Exclude<V, number | null>) => number | CErr | undefined

/**
 * A list kind: a range gives its numbers and its other non-blank cells by `inRange` (a number, an error, or
 * undefined to skip); a direct argument goes through `direct`.
 */
function list(direct: (v: V, rules: Rules) => number | CErr, inRange: ListOther, inArray = inRange) {
  return (x: Arg, cx: Cx): number[] | CErr => {
    const out: number[] = []
    const { rules } = cx
    const other = x instanceof Arr ? inArray : inRange
    // Keep the output and conversion rule local so the JIT can specialize this shared loop.
    const numbers = (values: readonly V[]): CErr | undefined => {
      for (const v of values) {
        if (v === null) continue
        const n = typeof v === 'number' ? v : other(v)
        if (n === undefined) continue
        if (isErr(n)) return n
        out.push(n)
      }
    }
    if (x instanceof Arr) {
      // in order: the occupied rows, then the fill past them (a whole column: 0 for B:B*C:C)
      const a = x.w < x.cols && x.fill !== null ? x.full() : x
      const e = numbers(a.v)
      if (e) return e
      const rest = a.rows * a.cols - a.h * a.w
      if (rest <= 0 || a.fill === null) return out
      // the fill once, then repeated: a whole column adds a million numbers
      const before = out.length
      const error = numbers([a.fill])
      if (error) return error
      return out.length > before ? out.concat(new Array<number>(rest - 1).fill(out[before])) : out
    }
    if (x instanceof Rng) return numbers(cx.flat(x)) ?? out
    // each area of several
    if (x instanceof Areas) {
      for (const g of x.list) {
        const e = numbers(cx.flat(g))
        if (e) return e
      }
      return out
    }
    const n = direct(x, rules)
    if (isErr(n)) return n
    out.push(n)
    return out
  }
}

const errOrSkip = (v: V) => (isErr(v) ? v : undefined)
const countOther = (v: V) => isErr(v) ? v : typeof v === 'boolean' ? Number(v) : 0

/**
 * A reference to one cell as Excel reads it before a function compares its arrays: the cell's error, #VALUE! for a
 * blank cell; undefined for anything else, an array of one too (the sets `sumx-cell-held`, `sumproduct-held`).
 */
export function oneCell(x: Arg, cx: Cx): CErr | undefined {
  if (!(x instanceof Rng) || x.r1 !== x.r2 || x.c1 !== x.c2) return undefined
  const v = cx.scalar(x)
  return isErr(v) ? v : v === null ? E.VALUE : undefined
}

/** Coercion of an argument by kind; a CErr result stops the call, except for the kinds in KEEP_ERRORS. */
const num = (v: V, rules: Rules) => rules.toNum(v)

const COERCE: Record<string, (x: Arg, cx: Cx) => unknown> = {
  n: (x, cx) => cx.rules.toNum(cx.scalar(x)),
  i: (x, cx) => {
    const n = cx.rules.toNum(cx.scalar(x))
    return isErr(n) ? n : Math.trunc(n)
  },
  w: (x, cx) => {
    const n = cx.rules.toNum(cx.scalar(x))
    return isErr(n) ? n : whole(n)
  },
  x: (x, cx) => noBool(cx.scalar(x), cx.rules),
  d: (x, cx) => toDate(cx.scalar(x), cx.rules),
  t: (x, cx) => toTime(cx.scalar(x), cx.rules),
  h: (x, cx) => timeOfDay(cx.scalar(x), cx.rules),
  s: (x, cx) => cx.rules.toText(cx.scalar(x)),
  b: (x, cx) => cx.rules.toBool(cx.scalar(x)),
  v: (x, cx) => cx.scalar(x),
  a: (x, cx) => cx.scalar(x),
  N: list(num, errOrSkip),
  // in an array MAXA, STDEVA and the like take the numbers only; AVERAGEA (`Y`) counts text there as 0 (Excel 365)
  A: list(num, countOther, errOrSkip),
  Y: list(num, countOther, (v) => (isErr(v) ? v : typeof v === 'boolean' ? undefined : 0)),
  Z: list(noBool, (v) => (isErr(v) ? v : E.VALUE)),
  // an empty element of an array counts for COUNTA (as Excel's empty value does): it is read as empty text
  V: (x, cx) => (x instanceof Rng ? cx.flat(x) : x instanceof Areas ? x.list.flatMap((g) => cx.flat(g)) : x instanceof Arr ? x.values().map((v) => (v === null ? '' : v)) : x),
  // several areas where one rectangle is wanted: #REF! of a union (ROWS), else #VALUE!
  G: (x, cx) => (isErr(x) ? x : cx.grid(x)),
  D: (x, cx) => (isErr(x) || x instanceof Rng ? x : cx.grid(x)),
  C: (x, cx) => (isErr(x) ? x : (oneCell(x, cx) ?? cx.grid(x))),
  M: (x, cx) => {
    if (isErr(x)) return x
    const g = cx.grid(x)
    return gridError(g) ?? g
  },
  R: (x) => x,
  r: (x) => x,
  p: (x) => x,
  P: (x) => x,
}


/**
 * The errors a function answers for a union and for a reference across sheets where it does not take them; `inside`:
 * the error counts as one inside a range of values, after an error given as a value (CONCAT).
 */
export interface AreaErrors {
  union?: ErrorCode
  cube?: ErrorCode
  inside?: true
}

/**
 * A reference of several areas as an argument of kind `k`: the lists take each area (SUM, COUNT), INDEX and the
 * functions that say so take it whole; elsewhere it is an error, by default #VALUE! for a union, #REF! for one
 * across sheets in an array argument (`G`, `C`, `M`) and #VALUE! in the others.
 */
function areasArg(own: FnDef['areas'], p: number, k: string, x: Areas, cx: Cx): unknown {
  if (own === true && k === 'R') return x
  const errors = typeof own === 'object' ? (own.at?.[p] ?? own) : undefined
  const error = errors && (x.cube ? errors.cube : errors.union)
  if (error) return k === 'V' && errors.inside ? [E[error]] : E[error]
  if (LIST_KINDS.includes(k) || k === 'V') return COERCE[k](x, cx)
  if (k === 'p' || k === 'P') return x
  return x.cube && 'GDCM'.includes(k) ? E.REF : E.VALUE
}

/** An argument coerced by kind inside a function that reads it after others (FVSCHEDULE's schedule after the principal). */
export const coerce = (kind: string, x: Arg, cx: Cx): unknown => COERCE[kind](x, cx)

/** Kinds whose argument may be an error without stopping the call. */
const KEEP_ERRORS = 'aRVrpP'
/**
 * Excel reads the arguments given as values before those given as references or arrays (`LARGE(A1:A9,"x")` and
 * `SUM(A1:A9,"x")` are #VALUE! even with an error in A1:A9).
 */
const late = (x: unknown) => x instanceof Rng || x instanceof Arr || x instanceof Areas

/** The callable the generated code uses for a function: coercion by kind, the call, the answer rules. */
export function makeCaller(def: FnDef, cx: Cx): (...a: any[]) => Arg {
  const { fn } = def
  if (def.lazy) return def.cx ? (...a) => fn(cx, ...a) : fn
  const kinds = parseKinds(def.args).map((k) => k.k)
  const rep = def.rep ?? 0
  const fixed = kinds.length - rep
  const head = (def.cx ? 1 : 0) + (def.cell ? 3 : 0)
  const kindAt = kindOf(kinds, rep)
  return (...a: unknown[]) => {
    const out: unknown[] = []
    if (def.cx) out.push(cx)
    const cells = def.cell ? 3 : 0
    for (let i = 0; i < cells; i++) out.push(a[i])
    const vals: unknown[] = new Array(a.length - cells)
    // a reference across sheets the function does not take is found before any other argument (SORT's #REF!)
    for (let p = 0; p < vals.length; p++) {
      const x = a[cells + p]
      if (!(x instanceof Areas) || !x.cube) continue
      const k = kindAt(p)
      const v = areasArg(def.areas, p, k, x, cx)
      if (isErr(v) && !KEEP_ERRORS.includes(k)) return v
    }
    for (const pass of [false, true])
      for (let p = 0; p < vals.length; p++) {
        const k = kindAt(p)
        const x = a[cells + p] as Arg | undefined
        // an argument left out that the function takes as not given (`omitted`) stays undefined
        if (x === undefined || late(x) !== pass) continue
        const v = x instanceof Areas ? areasArg(def.areas, p, k, x, cx) : COERCE[k](x, cx)
        if (isErr(v) && !KEEP_ERRORS.includes(k)) return v
        vals[p] = v
      }
    const tail: unknown[] = []
    let lists = 0
    let nums = 0
    for (let p = 0; p < vals.length; p++) {
      const k = kindAt(p)
      const v = vals[p]
      // a repeated argument not given is left out (PRODUCT(2,), LCM(6,,4)); in pairs it keeps its place
      if (v === undefined && rep === 1 && p >= fixed) continue
      const list = LIST_KINDS.includes(k) && v !== undefined
      if (list) {
        lists++
        nums += (v as number[]).length
      }
      if (rep && p >= fixed) {
        if (list) for (const e of v as number[]) tail.push(e)
        else tail.push(v)
      } else out.push(v)
    }
    if (def.empty && lists && !nums) return E[def.empty]
    if (def.criteria) {
      const arrays = vals.filter((v, p) => v instanceof Arr && kindAt(p) === 'R') as Arr[]
      if (arrays.length) return valueErrors(Math.max(...arrays.map((a) => a.rows)), Math.max(...arrays.map((a) => a.cols)))
    }
    if (rep) {
      while (out.length < head + fixed) out.push(undefined)
      out.push(tail)
    }
    let r: Arg
    try {
      r = fn(...out)
    } catch (e) {
      // a stopped computation goes on stopping; a defect in one function must not abort the whole book
      if (e instanceof CalcError) throw e
      return E.VALUE
    }
    if (typeof r !== 'number') return r
    if (def.date && cx.rules.epoch) {
      r -= cx.rules.epoch
      if (r < 0 && def.date === 'book') return E.NUM
    }
    return fitNumber(r)
  }
}

/** An array of #VALUE! only. */
const valueErrors = (rows: number, cols: number) => new Arr(rows, cols, rows, cols, new Array<V>(rows * cols).fill(E.VALUE))

// ---------- helpers for function implementations ----------

/** The value at a position of the referenced area (row by row); blank outside the occupied part. */
export function gridAt(g: Grid, i: number): V {
  const r = Math.floor(i / g.cols)
  const c = i - r * g.cols
  return r < g.h && c < g.w ? g.v[r * g.w + c] : null
}

/** Positions of a grid that hold cells, row by row in the referenced width; past them everything is blank. */
export const occupied = (g: Grid) => (g.h === 0 ? 0 : (g.h - 1) * g.cols + g.w)

/** The first error in the occupied part of a grid. */
export const gridError = (g: Grid): CErr | undefined => g.v.find(isErr) as CErr | undefined

/** The numbers of a grid, row by row; an error among them is it (the `M` kind answers it first), anything else #VALUE!. */
export function values(g: Grid): number[] | CErr {
  const out: number[] = []
  for (let i = 0; i < g.rows * g.cols; i++) {
    const v = gridAt(g, i)
    if (isErr(v)) return v
    if (typeof v !== 'number') return E.VALUE
    out.push(v)
  }
  return out
}

/** A flat list of n·k numbers as n observations of k variables: row by row when `vertical`, else column by column. */
export const observations = (xs: readonly number[], n: number, k: number, vertical: boolean): number[][] =>
  Array.from({ length: n }, (_, i) => Array.from({ length: k }, (_, j) => (vertical ? xs[i * k + j] : xs[j * n + i])))

/** Rule 4: number pairs from two arrays; pairs with a non-number are dropped; different sizes are #N/A. */
export function pairs(a: Grid, b: Grid): [number[], number[]] | CErr {
  const n = a.rows * a.cols
  if (n !== b.rows * b.cols) return E.NA
  // past the occupied parts every pair is blank
  const end = Math.min(n, Math.max(occupied(a), occupied(b)))
  const xs: number[] = []
  const ys: number[] = []
  for (let i = 0; i < end; i++) {
    const x = gridAt(a, i)
    const y = gridAt(b, i)
    if (typeof x === 'number' && typeof y === 'number') {
      xs.push(x)
      ys.push(y)
    }
  }
  return [xs, ys]
}

/** The first error given as a value among `V` arguments: Excel finds it before any error inside a range or array (`AND({TRUE,#DIV/0!},NA())` is #N/A). */
export const directError = (segs: readonly Segment[]): CErr | undefined => segs.find((s): s is CErr => !Array.isArray(s) && isErr(s))

/** Visits the values of `V` arguments in order, telling whether each came from a range; an error returned by `visit` stops and is returned. */
export function eachValue(segs: readonly Segment[], visit: (v: V, inRange: boolean) => CErr | void): CErr | undefined {
  for (const s of segs) {
    if (Array.isArray(s)) {
      for (const v of s as readonly V[]) {
        const e = visit(v, true)
        if (e) return e
      }
    } else {
      const e = visit(s as V, false)
      if (e) return e
    }
  }
}
