/**
 * The shared argument layer. Every function declares the kinds of its arguments (`FnDef.args`), and the
 * runtime coerces the arguments by kind before the call: a function never repeats the coercion, it names
 * the group it belongs to. An error in a coerced argument is the answer, without calling the function.
 *
 * The kinds follow the rules derived from the coverage measurement against Excel en-US
 * (`docs/coverage-formulajs.md`, "Предложение"):
 *
 * 1. Coercion by group, not one rule for all:
 *    - `n` number: a boolean is 1 or 0, a blank 0, text that reads as a number is that number, other text
 *      #VALUE!; `i` the same, truncated to an integer; `x` the same, but a boolean is #VALUE! (functions
 *      of the former Analysis ToolPak: engineering and part of the financial ones);
 *    - `s` text (a number in the General format, TRUE/FALSE); `b` boolean (text "TRUE"/"FALSE" too, as in
 *      Excel en-US with the English UI); `v` any value (an error stops the call); `a` any value, errors
 *      included (ISERROR, IFS);
 *    - lists of numbers from ranges and arguments: `N` the SUM rule (a range gives its numbers, a direct
 *      argument is coerced like `n`), `A` the MAXA rule (a range gives numbers, booleans as 1/0 and
 *      text as 0; a direct argument is coerced like `n`), `Y` the same for AVERAGEA, `Z` strict (text or a boolean in a range and a
 *      boolean argument are #VALUE!: GCD, LCM, MULTINOMIAL);
 *    - `V` values in order, errors included, each range as one segment, so a function can tell a range
 *      from a direct argument (COUNTA counts every direct argument); `G` a range as a grid with its
 *      referenced size, errors inside kept; `M` a numeric grid (an error inside is the answer); `R` the
 *      argument itself, a reference or a value (criteria functions, which work on the reference through
 *      the function context); `r` one value that may be a reference (the function gets a reference to one
 *      or more cells as it is; an array is lifted); `p` any value, reference or array, passed as it is
 *      unless the call is lifted (CHOOSE, the values of IF and IFS).
 * 2. Date and time text in a date argument (`d`, `t`, `D`) is read as Excel en-US reads it
 *    (`2025-01-31`, `1/31/2025`, `18:30`; see `dates.ts`); a date before 0 or after 9999-12-31 is #NUM!.
 * 3. Empty input: `FnDef.empty` names the error of a function whose lists (`N`, `A`, `Z`, `D`) hold no
 *    numbers at all (AVERAGE #DIV/0!, MEDIAN #NUM!, MODE #N/A); a blank cell as a whole answer is null,
 *    which a cell shows as 0.
 * 4. Paired arrays (`pairs`): pairs where either value is not a number are dropped; arrays with different
 *    numbers of values are #N/A.
 * 5. The answer: a non-finite number is #NUM!, one below the smallest normal number is 0 (`fitNumber`);
 *    radix digits are upper-case and complex text is written as Excel writes it by the functions themselves
 *    (`functions/engineering/radix.ts`, `complex.ts`).
 * 6. Holidays and payment dates (`D`): a flat list of whole serial numbers.
 * 7. Arrays (`docs/arrays-incell.md`): the list and grid kinds take an array as they take a range (`A`
 *    takes only an array's numbers, `Y` counts its text as 0); a one-value kind given an array or a range
 *    of several cells is lifted element by element (`src/arrays/call.ts`), except where `FnDef.ranges` says.
 */
import type { ErrorCode } from './contract.js'
import { Arr } from './arrays/value.js'
import { MAX_SERIAL, parseDateText } from './dates.js'
import { E, fitNumber, isErr, Rng, toBool, toNum, toText, type CErr, type V } from './values.js'

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

/** What the runtime passes a function: a value, a reference, or an array. */
export type Arg = V | Rng | Arr

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
  sel: number[]
  rest: number
  read(g: Rng): readonly V[]
}

/** What a function declared with `cx` gets as its first argument. */
export interface Cx {
  /** The clock moment, a serial number with a day fraction (UTC). */
  readonly now: number
  /** Pseudo-random [0, 1) from the clock seed, a cell and an element of an array (RANDARRAY). */
  rand(s: number, r: number, c: number, k?: number): number
  /** Range values row by row (the occupied part); shared with caches, never to be modified. */
  flat(g: Rng): readonly V[]
  /** The same with the cells whose formula calls SUBTOTAL or AGGREGATE blank: SUBTOTAL and AGGREGATE skip them. */
  flatOuter(g: Rng): readonly V[]
  bounds(g: Rng): Bounds
  /** A value in a scalar position: the single cell of a range or array, or #VALUE!. */
  scalar(x: Arg): V
  grid(x: Arg): Grid
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
  /** The implementation, given the coerced arguments (an omitted optional argument is undefined). */
  fn: (...a: any[]) => Arg
  /** The error when all list arguments (`N`, `A`, `Z`, `D`) are empty (rule 3). */
  empty?: ErrorCode
  /** `fn` gets the function context (`Cx`) first. */
  cx?: boolean
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
}

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
  /** Parity of the argument count (odd for SUMIFS, even for COUNTIFS). */
  parity?: 'odd' | 'even'
  volatile?: boolean
  lazy?: boolean
  cell?: boolean
}

/** Excel's limit on the number of arguments. */
const MAX_ARGS = 255
const RANGE_KINDS = 'NAYZVGMRDr'
const LIST_KINDS = 'NAYZD'
const WHOLE_KINDS = 'NAYZVGMRD'

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

export function specOf(name: string, def: FnDef): FunctionSpec {
  const kinds = parseKinds(def.args)
  for (const k of kinds) if (!(k.k in COERCE)) throw new Error(`${name}: unknown argument kind ${k.k}`)
  const rep = def.rep ?? 0
  const fixed = kinds.length - rep
  const kindAt = (i: number) => (i < kinds.length ? kinds[i].k : kinds[fixed + ((i - fixed) % rep)].k)
  const min = kinds.filter((k) => !k.optional).length
  let max = rep ? MAX_ARGS : kinds.length
  let parity: FunctionSpec['parity']
  if (rep === 2) {
    parity = fixed % 2 === 0 ? 'even' : 'odd'
    if ((max - fixed) % 2) max--
  }
  const mode = (i: number): ArgMode => {
    const k = kindAt(i)
    return k === 'r' ? 'array' : k === 'p' ? 'pass' : WHOLE_KINDS.includes(k) ? 'whole' : 'lift'
  }
  const { ranges } = def
  const rangeAt = (i: number) => (ranges === undefined ? 'lift' : ranges === 'first' || ranges === 'error' ? ranges : ranges.includes(i) ? 'error' : 'lift')
  return { min, max, range: (i) => RANGE_KINDS.includes(kindAt(i)), mode, rangeAt, parity, volatile: def.volatile, lazy: def.lazy, cell: def.cell, array: def.array, liftAt: def.liftAt ?? ((i) => i === 0) }
}

// ---------- coercion by kind ----------

/** A date argument: a whole serial number, from a number or date text; outside 0…9999-12-31 #NUM!. */
export function toDate(v: V): number | CErr {
  const n = toTime(v)
  return isErr(n) ? n : Math.floor(n)
}

/** A time or date-time argument: a serial number with a day fraction, from a number or text. */
export function toTime(v: V): number | CErr {
  const t = typeof v === 'string' ? parseDateText(v) : null
  const n = t ?? toNum(v)
  if (isErr(n)) return n
  return n < 0 || n >= MAX_SERIAL + 1 ? E.NUM : n
}

const noBool = (v: V): number | CErr => (typeof v === 'boolean' ? E.VALUE : toNum(v))

/**
 * A list kind: a range gives its numbers (through `num`, if given) and its other non-blank cells by
 * `inRange` (a number, an error, or undefined to skip); a direct argument goes through `direct`.
 */
function list(direct: (v: V) => number | CErr, inRange: (v: Exclude<V, number | null>) => number | CErr | undefined, num?: (n: number) => number | CErr, inArray = inRange) {
  return (x: Arg, cx: Cx): number[] | CErr => {
    const out: number[] = []
    if (x instanceof Arr) {
      const take = (v: V): CErr | undefined => {
        if (v === null) return
        const n = typeof v === 'number' ? (num ? num(v) : v) : inArray(v)
        if (n === undefined) return
        if (isErr(n)) return n
        out.push(n)
      }
      // in order: the occupied rows, then the fill past them (a whole column: 0 for B:B*C:C)
      const a = x.w < x.cols && x.fill !== null ? x.full() : x
      for (const v of a.v) {
        const e = take(v)
        if (e) return e
      }
      const rest = a.rows * a.cols - a.h * a.w
      if (rest <= 0 || a.fill === null) return out
      // the fill once, then repeated: a whole column adds a million numbers
      const before = out.length
      const e = take(a.fill)
      if (e) return e
      return out.length > before ? out.concat(new Array<number>(rest - 1).fill(out[before])) : out
    }
    if (!(x instanceof Rng)) {
      const n = direct(x)
      if (isErr(n)) return n
      out.push(n)
      return out
    }
    for (const v of cx.flat(x)) {
      if (v === null) continue
      const n = typeof v === 'number' ? (num ? num(v) : v) : inRange(v)
      if (n === undefined) continue
      if (isErr(n)) return n
      out.push(n)
    }
    return out
  }
}

const errOrSkip = (v: V) => (isErr(v) ? v : undefined)

/** Coercion of an argument by kind; a CErr result stops the call, except for the kinds in KEEP_ERRORS. */
const COERCE: Record<string, (x: Arg, cx: Cx) => unknown> = {
  n: (x, cx) => toNum(cx.scalar(x)),
  i: (x, cx) => {
    const n = toNum(cx.scalar(x))
    return isErr(n) ? n : Math.trunc(n)
  },
  x: (x, cx) => noBool(cx.scalar(x)),
  d: (x, cx) => toDate(cx.scalar(x)),
  t: (x, cx) => toTime(cx.scalar(x)),
  s: (x, cx) => toText(cx.scalar(x)),
  b: (x, cx) => toBool(cx.scalar(x)),
  v: (x, cx) => cx.scalar(x),
  a: (x, cx) => cx.scalar(x),
  N: list(toNum, errOrSkip),
  // in an array MAXA, STDEVA and the like take the numbers only; AVERAGEA (`Y`) counts text there as 0 (Excel 365)
  A: list(toNum, (v) => (isErr(v) ? v : typeof v === 'boolean' ? Number(v) : 0), undefined, errOrSkip),
  Y: list(toNum, (v) => (isErr(v) ? v : typeof v === 'boolean' ? Number(v) : 0), undefined, (v) => (isErr(v) ? v : typeof v === 'boolean' ? undefined : 0)),
  Z: list(noBool, (v) => (isErr(v) ? v : E.VALUE)),
  D: list(toDate, (v) => (isErr(v) ? v : E.VALUE), toDate),
  // an empty element of an array counts for COUNTA (as Excel's empty value does): it is read as empty text
  V: (x, cx) => (x instanceof Rng ? cx.flat(x) : x instanceof Arr ? x.values().map((v) => (v === null ? '' : v)) : x),
  G: (x, cx) => (isErr(x) ? x : cx.grid(x)),
  M: (x, cx) => {
    if (isErr(x)) return x
    const g = cx.grid(x)
    return g.v.find(isErr) ?? g
  },
  R: (x) => x,
  r: (x) => x,
  p: (x) => x,
}

/** An argument coerced by kind inside a function that reads it after others (FVSCHEDULE's schedule after the principal). */
export const coerce = (kind: string, x: Arg, cx: Cx): unknown => COERCE[kind](x, cx)

/** Kinds whose argument may be an error without stopping the call. */
const KEEP_ERRORS = 'aRVrp'
/**
 * Excel reads the arguments given as values before those given as references or arrays (`LARGE(A1:A9,"x")` and
 * `SUM(A1:A9,"x")` are #VALUE! even with an error in A1:A9).
 */
const late = (x: unknown) => x instanceof Rng || x instanceof Arr

/** The callable the generated code uses for a function: coercion by kind, the call, the answer rules. */
export function makeCaller(def: FnDef, cx: Cx): (...a: any[]) => V | Arr | Rng {
  const { fn } = def
  if (def.lazy) return def.cx ? (...a) => fn(cx, ...a) : fn
  const kinds = parseKinds(def.args).map((k) => k.k)
  const rep = def.rep ?? 0
  const fixed = kinds.length - rep
  const head = (def.cx ? 1 : 0) + (def.cell ? 3 : 0)
  const kindAt = (p: number) => (p < kinds.length ? kinds[p] : kinds[fixed + ((p - fixed) % rep)])
  return (...a: unknown[]) => {
    const out: unknown[] = []
    if (def.cx) out.push(cx)
    const cells = def.cell ? 3 : 0
    for (let i = 0; i < cells; i++) out.push(a[i])
    const vals: unknown[] = new Array(a.length - cells)
    for (const pass of [false, true])
      for (let p = 0; p < vals.length; p++) {
        const k = kindAt(p)
        const x = a[cells + p] as Arg
        if (late(x) !== pass) continue
        const v = COERCE[k](x, cx)
        if (isErr(v) && !KEEP_ERRORS.includes(k)) return v
        vals[p] = v
      }
    const tail: unknown[] = []
    let lists = 0
    let nums = 0
    for (let p = 0; p < vals.length; p++) {
      const k = kindAt(p)
      const v = vals[p]
      if (LIST_KINDS.includes(k)) {
        lists++
        nums += (v as number[]).length
      }
      if (rep && p >= fixed) {
        if (LIST_KINDS.includes(k)) for (const e of v as number[]) tail.push(e)
        else tail.push(v)
      } else out.push(v)
    }
    if (def.empty && lists && !nums) return E[def.empty]
    if (rep) {
      while (out.length < head + fixed) out.push(undefined)
      out.push(tail)
    }
    let r: V | Arr | Rng
    try {
      r = fn(...out)
    } catch {
      // a defect in one function must not abort the whole book
      return E.VALUE
    }
    return typeof r === 'number' ? fitNumber(r) : r
  }
}

// ---------- helpers for function implementations ----------

/** The value at a position of the referenced area (row by row); blank outside the occupied part. */
export function gridAt(g: Grid, i: number): V {
  const r = Math.floor(i / g.cols)
  const c = i - r * g.cols
  return r < g.h && c < g.w ? g.v[r * g.w + c] : null
}

/** Rule 4: number pairs from two arrays; pairs with a non-number are dropped; different sizes are #N/A. */
export function pairs(a: Grid, b: Grid): [number[], number[]] | CErr {
  const n = a.rows * a.cols
  if (n !== b.rows * b.cols) return E.NA
  // past the occupied parts every pair is blank
  const extent = (g: Grid) => (g.h === 0 ? 0 : (g.h - 1) * g.cols + g.w)
  const end = Math.min(n, Math.max(extent(a), extent(b)))
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
