/**
 * What makes Excel refuse a whole file over one formula, as the sets `bench/coverage/links/file-refusals*.ts` measured it on Excel
 * for Windows and for Mac. A refusal by either counts: files are for the oldest Excel that opens them (Excel for Mac 16.114 takes a
 * longer text literal as `_LONGTEXT`, Excel for Windows 16.0.20430 refuses the file). The specifications say more than Excel
 * does: nothing here that the sets did not show.
 */
import { holds, type Positions } from '../../arity.js'
import { FUNCTION_COUNTS } from '../../function-words.js'
import { isRefOp, type Ast, type RangeRef } from '../../syntax/ast.js'
import { walk } from '../../syntax/print.js'
import { XLFN, XLWS } from '../formulas/prefixes.js'

/** The UTF-16 units of a text literal, once its quotes are undoubled. */
const MAX_LITERAL = 255
/** The arguments of a call of a user's function, one fewer than a built-in's. */
const MAX_USER_ARGS = 254

const positions = (names: string, p: Positions) => names.split(' ').map((n) => [n, p] as const)

/** The arguments only a reference fills, by [MS-XLSX] 2.2.2 `ref-argument-expression`, which the sets confirmed at each position. */
const REFERENCE_ONLY: ReadonlyMap<string, Positions> = /* @__PURE__ */ new Map([
  ...positions('AREAS COLUMN COUNTBLANK COUNTIF DAVERAGE DCOUNT DCOUNTA DGET DMAX DMIN DPRODUCT DSTDEV DSTDEVP DSUM DVAR DVARP FORMULATEXT ISFORMULA OFFSET PHONETIC ROW', { at: [0] }),
  ...positions('SUMIF AVERAGEIF', { at: [0, 2] }),
  ...positions('SUMIFS AVERAGEIFS MAXIFS MINIFS', { at: [0], from: 1, odd: true }),
  ...positions('COUNTIFS', { from: 0, even: true }),
  ...positions('CELL RANK RANK.AVG RANK.EQ', { at: [1] }),
  ...positions('SUBTOTAL', { from: 1 }),
  ...positions('AGGREGATE', { from: 4 }),
])

/** Functions of the whitelist no Excel measured has: a call of any count, anywhere, is a user's call, `#NAME?`. */
const NOT_IN_EXCEL: ReadonlySet<string> = /* @__PURE__ */ new Set('EUROCONVERT JIS REGEXEXTRACT REGEXREPLACE REGEXTEST'.split(' '))

/** The functions Excel takes in a reference's place, whatever they return there. */
const MAY_RETURN_REFERENCE: ReadonlySet<string> = /* @__PURE__ */ new Set('CHOOSE DROP IF IFS INDEX INDIRECT OFFSET REDUCE SWITCH TAKE TRIMRANGE XLOOKUP'.split(' '))

function fillsReference(a: Ast): boolean {
  switch (a.k) {
    case 'paren':
      return fillsReference(a.a)
    case 'num':
    case 'str':
    case 'bool':
    case 'arr':
    case 'un':
    case 'pct':
    case 'missing':
      return false
    case 'err':
      return a.code === 'REF'
    case 'bin':
      return isRefOp(a.op)
    case 'fn':
      return !FUNCTION_COUNTS.has(a.name) || NOT_IN_EXCEL.has(a.name) || MAY_RETURN_REFERENCE.has(a.name)
    default:
      return true
  }
}

/** The number of arguments of a built-in Excel takes, the engine's but where the sets showed otherwise. */
function takesCount(name: string, k: number): boolean {
  const c = FUNCTION_COUNTS.get(name)
  if (!c || NOT_IN_EXCEL.has(name)) return k <= MAX_USER_ARGS
  if (k === 1 && (name === 'BYROW' || name === 'BYCOL')) return true
  return k >= c[0] && k <= c[1] && (k - c[0]) % c[2] === 0
}

/** Whether Excel loads a call: its count, and a reference in each place only one fills. */
function loadsCall(n: Extract<Ast, { k: 'fn' }>): boolean {
  if (!takesCount(n.name, n.args.length)) return false
  const only = REFERENCE_ONLY.get(n.name)
  return !only || n.args.every((a, i) => !holds(only, i) || fillsReference(a))
}

/** Whether Excel loads a formula of a cell, a table, a rule or a name. */
export function excelLoads(ast: Ast): boolean {
  let ok = true
  walk(ast, (n) => {
    if (n.k === 'str') ok &&= n.v.length <= MAX_LITERAL
    else if (n.k === 'arr') ok &&= n.rows.every((row) => row.every((x) => x.k !== 'str' || x.v.length <= MAX_LITERAL))
    else if (n.k === 'fn') ok &&= loadsCall(n)
    else if (n.k === 'call') ok &&= n.args.length <= MAX_USER_ARGS
  })
  return ok
}

/**
 * What the writer takes of a formula's tokens and arguments together (`fitsTokens`). Not Excel's limit: Microsoft gives 16,384 bytes,
 * and [MS-XLSB]'s sizes do not add up to what Excel refuses. A chain of operators opens at 16,383 and is refused at 16,384; a function's
 * arguments cost something the files do not show, about a byte each and not exactly (`links/formula-bytes-held`, `-args`, on Excel for
 * Windows). So a margin under what the sets opened, large as no book holds a formula near it: the longest of a corpus of 1,123 is 683 characters.
 */
const MAX_WEIGHT = 15_000

/** No token and its argument takes 10 bytes for a character of a formula's text, so a text this long or shorter fits without counting. */
export const SHORT_FORMULA = 1500

/** The functions [MS-XLSB] has Excel jump over with their arguments: one control token of 4 bytes, and one after each argument. */
const JUMPS: ReadonlySet<string> = /* @__PURE__ */ new Set('IF IFERROR IFNA IFS SWITCH'.split(' '))
/** The volatile functions, which mark the formula with a token of 4 bytes, here counted for each call. */
const VOLATILE: ReadonlySet<string> = /* @__PURE__ */ new Set('CELL INDIRECT INFO NOW OFFSET RAND RANDARRAY RANDBETWEEN TODAY'.split(' '))

/** A function's name as a file spells it takes a defined name's token where it has a prefix or is not Excel's. */
const FAR_CALL = 9
const CALL = 4

/** A reference, with a sheet (`S!A1`) or a sheet's index, and a spill reference (`A1#`), a call around it. */
const refBytes = (r: RangeRef, sheet: boolean) => (r.kind === 'cell' ? (sheet ? 9 : 7) : sheet ? 15 : 13) + (r.spill ? FAR_CALL : 0)

function callBytes(n: Extract<Ast, { k: 'fn' }>): number {
  const count = n.args.length
  const near = FUNCTION_COUNTS.has(n.name) && !XLFN.has(n.name) && !XLWS.has(n.name)
  const own = near ? CALL : FAR_CALL
  // CHOOSE: its table of 2 bytes for each argument on top of what a jump takes
  if (n.name === 'CHOOSE') return own + 4 + 2 * count + 4 * count
  return own + (JUMPS.has(n.name) ? 4 * count : 0) + (VOLATILE.has(n.name) ? 4 : 0)
}

/**
 * A formula's weight: its tokens by the sizes of [MS-XLSB], the most they can take (every function of Excel's table as one of variable
 * count, one byte over a fixed count's; every jump as if each argument had it; the operators on references with the most they can
 * bring), and a byte for each argument. Exported for the tests, which hold it against what Excel refused.
 */
export function formulaWeight(ast: Ast): number {
  let bytes = 0
  walk(ast, (n) => {
    switch (n.k) {
      case 'num':
        bytes += Number.isInteger(n.v) && n.v >= 0 && n.v <= 65_535 ? 3 : 9
        break
      case 'str':
        bytes += 3 + 2 * n.v.length
        break
      case 'bool':
        bytes += 2
        break
      case 'err':
        bytes += n.sheet === undefined && n.ref === undefined && n.book === undefined ? 2 : 15
        break
      case 'ref':
        bytes += refBytes(n.ref, n.sheet !== null)
        break
      case 'ref3d':
      case 'xref':
        bytes += refBytes(n.ref, true)
        break
      case 'table':
        bytes += 14
        break
      case 'name':
        bytes += n.sheet === undefined ? 5 : 7
        break
      case 'xname':
        bytes += 7
        break
      case 'local':
      case 'eta':
        bytes += 5
        break
      case 'at':
        bytes += FAR_CALL
        break
      case 'union':
        bytes += n.items.length + 8
        break
      case 'fn':
        bytes += callBytes(n) + n.args.length
        break
      case 'un':
      case 'pct':
      case 'paren':
      case 'missing':
        bytes += 1
        break
      case 'bin':
        bytes += isRefOp(n.op) ? 8 : 1
        break
      case 'arr':
        bytes += 15
        break
      case 'let':
        bytes += FAR_CALL + 5 * n.binds.length
        break
      case 'lambda':
        bytes += FAR_CALL + 5 * n.params.length
        break
      case 'call':
        bytes += CALL + n.args.length
        break
    }
  })
  return bytes
}

/** Whether a formula is under the margin (`MAX_WEIGHT`): where it is not, it goes as its value, though Excel may well have opened it. */
export const fitsTokens = (ast: Ast): boolean => formulaWeight(ast) <= MAX_WEIGHT

/** Whether Excel loads a formula as a conditional format's or a validation's: neither takes these, in a block of Excel 2007 or of 2010. */
export function excelLoadsRule(ast: Ast): boolean {
  let ok = true
  walk(ast, (n) => {
    if (n.k === 'arr' || n.k === 'table' || n.k === 'union' || n.k === 'lambda' || (n.k === 'bin' && isRefOp(n.op))) ok = false
  })
  return ok
}

/** Whether a defined name's formula names the sheet of every reference: Excel loads no name without. */
export function namesSheets(ast: Ast): boolean {
  let ok = true
  walk(ast, (n) => {
    if (n.k === 'ref' && n.sheet === null) ok = false
  })
  return ok
}
