/**
 * What makes Excel refuse a whole file over one formula, as the sets `bench/coverage/links/file-refusals*.ts` measured it on Excel
 * for Windows and for Mac. A refusal by either counts: files are for the oldest Excel that opens them (Excel for Mac 16.114 takes a
 * longer text literal as `_LONGTEXT`, Excel for Windows 16.0.20430 refuses the file). The specifications say more than Excel
 * does: nothing here that the sets did not show.
 */
import { holds, type Positions } from '../../arity.js'
import { FUNCTION_COUNTS } from '../../function-words.js'
import { isRefOp, type Ast } from '../../syntax/ast.js'
import { walk } from '../../syntax/print.js'

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

/** The functions Excel takes in a reference's place, whatever they return there; a function it does not have is a user's call. */
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
      return !FUNCTION_COUNTS.has(a.name) || MAY_RETURN_REFERENCE.has(a.name)
    default:
      return true
  }
}

/** The number of arguments of a built-in Excel takes, the engine's but where the sets showed otherwise. */
function takesCount(name: string, k: number): boolean {
  const c = FUNCTION_COUNTS.get(name)
  if (!c) return k <= MAX_USER_ARGS
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
