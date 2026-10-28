/**
 * Paste Special's operation (`copyRange` with `operation`), as Excel pastes it (the paste set `ops`): a cell of the copy
 * combined with the cell of the destination it lands on. A blank counts as 0. Two numbers give a number, as Excel's
 * arithmetic gives it (10 divided by a blank is #DIV/0!); a formula on either side gives a formula of both, the
 * destination's first, a number as it is and a formula in parentheses (10 plus `B1*2` is `10+(B1*2)`). Text, booleans
 * and errors of the copy, or of the destination, leave the destination's cell as it was, and so do two blanks.
 */
import type { CellInput, PasteOperation } from '../contract.js'
import { ERROR_TEXT } from '../syntax.js'
import { E, fitNumber, isErr, nearZero } from '../values.js'

/** What a cell is to an operation: a number (a blank is 0), a formula, or null for what it leaves. */
type Side = { n: number } | { f: string } | null

const SIGN: Record<PasteOperation, string> = { add: '+', subtract: '-', multiply: '*', divide: '/' }

/** An error constant, as a cell holds it (`#N/A` read from a file, a paste of values' error). */
const ERRORS = new Set<string>(Object.values(ERROR_TEXT))

/** Whether a cell holds no value or formula. */
export const isBlank = (cell: CellInput | undefined): boolean => cell === undefined || (cell.formula === undefined && cell.value == null)

function side(cell: CellInput | undefined): Side {
  if (isBlank(cell)) return { n: 0 }
  if (cell!.formula !== undefined) return ERRORS.has(cell!.formula) ? null : { f: cell!.formula }
  return typeof cell!.value === 'number' ? { n: cell!.value } : null
}

/** Excel's answer of two numbers: an addition's or subtraction's 0 near the destination's number, a division by 0 its error. */
function answer(op: PasteOperation, a: number, b: number): CellInput {
  const r = op === 'add' ? nearZero(a, a + b) : op === 'subtract' ? nearZero(a, a - b) : op === 'multiply' ? a * b : b === 0 ? E.DIV0 : a / b
  const v = isErr(r) ? r : fitNumber(r)
  return isErr(v) ? { formula: ERROR_TEXT[v.code] } : { value: v }
}

/**
 * The contents of the destination's cell `t` after the copy's cell `x` is combined with it: a value or a formula, or
 * null where the destination's cell stays as it was.
 */
export function combined(op: PasteOperation, t: CellInput | undefined, x: CellInput | undefined): CellInput | null {
  const [a, b] = [side(t), side(x)]
  if (!a || !b || (isBlank(t) && isBlank(x))) return null
  if ('n' in a && 'n' in b) return answer(op, a.n, b.n)
  const text = (s: { n: number } | { f: string }) => ('n' in s ? String(s.n) : `(${s.f})`)
  return { formula: text(a) + SIGN[op] + text(b) }
}
