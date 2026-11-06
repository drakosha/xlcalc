/**
 * Why `formula-bytes-held` missed: formulas of 16,384 bytes by the sizes of [MS-XLSB] were refused wherever they held functions' arguments,
 * and a chain of operators of 16,383 had opened. This set asks one thing of the hypothesis that an argument costs a byte: each family of
 * functions Excel 2003 had on both sides of 16,384 of bytes and arguments together, each cell of a `SUM` an argument of it, each
 * argument inside a call too.
 *
 * What the files say of a price: nothing. [MS-XLSB] 2.5.98.88 has `function-call` as its parameters and `PtgFuncVar` (`cparams`, 1 byte,
 * and `tab`, 2), `PtgFunc` (`iftab`, 2) or `PtgAttrSum`; a comma leaves no token, `PtgAttrSpace` stands for white space alone and goes on
 * no argument, `RgbExtra` holds an array's data and a union's ranges and no argument's. `cce`, the length of `rgce`, stays under 16,385.
 * ECMA-376 has only the formula's text. So the files give the sizes of `formula-bytes-held`, and the limit stands on another form, Excel's
 * own in memory, which they do not describe: a byte for each argument is a hypothesis from the answers until this set measures it.
 *
 * Committed before Excel ran it, the rule of decision: the hypothesis is taken if every family here opens on 16,384 and is refused on
 * 16,385, the chain with its arguments of none included. A single miss, on either side, and it is not taken, with no fitting to what
 * came back: the writer's count then is a margin (bytes and arguments over 16,384 are over, more if that does not stop them all), not a limit.
 * Functions with a prefix and `IFERROR`, `SINGLE`, `ANCHORARRAY`, `LET` opened on both sides in `formula-bytes-held` and are left out of the rule, as of this set.
 */
import type { LinksCase, LinksCases } from './cases.js'
import { cell, list, REFUSED, type Extra } from './file-refusals.js'
import { exact, U, type Unit } from './formula-bytes-held.js'

const LIMIT = 16_384
const B1 = U('B1', 7)
const MORE: Extra = { more: [{ name: 'T', cells: { B1: { value: 1 } } }] }

/** Each unit's bytes and the arguments inside it. */
const FAMILIES: [string, string, Unit, number, Extra?][] = [
  ['cell', 'cells', U('C1', 7), 100],
  ['range', 'ranges', U('B1:B2', 13), 300],
  ['number', 'numbers with a point', U('1.5', 9), 400],
  ['integer', 'small integers', U('2', 3), 400],
  ['text-short', 'short text', U('"ab"', 7), 200],
  ['text-long', 'long text', U(`"${'a'.repeat(100)}"`, 203), 40],
  ['sheet', 'cells of another sheet', U('T!B1', 9), 300, MORE],
  ['max', 'MAX of two', U('MAX(B1,C1)', 18 + 2), 200],
  ['abs', 'ABS of a cell', U('ABS(B1)', 10 + 1), 200],
  ['nested', 'calls in calls', U('ABS(MAX(B1,ABS(C1)))', 24 + 4), 150],
  ['if', 'IF of three', U('IF(B1,1,2)', 29 + 3), 150],
  ['mixed', 'all of them in turn', U(`C1,B1:B2,1.5,2,"ab",MAX(B1,C1),ABS(B1),ABS(MAX(B1,ABS(C1))),IF(B1,1,2)`, 7 + 13 + 9 + 3 + 7 + 20 + 11 + 28 + 32 + 8), 40],
]

const SIDES = [['open', LIMIT, {}], ['over', LIMIT + 1, REFUSED]] as const

function family([id, about, u, count, e]: (typeof FAMILIES)[number]): LinksCase[] {
  return SIDES.map(([side, target, more]) => {
    const f = exact(u, count, target, B1, 0, 1)
    if (f.length > 8192) throw new Error(`${id} is ${f.length} characters`)
    return cell(`${id}-${side}`, `${about}: ${target} of bytes and arguments, ${f.length} characters`, f, { ...e, ...more })
  })
}

/** No argument at all: operators and parentheses alone, as the chain of `file-refusals-4`. */
const chain = (): LinksCase[] => SIDES.map(([side, target, more]) => cell(`chain-${side}`, `2,048 cells joined with ${target - 16_383} in parentheses, ${target} of bytes`, `${list(target - 16_383, () => 'B1', '+').replaceAll('B1', '(B1)')}+${list(2048 - (target - 16_383), () => 'B1', '+')}`, more))

export function buildLinksCases(): LinksCases {
  return { cases: [...chain(), ...FAMILIES.flatMap(family)] }
}
