/**
 * Held out from `file-refusals-4`, which showed the limit of a formula's bytes on two families alone: each kind of token on both sides of
 * 16,384 bytes, the rest of its formula cells and small numbers, whose sizes that set measured. What is asked of each is whether its size
 * is the one [MS-XLSB] gives: a miss on the side under the limit is a token bigger than it says, which the writer's count (`fitsTokens`,
 * `src/xlsx/write/excel-loads.ts`) could let past. The formulas are grouped in `SUM`s of 250, as the parser takes 2,000 operators.
 *
 * `refused` was committed before Excel ran this set. The model: a formula holds up to 16,384 bytes of tokens, its data not counting, at
 * the sizes of [MS-XLSB] 2.5.98: a cell 7, a range 13, either on another sheet or the formula's own by name 9 and 15, across sheets and
 * in another workbook the same; an integer to 65,535 3, another number 9; text 3 and 2 for each unit; a boolean or an error 2; a sheet's
 * lost reference 9; a name 5, a sheet's name or another workbook's 7; an operator, a parenthesis or an argument left out 1; a function of
 * Excel's table of fixed count 3, of variable count 4, one with a prefix or of none 5 more; `IF` and `IFERROR` 4 for each control token,
 * `CHOOSE` 2 for each entry of its table on top; a volatile function 4, once a formula; a table's reference 14; an array constant 15; a
 * union or an intersection, its operator 1 and a token of 3 around it. A rule and a name by the same count, and a dynamic array.
 */
import { list, one, cell, inSheet, subs, REFUSED, type Extra } from './file-refusals.js'
import type { LinkedFormula, LinksCase, LinksCases } from './cases.js'

const LIMIT = 16_384
const GROUP = 250
const MAX_TEXT = 8192
const SUM = 4
const PLUS = 1

interface Unit {
  text: string
  bytes: number
}
const U = (text: string, bytes: number): Unit => ({ text, bytes })

/** A formula of `count` of `unit` and as many of `fill` and of the int `1` as make `target` bytes, `extra` of them outside the `SUM`s. */
function exact(unit: Unit, count: number, target: number, fill: Unit, extra: number): string {
  const size = (items: number, bytes: number) => {
    const groups = Math.ceil(items / GROUP)
    return bytes + SUM * groups + PLUS * (groups - 1) + extra
  }
  const known = count * unit.bytes
  const times = (n: number, text: string) => Array.from({ length: n }, () => text)
  for (const parens of [0, 1, 2])
    for (let n = Math.floor((target - known) / fill.bytes); n >= 0; n--)
      for (let ints = 0; ints < 40; ints++) {
        const items = count + n + ints + parens
        if (size(items, known + n * fill.bytes + ints * 3 + parens * 4) !== target) continue
        const texts = [...times(count, unit.text), ...times(n, fill.text), ...times(ints, '1'), ...times(parens, '(1)')]
        const groups: string[] = []
        for (let i = 0; i < texts.length; i += GROUP) groups.push(`SUM(${texts.slice(i, i + GROUP).join(',')})`)
        return groups.join('+')
      }
  throw new Error(`formula-bytes: no formula of ${target} bytes over ${unit.text}`)
}

const B1 = U('B1', 7)
const SHEET_CELLS = { C1: { value: 2 }, E1: { value: 'Qty' }, F1: { value: 'F' }, E2: { value: 1 } }
const MORE: Extra = { more: [{ name: 'T', cells: { B1: { value: 1 }, B2: { value: 2 } } }, { name: 'U', cells: { B1: { value: 1 } } }] }
const NAMES: Extra = { names: [{ name: 'rate', ref: 'S!$B$1' }, { name: 'sn', ref: 'S!$B$1', sheet: 'S' }] }
const LINK: Extra = { links: [{ target: 'other.xlsx', sheets: [{ name: 'T', cells: { B1: 1, B2: 2 } }], names: [{ name: 'ext', ref: '=T!$B$1' }] }] }
const TABLE: Extra = { cells: SHEET_CELLS, sheet: { tables: [{ name: 'Sales', range: 'E1:F2', columns: [{ name: 'Qty' }, { name: 'F' }] }] } }

/** The count of a unit a formula of the limit can hold within 8,192 characters. */
function fitting(u: Unit, fill: Unit): number {
  const perChar = fill.bytes / (fill.text.length + 1)
  const net = u.text.length + 1 - u.bytes / perChar
  const room = MAX_TEXT - LIMIT / perChar - 80
  return Math.min(400, net <= 0 ? 400 : Math.floor(room / net))
}

/** Both sides of the limit for a unit. */
function family(id: string, about: string, u: Unit, e: Extra = {}, fill: Unit = B1, extra = 0): LinksCase[] {
  const count = fitting(u, fill)
  return ([['open', LIMIT, {}], ['over', LIMIT + 1, REFUSED]] as const).map(([side, target, more]) => {
    const f = exact(u, count, target, fill, extra)
    if (f.length > MAX_TEXT) throw new Error(`formula-bytes: ${id} is ${f.length} characters`)
    return cell(`${id}-${side}`, `${about}: ${count} of ${u.bytes} bytes, ${target} in all, ${f.length} characters`, f, { ...e, ...more })
  })
}

const QUOTED = (n: number, c = 'a') => `"${c.repeat(n)}"`

const FAMILIES: [string, string, Unit, Extra?][] = [
  ['cell', 'cells alone', U('C1', 7)],
  ['cell-abs', 'absolute cells', U('$C$1', 7)],
  ['add', 'an operator between two cells', U('C1+B1', 15)],
  ['paren', 'a parenthesis', U('(B1)', 8)],
  ['minus', 'a sign', U('-1', 4)],
  ['plus-sign', 'a plus sign', U('+B1', 8)],
  ['percent', 'a percent', U('1%', 4)],
  ['range', 'a range', U('B1:B2', 13)],
  ['column', 'a whole column', U('B:B', 13)],
  ['row', 'a whole row', U('1:1', 13)],
  ['int-top', 'the largest integer of its token', U('65535', 3)],
  ['int-past', 'an integer past it', U('65536', 9)],
  ['decimal', 'a number with a point', U('1.5', 9)],
  ['bool', 'a boolean', U('TRUE', 2)],
  ['error', 'an error', U('#N/A', 2)],
  ['error-ref', 'a lost reference', U('#REF!', 2)],
  ['error-sheet', 'a lost reference of a sheet', U('T!#REF!', 9), MORE],
  ['text-short', 'short text', U(QUOTED(2), 7)],
  ['text-empty', 'empty text', U('""', 3)],
  ['text-quote', 'a doubled quote', U('""""', 5)],
  ['text-long', 'long text', U(QUOTED(100), 203)],
  ['text-cyrillic', 'Cyrillic text', U(QUOTED(100, 'я'), 203)],
  ['text-255', 'text of the most', U(QUOTED(255), 513)],
  ['sheet-cell', 'a cell of another sheet', U('T!B1', 9), MORE],
  ['sheet-range', 'a range of another sheet', U('T!B1:B2', 15), MORE],
  ['sheet-own', 'a cell by its own sheet’s name', U('S!B1', 9)],
  ['across-cell', 'a cell across sheets', U("'T:U'!B1", 9), MORE],
  ['across-range', 'a range across sheets', U("'T:U'!B1:B2", 15), MORE],
  ['link-cell', 'a cell of another workbook', U('[1]T!B1', 9), LINK],
  ['link-range', 'a range of another workbook', U('[1]T!B1:B2', 15), LINK],
  ['link-name', 'a name of another workbook', U('[1]!ext', 7), LINK],
  ['name', 'a name', U('rate', 5), NAMES],
  ['name-sheet', 'a sheet’s name', U('S!sn', 7), NAMES],
  ['fixed', 'a function of fixed count', U('ABS(B1)', 10)],
  ['fixed-none', 'a function of no argument', U('PI()', 3)],
  ['fixed-two', 'a function of fixed count of two', U('MOD(B1,2)', 13)],
  ['var', 'a function of variable count', U('MAX(B1,C1)', 18)],
  ['var-one', 'one argument of a variable count', U('MAX(B1)', 11)],
  ['sum-one', 'a SUM of one argument', U('SUM(B1)', 11)],
  ['missing', 'an argument left out', U('MAX(,)', 6)],
  ['future', 'a function with a prefix', U('_xlfn.CONCAT(B1)', 16)],
  ['future-ws', 'a function with both prefixes', U('_xlfn._xlws.SORT(B1:B2)', 22)],
  ['user', 'a call of a user’s function', U('own(B1)', 16)],
  ['if-two', 'IF of two', U('IF(B1,1)', 22)],
  ['if-three', 'IF of three', U('IF(B1,1,2)', 29)],
  ['iferror', 'IFERROR', U('IFERROR(B1,1)', 22)],
  ['ifna', 'IFNA with its prefix', U('_xlfn.IFNA(B1,1)', 27)],
  ['ifs', 'IFS with its prefix', U('_xlfn.IFS(B1,1)', 19)],
  ['switch', 'SWITCH with its prefix', U('_xlfn.SWITCH(B1,1,2)', 22)],
  ['choose', 'CHOOSE of two', U('CHOOSE(1,B1,C1)', 39)],
  ['array', 'an array constant', U('{1,2}', 15)],
  ['array-rows', 'an array of rows', U('{1;2;3}', 15)],
  ['table', 'a table’s column', U('Sales[Qty]', 14), TABLE],
  ['table-row', 'a table’s row of a column', U('Sales[[#This Row],[Qty]]', 14), TABLE],
  ['union', 'a union', U('(B1,C1)', 19)],
  ['intersection', 'an intersection', U('B1:B2 B1:C1', 30)],
  ['range-op', 'the range operator over a call', U('B1:INDEX(B1:B2,1)', 31)],
  ['single', 'an implicit intersection', U('_xlfn.SINGLE(B1:B2)', 22)],
  ['spill', 'a spill reference', U('_xlfn.ANCHORARRAY(B1)', 16)],
  ['let', 'LET of one name', U('_xlfn.LET(_xlpm.x,1,_xlpm.x)', 22)],
  ['lambda', 'a LAMBDA called', U('_xlfn.LAMBDA(_xlpm.x,_xlpm.x)(1)', 26)],
]

/** Every family in one formula: the mix of `step`, each once in turn. */
function mixed(step: number, side: string, target: number, more: Extra): LinksCase {
  const plain = FAMILIES.filter(([id]) => !['error-sheet', 'link-cell', 'link-range', 'link-name', 'table', 'table-row', 'name', 'name-sheet', 'sheet-cell', 'sheet-range', 'across-cell', 'across-range', 'text-255', 'text-long', 'text-cyrillic'].includes(id)).filter((_, i) => i % step === 0)
  const u = U(plain.map(([, , x]) => x.text).join(','), plain.reduce((n, [, , x]) => n + x.bytes, 0))
  const f = exact(u, 6, target, B1, 0)
  return cell(`mix-${step}-${side}`, `every ${step}th family six times over, ${target} in all, ${f.length} characters`, f, more)
}

/** A rule's and a name's formula, and a dynamic array's, on both sides. */
const wrapped = (): LinksCase[] => {
  const SIDES = [['open', LIMIT, {}], ['over', LIMIT + 1, REFUSED]] as const
  const SHEET_B1 = U('S!B1', 9)
  const FIT = 100
  return SIDES.flatMap(([side, target, more]) => {
    const rule = exact(U('C1', 7), FIT, target - 4, B1, 0)
    const named = exact(U('S!C1', 9), FIT, target, SHEET_B1, 0)
    const dynamic = exact(U('C1', 7), FIT, target, B1, 0)
    const FORMULA: LinkedFormula = { formula: dynamic, array: 'A1' }
    return [
      one(`cf-${side}`, `a conditional format’s formula, ${target} in all`, {}, { edit: subs(inSheet(`<conditionalFormatting sqref="D1"><cfRule type="expression" dxfId="0" priority="1"><formula>${rule}&gt;0</formula></cfRule></conditionalFormatting>`), ['xl/styles.xml', /<dxfs count="0"\/>|<\/styleSheet>/, '<dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>']), ...more }),
      one(`dv-${side}`, `a validation’s formula, ${target} in all`, {}, { edit: subs(inSheet(`<dataValidations count="1"><dataValidation type="custom" sqref="D1"><formula1>${rule}&gt;0</formula1></dataValidation></dataValidations>`)), ...more }),
      cell(`defined-${side}`, `a name’s formula, ${target} in all`, 'k', { names: [{ name: 'k', ref: named }], ...more }),
      one(`dynamic-${side}`, `a dynamic array of ${target} in all`, { A1: FORMULA } as unknown as Record<string, string>, more),
    ]
  })
}

/** A volatile function in a formula: its token once. */
const volatile = (): LinksCase[] =>
  ([['open', LIMIT, {}], ['over', LIMIT + 1, REFUSED]] as const).map(([side, target, more]) => {
    const f = exact(U('NOW()', 3), 40, target, B1, 4)
    return cell(`volatile-${side}`, `40 volatile calls, ${target} in all with the token of 4`, f, more)
  })

/** An array constant's data counts for nothing: 8,000 characters of it. */
const data = (): LinksCase[] => [
  cell('array-long', 'SUM of an array of 3,990 numbers, 35,000 bytes of data beside 19 of tokens', `SUM({${list(3990, () => '1', ',')}})`, { value: 3990 }),
  cell('array-long-text', 'an array of 400 texts of 16', `COUNTA({${list(400, () => QUOTED(16), ',')}})`, { value: 400 }),
]

export function buildLinksCases(): LinksCases {
  return {
    cases: [
      ...FAMILIES.flatMap(([id, about, u, e]) => family(id, about, u, e)),
      ...[2, 3, 5].flatMap((step) => [mixed(step, 'open', LIMIT, {}), mixed(step, 'over', LIMIT + 1, REFUSED)]),
      ...volatile(),
      ...wrapped(),
      ...data(),
    ],
  }
}
