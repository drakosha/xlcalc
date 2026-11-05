/**
 * The second round of `file-refusals`, over what its answers left open on Excel for Windows. Every function of the whitelist in a
 * reference's place, since the first round showed `_xlfn.` is no rule: `COUNTIF(UNIQUE(…),1)` costs the file and `COUNTIF(TAKE(…),1)`
 * does not. The boundaries that round crossed without finding: how deep calls and parentheses go, 8,192 characters of a file's text,
 * the number format and font counts. What opens a table's totals and a slicer.
 *
 * `refused` and `predicted` were committed before Excel ran this set, with the first round's Windows answers in hand. The model
 * behind them: in a reference's place the five old functions and the calls of 365 that return what they were given (`IFS`, `SWITCH`,
 * `XLOOKUP`, `LET`, the `TAKE` family) open and all others are refused; a file's text is limited to 8,192 characters as it stands; a
 * totals cell with neither function nor label costs the file while a label must be the cell's text.
 */
import { MAX_FORMULA_LENGTH } from '../../../src/contract.js'
import { FUNCTION_ARITY } from '../../../src/function-arity.js'
import { arityOf } from '../../../src/arity.js'
import { toFile } from '../../../src/xlsx/formulas/to-file.js'
import { fileTree } from '../../../src/xlsx/write/formula.js'
import type { LinksCase, LinksCases } from './cases.js'
import { BASE_XF, CACHE, cell, cfEdit, DXF, dvEdit, FIVE, inSheet, labelled, list, manyFonts, numFmts, one, onStyle, pivoted, REFUSED, rep, STYLES, subs, total, XFS } from './file-refusals.js'
import { CALLS } from './no-prefix.js'

// --- every function in a reference's place

/** The calls predicted to open in a reference's place. */
const OPENS = new Set(['IF', 'CHOOSE', 'INDEX', 'INDIRECT', 'OFFSET', 'IFS', 'SWITCH', 'XLOOKUP', 'LET', 'LAMBDA', 'TAKE', 'DROP', 'EXPAND', 'CHOOSECOLS', 'CHOOSEROWS', 'TRIMRANGE'])

/** Calls that need more than references: a function of the second kind goes by what the no-prefix set calls it with. */
const OWN: Record<string, string> = { LET: 'LET(x,B1:B3,x)', LAMBDA: 'LAMBDA(x,x)(B1:B3)', INDIRECT: 'INDIRECT("B1:B3")', OFFSET: 'OFFSET(B1,0,0,3)', IF: 'IF(TRUE,B1:B3)', CHOOSE: 'CHOOSE(1,B1:B3)', INDEX: 'INDEX(B1:B3,1)', IFS: 'IFS(TRUE,B1:B3)', SWITCH: 'SWITCH(1,1,B1:B3)', XLOOKUP: 'XLOOKUP(1,B1:B3,C1:C3)', TAKE: 'TAKE(B1:B3,1)', DROP: 'DROP(B1:B3,1)', EXPAND: 'EXPAND(B1:B3,3)', CHOOSECOLS: 'CHOOSECOLS(B1:C3,1)', CHOOSEROWS: 'CHOOSEROWS(B1:C3,1)', TRIMRANGE: 'TRIMRANGE(B1:B3)' }

const FILE_NAMES = { sheets: new Set(['S']), tables: [] }
const asFile = (formula: string) => {
  const tree = fileTree(formula, FILE_NAMES)
  if (!tree) throw new Error(`file-refusals-2: the writer takes no ${formula}`)
  return toFile(tree)
}
const callOf = (fn: string) => OWN[fn] ?? CALLS[fn] ?? `${fn}(${list(arityOf(FUNCTION_ARITY[fn]).min, () => 'B1', ',')})`

const everyFunction = (): LinksCase[] =>
  Object.keys(FUNCTION_ARITY)
    .sort()
    .map((fn) => cell(`fn-${fn.toLowerCase().replaceAll('.', '-')}`, `COUNTIF over a call of ${fn}`, `COUNTIF(${asFile(callOf(fn))},1)`, OPENS.has(fn) ? {} : REFUSED))

// --- how deep and how long

const nested = (n: number) => `${'ABS('.repeat(n)}1${')'.repeat(n)}`
const parens = (n: number) => `${'('.repeat(n)}1${')'.repeat(n)}`
/** A formula of exactly `n` characters, spaces padding `tail` after `head`. */
const padded = (head: string, n: number, tail = '') => `${head}${rep(' ', n - head.length - tail.length)}${tail}`
/** Calls joined by `+` up to about `n` characters, the rest in spaces: the formula is `n` long as a file holds it. */
function chain(unit: string, n: number): string {
  const count = Math.floor((n - 1) / (unit.length + 1))
  return padded(list(count, () => unit, '+'), n)
}
const QTY = 'Sales[[#This Row],[Qty]]'
const SALES = { cells: { E1: { value: 'Qty' }, F1: { value: 'F' }, E2: { value: 1 } }, sheet: { tables: [{ name: 'Sales', range: 'E1:F2', columns: [{ name: 'Qty' }, { name: 'F' }] }] } }
const ruleText = (n: number) => padded('B1>0', n)

const long = (): LinksCase[] => [
  ...[66, 67, 68, 69].map((n) => cell(`nest-${n}`, `${n} calls one in another`, nested(n), REFUSED)),
  ...[100, 128, 200, 255].map((n) => cell(`parens-${n}`, `${n} parentheses`, parens(n), { value: 1 })),
  cell('parens-256', '256 parentheses', parens(256), REFUSED),
  cell('text-spaces-8192', `${MAX_FORMULA_LENGTH} characters, mostly spaces`, padded('1+1', MAX_FORMULA_LENGTH), { value: 2 }),
  cell('text-spaces-8193', '8,193 characters, mostly spaces', padded('1+1', MAX_FORMULA_LENGTH + 1), REFUSED),
  cell('text-xlfn-8192', '8,192 characters of XLOOKUP with its prefix', chain('_xlfn.XLOOKUP(1,B1:B2,C1:C2)', MAX_FORMULA_LENGTH)),
  cell('text-xlfn-8193', '8,193 characters of XLOOKUP with its prefix', chain('_xlfn.XLOOKUP(1,B1:B2,C1:C2)', MAX_FORMULA_LENGTH + 1), REFUSED),
  one('text-table-8192', '8,192 characters of a table’s this-row references', { F2: chain(QTY, MAX_FORMULA_LENGTH) }, SALES),
  one('text-table-8193', '8,193 characters of a table’s this-row references', { F2: chain(QTY, MAX_FORMULA_LENGTH + 1) }, { ...SALES, ...REFUSED }),
  one('text-cf-8192', 'a rule of 8,192 characters', {}, { edit: cfEdit(ruleText(MAX_FORMULA_LENGTH)), holds: 'type="expression"' }),
  one('text-cf-8193', 'a rule of 8,193 characters', {}, { edit: cfEdit(ruleText(MAX_FORMULA_LENGTH + 1)), ...REFUSED }),
  one('text-dv-8192', 'a validation rule of 8,192 characters', {}, { edit: dvEdit('custom', ruleText(MAX_FORMULA_LENGTH)), holds: 'type="custom"' }),
  one('text-dv-8193', 'a validation rule of 8,193 characters', {}, { edit: dvEdit('custom', ruleText(MAX_FORMULA_LENGTH + 1)), ...REFUSED }),
  cell('text-name-8192', 'a name of 8,192 characters', 'k', { names: [{ name: 'k', ref: padded('1+1', MAX_FORMULA_LENGTH) }], value: 2 }),
  cell('text-name-8193', 'a name of 8,193 characters', 'k', { names: [{ name: 'k', ref: padded('1+1', MAX_FORMULA_LENGTH + 1) }], ...REFUSED }),
]

// --- the rule forms the first round refused

const rangeOp = (): LinksCase[] => [
  one('cf-range-sum', 'a plain range in a rule', {}, { edit: cfEdit('SUM(B1:B3)>0'), holds: 'type="expression"' }),
  one('cf-range-abs', 'the range operator over a call, absolute, as a dynamic range is written', {}, { edit: cfEdit('SUM($B$1:INDEX($B$1:$B$3,2))>0'), ...REFUSED }),
  one('cf-range-offset', 'OFFSET in a rule', {}, { edit: cfEdit('SUM(OFFSET(B1,0,0,2))>0'), holds: 'type="expression"' }),
  one('dv-range-op', 'the range operator over a call in a custom rule', {}, { edit: dvEdit('custom', 'SUM(B1:INDEX(B1:B3,2))>0'), ...REFUSED }),
  one('dv-list-range-op', 'a dynamic list, as people write one', {}, { edit: dvEdit('list', '$B$1:INDEX($B$1:$B$3,2)'), holds: 'type="list"' }),
]

// --- number formats, fonts

/** 207 formats with ids of its own, A1 in the last. */
const numFmtsFrom = (n: number, first: number, used: boolean) =>
  subs(
    [STYLES, '<fonts ', `<numFmts count="${n}">${list(n, (i) => `<numFmt numFmtId="${first + i}" formatCode="0.00&quot;i${i + 1}&quot;"/>`)}</numFmts><fonts `],
    ...(used ? ([[STYLES, XFS, `<cellXfs count="2">${BASE_XF}<xf numFmtId="${first + n - 1}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>`], onStyle(1)] as [string, string | RegExp, string][]) : []),
  )
/** `n` conditional styles with a number format each: what the 6 books over 206 of the corpus hold, none of them in `numFmts`. */
const dxfFormats = (n: number) => subs(inSheet('<conditionalFormatting sqref="D1"><cfRule type="expression" dxfId="0" priority="1"><formula>B1&gt;0</formula></cfRule></conditionalFormatting>'), [DXF[0], DXF[1], `<dxfs count="${n}">${list(n, (i) => `<dxf><numFmt numFmtId="${200 + i}" formatCode="0.00&quot;d${i}&quot;"/></dxf>`)}</dxfs></styleSheet>`])

const styles = (): LinksCase[] => [
  one('numfmt-207-gap', '207 number formats with ids from 1,000', {}, { ...FIVE, edit: numFmtsFrom(207, 1000, true), ...REFUSED }),
  one('numfmt-207-unused', '207 number formats, none in use', {}, { edit: numFmtsFrom(207, 164, false), ...REFUSED }),
  one('numfmt-206-first', '206 number formats, as the first round with its last in use', {}, { ...FIVE, edit: numFmts(206), holds: 'i206' }),
  one('numfmt-dxf-207', '207 number formats in conditional styles', {}, { edit: dxfFormats(207), holds: 'cfRule' }),
  one('numfmt-dxf-600', '600 number formats in conditional styles, as saved books hold them', {}, { edit: dxfFormats(600), holds: 'cfRule' }),
  one('font-65430', '65,430 fonts', {}, { edit: manyFonts(65430) }),
  one('font-65431', '65,431 fonts', {}, { edit: manyFonts(65431) }),
  one('font-65490', '65,490 fonts', {}, { edit: manyFonts(65490) }),
]

// --- a table's totals

const LABELLED = (text: string, label: string) => ({ cells: { E1: { value: 'Q' }, E2: { value: 1 }, E3: { value: text } }, sheet: { tables: [{ name: 'Tbl', range: 'E1:E3', totals: true, columns: [{ name: 'Q', totals: { label: 'L' } }] }] }, edit: subs(['xl/tables/table1.xml', 'totalsRowLabel="L"', `totalsRowLabel="${label}"`]) })
const NUMBER_CUSTOM = subs(['xl/tables/table1.xml', '<tableColumn id="1" name="Q"/>', '<tableColumn id="1" name="Q" totalsRowFunction="custom"><totalsRowFormula>42</totalsRowFormula></tableColumn>'])

const totals = (): LinksCase[] => [
  total('totals-none-text', 'a text in the totals row, no label', '', {}, { cells: { E1: { value: 'Q' }, E2: { value: 1 }, E3: { value: 'T' } }, ...REFUSED }),
  total('totals-none-empty', 'an empty totals row, no function, no label', '', {}),
  total('totals-custom-const', 'a value in the totals row with the value as its formula', '', { value: 42 }, { edit: NUMBER_CUSTOM }),
  one('totals-label-same', 'a label that is the cell’s text', {}, LABELLED('Total', 'Total')),
  one('totals-label-case', 'a label that differs from the cell’s text in case', {}, { ...LABELLED('Total', 'total'), ...REFUSED }),
  one('totals-label-number', 'a label over a number of its text', {}, { ...LABELLED('Total', '42'), cells: { E1: { value: 'Q' }, E2: { value: 1 }, E3: { value: 42 } }, ...REFUSED }),
  one('totals-label-empty', 'a label over an empty cell', {}, { ...LABELLED('', 'Total'), cells: { E1: { value: 'Q' }, E2: { value: 1 } }, ...REFUSED }),
  labelled('totals-label-long', 'a label of 300 over a cell’s 300', 300, { cells: { E1: { value: 'Q' }, E2: { value: 1 }, E3: { value: rep('L', 300) } } }),
]

// --- slicers

const WITH_ITEMS = (b: { slicerCaches?: { items?: { item: number }[] }[] }) => {
  b.slicerCaches![0].items = [0, 1, 2, 3].map((item) => ({ item }))
}

const slicers = (): LinksCase[] => [
  pivoted('slicer-items', 'a slicer with its items, no defined name', { change: WITH_ITEMS, ...REFUSED }),
  pivoted('slicer-items-name', 'a slicer with its items and the defined name', { change: WITH_ITEMS, names: [{ name: 'Slicer_Region', ref: '#N/A' }], holds: 'Slicer_Region' }),
  pivoted('pivot-source-missing', 'a source sheet of 31 the book does not have', { slicer: false, edit: subs([CACHE, 'sheet="Data"', `sheet="${rep('s', 31)}"`]) }),
]

export function buildLinksCases(): LinksCases {

  return { cases: [...everyFunction(), ...long(), ...rangeOp(), ...styles(), ...totals(), ...slicers()] }
}
