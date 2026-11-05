/**
 * What a file can hold and Excel may refuse: one case in a file of its own, so a refused file costs only it. The specifications
 * ([MS-XLSX] 2.2.2, [MS-OI29500]) say what a file may contain, not what Excel does, so the writer refuses only what this set shows.
 * Everything the writer may come to refuse is put in by hand (`linked`, `names`, `edit`) over a file of legal input, so the files do
 * not depend on the writer. The harness of links takes the files; the saved files are not kept (`file-refusals-kept.ts` reads them).
 *
 * `refused` and `predicted` were committed before Excel ran the set. The model behind them: Excel compiles a formula on load, so an
 * argument class or a count its function table rejects costs the file. A reference-only argument takes a reference, a reference
 * operator, or a call that may return one: the five old functions (IF, CHOOSE, INDEX, INDIRECT, OFFSET) and any `_xlfn.` or
 * user-defined call, whose return class the table does not fix; an old function that returns a value is refused. A text literal is
 * limited to 255 UTF-16 units after undoubling, a formula to 16,384 bytes of tokens and 64 levels of calls. Rules of [MS-OI29500]
 * on structure (conditional format and validation grammars, font names) hold; those on thresholds of a rule do not.
 */
import { pivotBook } from '../../../test/pivot-books.js'
import type { CachedValue, LinkedBook, LinksCase, LinksCases } from './cases.js'

export type Edit = NonNullable<LinksCase['edit']>
export type Extra = Partial<Pick<LinksCase, 'refused' | 'holds' | 'names' | 'links' | 'book' | 'edit' | 'read'>> & { value?: CachedValue; cells?: LinksCase['sheets'][number]['cells']; sheet?: Partial<LinksCase['sheets'][number]>; more?: LinksCase['sheets'] }

export const CELLS = Object.fromEntries([1, 2, 3, 4].flatMap((r, i) => [[`B${r}`, { value: i + 1 }], [`C${r}`, { value: [2, 3, 5, 7][i] }]]))
export const REFUSED = { refused: true } as const
export const SHEET1 = 'xl/worksheets/sheet1.xml'
export const STYLES = 'xl/styles.xml'
const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const XM = 'http://schemas.microsoft.com/office/excel/2006/main'

export const xml = (t: string) => t.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;')
export const attr = (t: string) => xml(t).replaceAll('"', '&quot;')
export const rep = (s: string, n: number) => s.repeat(Math.ceil(n / s.length)).slice(0, n)
export const list = (n: number, f: (i: number) => string, sep = '') => Array.from({ length: n }, (_, i) => f(i)).join(sep)

/** Changes parts by text; a part's text that does not hold what it replaces throws, so no case goes out unchanged. */
export function subs(...changes: [part: string, from: string | RegExp, to: string][]): Edit {
  return (part, text) => {
    let x = text
    for (const [p, from, to] of changes) {
      if (p !== part) continue
      if (typeof from === 'string' ? !x.includes(from) : !from.test(x)) throw new Error(`file-refusals: ${part} holds no ${from}`)
      x = x.replace(from, () => to)
    }
    return x === text ? undefined : x
  }
}

export function one(id: string, about: string, linked: Record<string, string>, e: Extra = {}): LinksCase {
  const { value, cells, sheet, more, ...rest } = e
  const first = Object.keys(linked)[0] ?? 'A1'
  return {
    id,
    about,
    sheets: [{ name: 'S', cells: { ...CELLS, ...cells }, ...sheet }, ...(more ?? [])],
    linked: Object.keys(linked).length ? { S: linked } : {},
    links: [],
    read: [{ sheet: 'S', range: first }],
    ...(value === undefined ? {} : { predicted: { [`S!${first}`]: value } }),
    ...rest,
  }
}

/** A formula in S!A1. */
export const cell = (id: string, about: string, formula: string, e: Extra = {}) => one(id, about, { A1: formula }, e)

// --- B1: arguments that only a reference may fill (MS-XLSX 2.2.2, `ref-argument-expression`)

/** Every position of the grammar's table, `@` where the reference goes. */
const POSITIONS: Record<string, string> = {
  areas: 'AREAS(@)',
  countblank: 'COUNTBLANK(@)',
  formulatext: '_xlfn.FORMULATEXT(@)',
  isformula: '_xlfn.ISFORMULA(@)',
  offset: 'OFFSET(@,0,0)',
  phonetic: 'PHONETIC(@)',
  column: 'COLUMN(@)',
  row: 'ROW(@)',
  countif: 'COUNTIF(@,1)',
  'countifs-1': 'COUNTIFS(@,1)',
  'countifs-3': 'COUNTIFS(B1:B3,1,@,1)',
  'sumif-1': 'SUMIF(@,1)',
  'sumif-3': 'SUMIF(B1:B3,1,@)',
  'averageif-1': 'AVERAGEIF(@,1)',
  'averageif-3': 'AVERAGEIF(B1:B3,1,@)',
  'sumifs-1': 'SUMIFS(@,B1:B3,1)',
  'sumifs-2': 'SUMIFS(B1:B3,@,1)',
  'sumifs-4': 'SUMIFS(B1:B3,C1:C3,1,@,1)',
  'averageifs-1': 'AVERAGEIFS(@,B1:B3,1)',
  'averageifs-2': 'AVERAGEIFS(B1:B3,@,1)',
  'averageifs-4': 'AVERAGEIFS(B1:B3,C1:C3,1,@,1)',
  'maxifs-1': '_xlfn.MAXIFS(@,B1:B3,1)',
  'maxifs-2': '_xlfn.MAXIFS(B1:B3,@,1)',
  'maxifs-4': '_xlfn.MAXIFS(B1:B3,C1:C3,1,@,1)',
  'minifs-1': '_xlfn.MINIFS(@,B1:B3,1)',
  'minifs-2': '_xlfn.MINIFS(B1:B3,@,1)',
  'minifs-4': '_xlfn.MINIFS(B1:B3,C1:C3,1,@,1)',
  ...Object.fromEntries(['DAVERAGE', 'DCOUNT', 'DCOUNTA', 'DGET', 'DMAX', 'DMIN', 'DPRODUCT', 'DSTDEV', 'DSTDEVP', 'DSUM', 'DVAR', 'DVARP'].map((d) => [d.toLowerCase(), `${d}(@,"a",C1:C2)`])),
  rank: 'RANK(1,@)',
  'rank-avg': '_xlfn.RANK.AVG(1,@)',
  'rank-eq': '_xlfn.RANK.EQ(1,@)',
  'subtotal-2': 'SUBTOTAL(9,@)',
  'subtotal-3': 'SUBTOTAL(9,B1:B3,@)',
  'aggregate-5': '_xlfn.AGGREGATE(9,0,B1:B3,C1:C3,@)',
  cell: 'CELL("row",@)',
}

const at = (position: string, arg: string) => POSITIONS[position].replace('@', () => arg)

/** A value where a reference goes: refused at every position. */
const VALUES: Record<string, string> = { num: '1', arr: '{1,2,3}', expr: 'B1:B3*1' }

/** One position of each function: a call in a reference's place reads alike across the positions of the table, which these sample. */
const SAMPLE = ['areas', 'countblank', 'formulatext', 'offset', 'phonetic', 'row', 'countif', 'countifs-3', 'sumif-3', 'sumifs-2', 'maxifs-1', 'dsum', 'rank', 'rank-eq', 'subtotal-2', 'aggregate-5', 'cell']

/** Calls in a reference's place; `refused` the old functions that return a value. */
const CALLS: [string, string, boolean][] = [
  ['sum', 'SUM(B1:B3)', true],
  ['iferror', 'IFERROR(B1:B3,0)', true],
  ['if', 'IF(TRUE,B1:B3)', false],
  ['index', 'INDEX(B1:B3,1)', false],
  ['switch', '_xlfn.SWITCH(1,1,B1:B3)', false],
  ['ifs', '_xlfn.IFS(TRUE,B1:B3)', false],
  ['let', '_xlfn.LET(_xlpm.x,B1:B3,_xlpm.x)', false],
  ['xlookup', '_xlfn.XLOOKUP(1,B1:B3,C1:C3)', false],
  ['unique', '_xlfn.UNIQUE(B1:B3)', false],
  ['filter', '_xlfn._xlws.FILTER(B1:B3,C1:C3>0)', false],
  ['sort', '_xlfn._xlws.SORT(B1:B3)', false],
  ['take', '_xlfn.TAKE(B1:B3,1)', false],
  ['sequence', '_xlfn.SEQUENCE(3)', false],
  ['udf', 'FOO(B1:B3)', false],
]

const NM = { names: [{ name: 'nm', ref: 'S!$B$1:$B$3' }] }

const referenceOnly = (): LinksCase[] => [
  ...Object.keys(POSITIONS).flatMap((p) => Object.entries(VALUES).map(([k, v]) => cell(`ref-${p}-${k}`, `${POSITIONS[p]} with ${v} for @`, at(p, v), REFUSED))),
  ...CALLS.flatMap(([k, v, refused]) => SAMPLE.map((p) => cell(`call-${p}-${k}`, `${POSITIONS[p]} with ${v} for @`, at(p, v), refused ? REFUSED : {}))),
  cell('ok-paren', 'a reference in parentheses', 'COUNTIF((B1:B3),1)', { value: 1 }),
  cell('ok-ref-error', 'the error #REF!', 'COUNTIF(#REF!,1)', { value: { error: 'REF' } }),
  cell('ok-union', 'a union', 'COUNTIF((B1,B2),1)', { value: { error: 'VALUE' } }),
  cell('ok-intersect', 'an intersection', 'COUNTIF(B1:B3 B2:B3,1)', { value: 0 }),
  cell('ok-range-op', 'the range operator over a call', 'COUNTIF(B1:INDEX(B1:B3,2),1)', { value: 1 }),
  cell('ok-name', 'a defined name', 'COUNTIF(nm,1)', { ...NM, value: 1 }),
  cell('ok-sheet', 'a reference with its sheet', 'COUNTIF(S!B1:B3,1)', { value: 1 }),
  cell('ok-choose', 'CHOOSE', 'COUNTIF(CHOOSE(1,B1:B3),1)', { value: 1 }),
  cell('ok-indirect', 'INDIRECT', 'COUNTIF(INDIRECT("B1:B3"),1)', { value: 1 }),
  cell('ok-offset', 'OFFSET', 'COUNTIF(OFFSET(B1,0,0,3),1)', { value: 1 }),
  cell('ok-anchor', 'a spill reference', 'COUNTIF(_xlfn.ANCHORARRAY(D1),1)'),
  cell('ok-single', 'the implicit intersection operator', 'ROW(_xlfn.SINGLE(B1:B3))', { value: 1 }),
  cell('left-offset', 'OFFSET with its reference left out', 'OFFSET(,1,0)', REFUSED),
  cell('left-cell', 'CELL with its reference left out', 'CELL("row",)', REFUSED),
  cell('left-dsum', 'DSUM with its database left out', 'DSUM(,1,C1:C2)', REFUSED),
  cell('left-aggregate', 'AGGREGATE with its fifth argument left out', '_xlfn.AGGREGATE(14,6,B1:B3,1,)', REFUSED),
]

// --- B2: a text literal over 255 characters

const LEN = (t: string) => `LEN("${t}")`
const inRule = (t: string) => `LEN("${t}")>0`

/** The dxf a rule of the 2007 block points at. */
export const DXF = [STYLES, '</styleSheet>', '<dxfs count="1"><dxf><font><b/></font></dxf></dxfs></styleSheet>'] as [string, string, string]
export const inSheet = (x: string): [string, string, string] => [SHEET1, '</sheetData>', `</sheetData>${x}`]
const cfBlock = (rule: string, sqref = 'D1') => `<conditionalFormatting sqref="${sqref}">${rule}</conditionalFormatting>`
/** A rule of the 2007 block with the formula as the file holds it. */
export const cfEdit = (formula: string) => subs(inSheet(cfBlock(`<cfRule type="expression" dxfId="0" priority="1"><formula>${xml(formula)}</formula></cfRule>`)), DXF)
/** A rule of Excel 2010's block, where a rule over another sheet goes. */
const x14Edit = (formula: string) =>
  subs(inSheet(`<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${X14}"><x14:conditionalFormattings><x14:conditionalFormatting xmlns:xm="${XM}"><x14:cfRule type="expression" priority="1" id="{00000001-0000-4000-8000-000000000001}"><xm:f>${xml(formula)}</xm:f><x14:dxf><font><b/></font></x14:dxf></x14:cfRule><xm:sqref>D1</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>`))
const dvRule = (type: string, formula: string, sqref = 'D1') => `<dataValidation type="${type}" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="${sqref}"><formula1>${xml(formula)}</formula1></dataValidation>`
export const dvEdit = (type: string, formula: string) => subs(inSheet(`<dataValidations count="1">${dvRule(type, formula)}</dataValidations>`))
const csv = (n: number) => rep('abc,', n)

const literals = (): LinksCase[] => [
  cell('text-255', '255 characters', LEN(rep('a', 255)), { value: 255 }),
  cell('text-256', '256 characters', LEN(rep('a', 256)), REFUSED),
  cell('text-300-concat', 'the audit: 300 characters joined to another', `"${rep('a', 300)}"&"b"`, REFUSED),
  cell('text-two-255', 'two literals of 255 joined', `LEN("${rep('a', 255)}"&"${rep('b', 255)}")`, { value: 510 }),
  cell('text-quote-255', '254 characters and a doubled quote: 255 once undoubled, 256 in the file', LEN(rep('a', 254) + '""'), { value: 255 }),
  cell('text-quote-256', '255 characters and a doubled quote: 256 once undoubled', LEN(rep('a', 255) + '""'), REFUSED),
  cell('text-quotes-128', '128 doubled quotes: 256 in the file, 128 once undoubled', LEN(rep('""', 256)), { value: 128 }),
  cell('text-cyrillic-255', '255 Cyrillic characters', LEN(rep('я', 255)), { value: 255 }),
  cell('text-cyrillic-256', '256 Cyrillic characters', LEN(rep('я', 256)), REFUSED),
  cell('text-astral-255', '253 characters and an astral one: 254 code points, 255 UTF-16 units', LEN(rep('a', 253) + '😀'), { value: 255 }),
  cell('text-astral-256', '254 characters and an astral one: 255 code points, 256 UTF-16 units', LEN(rep('a', 254) + '😀'), REFUSED),
  cell('text-escaped-255', '255 ampersands, 1,275 bytes of XML', LEN(rep('&', 255)), { value: 255 }),
  one('dv-list-255', 'a list rule of 255 characters', {}, { edit: dvEdit('list', `"${csv(255)}"`), holds: 'type="list"' }),
  one('dv-list-256', 'a list rule of 256 characters', {}, { edit: dvEdit('list', `"${csv(256)}"`), ...REFUSED }),
  one('dv-list-400', 'the audit: a list rule of 400 characters', {}, { edit: dvEdit('list', `"${csv(400)}"`), ...REFUSED }),
  one('dv-custom-255', 'a literal of 255 in a custom rule', {}, { edit: dvEdit('custom', inRule(rep('a', 255))), holds: 'type="custom"' }),
  one('dv-custom-256', 'a literal of 256 in a custom rule', {}, { edit: dvEdit('custom', inRule(rep('a', 256))), ...REFUSED }),
  one('dv-custom-long', 'a custom rule of 301 characters with no long literal, over what the dialog takes', {}, { edit: dvEdit('custom', `${list(100, () => 'B1', '+')}>0`), holds: 'type="custom"' }),
  one('cf-text-255', 'a literal of 255 in a rule', {}, { edit: cfEdit(inRule(rep('a', 255))), holds: 'type="expression"' }),
  one('cf-text-256', 'a literal of 256 in a rule', {}, { edit: cfEdit(inRule(rep('a', 256))), ...REFUSED }),
  one('cf-long', 'a rule of 301 characters with no long literal', {}, { edit: cfEdit(`${list(100, () => 'B1', '+')}>0`), holds: 'type="expression"' }),
  cell('name-text-255', 'a name holding a literal of 255', 'lit', { names: [{ name: 'lit', ref: `"${rep('a', 255)}"` }], value: rep('a', 255) }),
  cell('name-text-256', 'a name holding a literal of 256', 'lit', { names: [{ name: 'lit', ref: `"${rep('a', 256)}"` }], ...REFUSED }),
]

// --- B3: argument counts, and how long and deep a formula goes

const ones = (n: number) => list(n, () => '1', ',')
const plus = (n: number) => list(n, () => '1', '+')
const nested = (n: number) => `${'ABS('.repeat(n)}1${')'.repeat(n)}`
const parens = (n: number) => `${'('.repeat(n)}1${')'.repeat(n)}`
const lambda = (n: number) => `_xlfn.LAMBDA(${list(n, (i) => `_xlpm.p${i}`, ',')},_xlpm.p0)`

const counts = (): LinksCase[] => [
  cell('count-abs-2', 'ABS of two', 'ABS(1,2)', REFUSED),
  cell('count-abs-0', 'ABS of none', 'ABS()', REFUSED),
  cell('count-sum-0', 'SUM of none', 'SUM()', REFUSED),
  cell('count-if-0', 'IF of none', 'IF()', REFUSED),
  cell('count-if-1', 'IF of one: the grammar takes it', 'IF(TRUE)', REFUSED),
  cell('count-if-4', 'IF of four: the grammar takes it', 'IF(TRUE,1,2,3)', REFUSED),
  cell('count-len-2', 'LEN of two', 'LEN("a","b")', REFUSED),
  cell('count-vlookup-2', 'VLOOKUP of two', 'VLOOKUP(1,B1:C3)', REFUSED),
  cell('count-hlookup-2', 'HLOOKUP of two', 'HLOOKUP(1,B1:C3)', REFUSED),
  cell('count-index-0', 'INDEX of none', 'INDEX()', REFUSED),
  cell('count-match-1', 'MATCH of one', 'MATCH(1)', REFUSED),
  cell('count-sumif-1', 'SUMIF of one', 'SUMIF(B1:B3)', REFUSED),
  cell('count-countif-1', 'COUNTIF of one', 'COUNTIF(B1:B3)', REFUSED),
  cell('count-offset-1', 'OFFSET of one', 'OFFSET(B1)', REFUSED),
  cell('count-indirect-0', 'INDIRECT of none', 'INDIRECT()', REFUSED),
  cell('count-xlookup-1', 'XLOOKUP of one', '_xlfn.XLOOKUP(1)', REFUSED),
  cell('count-xlookup-7', 'XLOOKUP of seven', '_xlfn.XLOOKUP(1,B1:B3,C1:C3,0,0,1,1)', REFUSED),
  cell('count-byrow-1', 'BYROW of one: the grammar takes it', '_xlfn.BYROW(B1:B3)'),
  cell('count-bycol-1', 'BYCOL of one: the grammar takes it', '_xlfn.BYCOL(B1:B3)'),
  cell('count-byrow-2', 'BYROW of two', '_xlfn.BYROW(B1:B3,_xlfn.LAMBDA(_xlpm.x,SUM(_xlpm.x)))'),
  cell('count-webservice-2', 'WEBSERVICE of two: the grammar takes it', '_xlfn.WEBSERVICE("a","b")', REFUSED),
  cell('count-webservice-1', 'WEBSERVICE of one', '_xlfn.WEBSERVICE("a")'),
  cell('count-filter-4', 'FILTER of four: the grammar takes it', '_xlfn._xlws.FILTER(B1:B3,C1:C3>0,0,0)', REFUSED),
  cell('count-sum-255', 'SUM of 255', `SUM(${ones(255)})`, { value: 255 }),
  cell('count-sum-256', 'SUM of 256', `SUM(${ones(256)})`, REFUSED),
  cell('count-sum-300', 'the audit: SUM of 300', `SUM(${ones(300)})`, REFUSED),
  cell('count-udf-255', 'a user-defined call of 255', `FOO(${ones(255)})`, { value: { error: 'NAME' } }),
  cell('count-udf-256', 'a user-defined call of 256', `FOO(${ones(256)})`, REFUSED),
  cell('count-lambda-253', 'LAMBDA of 253 parameters', `${lambda(253)}(${ones(253)})`, { value: 1 }),
  cell('count-lambda-254', 'LAMBDA of 254 parameters', `${lambda(254)}(${ones(254)})`, REFUSED),
  // hand-made, whatever the engine's parser takes: how long a chain Excel itself loads
  cell('plus-1000', '1,000 terms', plus(1000), { value: 1000 }),
  cell('plus-1025', '1,025 terms, 1,024 operators', plus(1025), { value: 1025 }),
  cell('plus-1026', '1,026 terms, 1,025 operators', plus(1026), { value: 1026 }),
  cell('plus-2000', '2,000 terms', plus(2000), { value: 2000 }),
  cell('plus-4096', '4,096 terms, 8,191 characters, 16,383 bytes of tokens', plus(4096), { value: 4096 }),
  cell('plus-4097', '4,097 terms, 8,193 characters, 16,387 bytes of tokens', plus(4097), REFUSED),
  cell('plus-4200', '4,200 terms', plus(4200), REFUSED),
  cell('plus-4600', '4,600 terms, 9,199 characters', plus(4600), REFUSED),
  cell('spaces-8195', '8,195 characters of which 8,192 spaces, few tokens', `1${rep(' ', 8192)}+1`, { value: 2 }),
  cell('grow-xlfn', 'N13: 350 XLOOKUP, 8,049 characters as typed, 10,149 with the prefix', list(350, () => '_xlfn.XLOOKUP(1,B1:B2,C1:C2)', '+')),
  one('grow-table', 'N13: 1,100 [@Qty], 7,699 characters as typed, about 27,500 as a file holds them', { F2: list(1100, () => 'Sales[[#This Row],[Qty]]', '+') }, {
    cells: { E1: { value: 'Qty' }, F1: { value: 'F' }, E2: { value: 1 } },
    sheet: { tables: [{ name: 'Sales', range: 'E1:F2', columns: [{ name: 'Qty' }, { name: 'F' }] }] },
    value: 1100,
  }),
  cell('nest-64', '64 calls one in another', nested(64), { value: 1 }),
  cell('nest-65', '65 calls one in another', nested(65), REFUSED),
  cell('nest-70', 'the audit: 70 calls one in another', nested(70), REFUSED),
  cell('parens-64', '64 parentheses', parens(64), { value: 1 }),
  cell('parens-65', '65 parentheses', parens(65), { value: 1 }),
  cell('parens-300', '300 parentheses', parens(300), { value: 1 }),
]

// --- B4: names

const named = (id: string, about: string, name: string, ref: string, e: Extra = {}) => cell(id, about, 'SUM(B1:B2)', { names: [{ name, ref }], ...e, holds: `name="${attr(name)}"` })

const names = (): LinksCase[] => [
  cell('name-rel', 'N1: a relative reference without its sheet, used from A1', 'rel', { names: [{ name: 'rel', ref: 'C1' }], ...REFUSED }),
  cell('name-abs', 'N1: an absolute reference without its sheet', 'abs', { names: [{ name: 'abs', ref: '$C$2' }], ...REFUSED }),
  cell('name-calc', 'N1: a calculation over a reference without its sheet', 'calc', { names: [{ name: 'calc', ref: 'C1*2' }], ...REFUSED }),
  cell('name-sum', 'N1: a function of a range without its sheet', 'tb', { names: [{ name: 'tb', ref: 'SUM(C1:C2)' }], ...REFUSED }),
  cell('name-const', 'a calculation over constants', 'k', { names: [{ name: 'k', ref: '1+1' }], value: 2 }),
  cell('name-sheet', 'a reference with its sheet', 'ok', { names: [{ name: 'ok', ref: 'S!$C$2' }], value: 3 }),
  named('name-xl-foo', 'N17: a name that starts _xl', '_xlfoo', 'S!$A$1'),
  named('name-xlnm-foo', 'N17: a built-in prefix with no such name', '_xlnm.Foo', 'S!$A$1'),
  named('name-xlcalc', 'N17: the prefix of the names writeKept carries', '_xlcalc.keep.zz', 'S!$A$1'),
  named('name-backslash-a', 'N17: backslash and one letter', '\\a', 'S!$A$1', REFUSED),
  named('name-backslash', 'N17: a backslash alone', '\\', 'S!$A$1'),
  named('name-backslash-cyr', 'N17: backslash and a word, as `validateName` takes', '\\Ставка', 'S!$A$1'),
]

// --- B5: conditional formats, validation, a table's totals

const TABLE = { cells: { E1: { value: 'Qty' }, E2: { value: 1 }, E3: { value: 2 } }, sheet: { tables: [{ name: 'Sales', range: 'E1:E3', columns: [{ name: 'Qty' }] }] } }
const T: Extra = { more: [{ name: 'T', cells: { B1: { value: 1 } } }] }
const LINK: LinkedBook = { target: 'B.xlsx', sheets: [{ name: 'A', cells: { A1: 5 } }] }

const CF_HOLDS = { holds: 'expression' }

const rules = (): LinksCase[] => [
  one('cf-plain', 'a rule of the first form', {}, { edit: cfEdit('B1>0'), ...CF_HOLDS }),
  one('cf-array', 'N4: an array constant', {}, { edit: cfEdit('OR(B1={1,2,3})'), ...REFUSED }),
  one('cf-structured', 'N4: a table column', {}, { ...TABLE, edit: cfEdit('B1>SUM(Sales[Qty])'), ...REFUSED }),
  one('cf-union', 'N4: a union', {}, { edit: cfEdit('SUM((B1,B2))>0'), ...REFUSED }),
  one('cf-intersect', 'N4: an intersection', {}, { edit: cfEdit('SUM(B1:B3 B2:B3)>0'), ...REFUSED }),
  one('cf-3d', 'N4: a reference across sheets', {}, { ...T, edit: cfEdit("SUM('S:T'!B1)>0"), ...REFUSED }),
  one('cf-lambda', 'the grammar: a LAMBDA', {}, { edit: cfEdit('_xlfn.LAMBDA(_xlpm.x,_xlpm.x)(1)>0'), ...REFUSED }),
  one('cf-let', 'a LET, not in the grammar’s list', {}, { edit: cfEdit('_xlfn.LET(_xlpm.x,1,_xlpm.x)>0'), ...CF_HOLDS }),
  one('cf-range-op', 'the grammar: the range operator', {}, { edit: cfEdit('SUM(B1:INDEX(B1:B3,2))>0'), ...CF_HOLDS }),
  one('cf-own-sheet', 'a reference naming its own sheet', {}, { edit: cfEdit('S!B1>0'), ...CF_HOLDS }),
  one('cf-other-sheet', 'another sheet in the first form, which Excel 2007 did not take', {}, { ...T, edit: cfEdit('T!B1>0'), ...CF_HOLDS }),
  one('cf-external', 'a link to another workbook', {}, { links: [LINK], edit: cfEdit('[1]A!A1>0'), ...REFUSED }),
  one('x14-plain', 'Excel 2010’s block', {}, { ...T, edit: x14Edit('T!B1>0'), holds: 'x14:cfRule' }),
  one('x14-array', 'N4: an array constant in Excel 2010’s block', {}, { ...T, edit: x14Edit('OR(T!B1={1,2,3})'), ...REFUSED }),
  one('x14-structured', 'N4: a table column in Excel 2010’s block', {}, { ...TABLE, ...T, edit: x14Edit('T!B1>SUM(Sales[Qty])'), ...REFUSED }),
  one('x14-union', 'N4: a union in Excel 2010’s block', {}, { ...T, edit: x14Edit('SUM((T!B1,T!B2))>0'), ...REFUSED }),
  one('x14-intersect', 'N4: an intersection in Excel 2010’s block', {}, { ...T, edit: x14Edit('SUM(T!B1:B3 T!B2:B3)>0'), ...REFUSED }),
  one('x14-3d', 'N4: a reference across sheets in Excel 2010’s block', {}, { ...T, edit: x14Edit("SUM('S:T'!B1)>0"), ...REFUSED }),
  one('cfvo-array', 'an array constant as a threshold', {}, { edit: subs(inSheet(cfBlock('<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="formula" val="{1,2}"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule>'))), ...REFUSED }),
  one('dv-plain', 'a custom rule', {}, { edit: dvEdit('custom', 'B1>0'), holds: 'type="custom"' }),
  one('dv-array', 'N4: an array constant', {}, { edit: dvEdit('custom', 'OR(B1={1,2,3})'), ...REFUSED }),
  one('dv-list-array', 'N4: an array constant as a list', {}, { edit: dvEdit('list', '{1,2,3}'), ...REFUSED }),
  one('dv-structured', 'N4: a table column as a list', {}, { ...TABLE, edit: dvEdit('list', 'Sales[Qty]'), ...REFUSED }),
  one('dv-union', 'N4: a union', {}, { edit: dvEdit('custom', 'SUM((B1,B2))>0'), ...REFUSED }),
  one('dv-intersect', 'N4: an intersection', {}, { edit: dvEdit('custom', 'SUM(B1:B3 B2:B3)>0'), ...REFUSED }),
  one('dv-3d', 'N4: a reference across sheets', {}, { ...T, edit: dvEdit('custom', "SUM('S:T'!B1)>0"), ...REFUSED }),
  one('dv-lambda', 'the grammar: a LAMBDA', {}, { edit: dvEdit('custom', '_xlfn.LAMBDA(_xlpm.x,_xlpm.x)(1)>0'), ...REFUSED }),
  one('dv-list-name', 'a defined name as a list', {}, { ...NM, edit: dvEdit('list', 'nm'), holds: 'type="list"' }),
  one('dv-list-other', 'another sheet’s range as a list', {}, { ...T, edit: dvEdit('list', 'T!$B$1:$B$3'), holds: 'type="list"' }),
  one('dv-list-indirect', 'a table column through INDIRECT', {}, { ...TABLE, edit: dvEdit('list', 'INDIRECT("Sales[Qty]")'), holds: 'type="list"' }),
  one('dv-external', 'a link to another workbook, which the grammar leaves to rules of validation', {}, { links: [LINK], edit: dvEdit('custom', '[1]A!A1>0'), holds: 'type="custom"' }),
  ...totals(),
]

/** A table over E1:E3 whose totals row is E3: its column's attributes in the table part, what the cell holds beside them. */
export function total(id: string, about: string, column: string, foot: { formula?: string; value?: number }, e: Extra = {}): LinksCase {
  return one(id, about, foot.formula ? { E3: foot.formula } : {}, {
    cells: { E1: { value: 'Q' }, E2: { value: 1 }, ...(foot.value === undefined ? {} : { E3: { value: foot.value } }) },
    sheet: { tables: [{ name: 'Tbl', range: 'E1:E3', totals: true, columns: [{ name: 'Q' }] }] },
    read: [{ sheet: 'S', range: 'E3' }],
    edit: subs(['xl/tables/table1.xml', ANY_COLUMN, `<tableColumn id="1" name="Q"${column}/></tableColumns>`]),
    ...e,
  })
}

/** The column as the writer wrote it, whatever it made of the cell under it: the case's own goes in by hand. */
export const ANY_COLUMN = /<tableColumn id="1" name="Q".*<\/tableColumns>/

const SUM_ATTR = ' totalsRowFunction="sum"'

const totals = (): LinksCase[] => [
  total('totals-sum-subtotal', 'N5: the function and its SUBTOTAL', SUM_ATTR, { formula: 'SUBTOTAL(109,Tbl[Q])' }, { value: 1 }),
  total('totals-sum-other', 'N5: sum over A2*2', SUM_ATTR, { formula: 'E2*2' }, REFUSED),
  total('totals-sum-sum', 'N5: sum over SUM(E2:E2)', SUM_ATTR, { formula: 'SUM(E2:E2)' }, REFUSED),
  total('totals-sum-value', 'N5: sum over a value', SUM_ATTR, { value: 42 }, { value: 42 }),
  total('totals-sum-empty', 'N5: sum over an empty cell, which a host makes by setting the column’s totals alone', SUM_ATTR, {}),
  total('totals-custom-empty', 'N5: custom with no formula and an empty cell', ' totalsRowFunction="custom"', {}),
  total('totals-custom-value', 'N5: custom with no formula and a value', ' totalsRowFunction="custom"', { value: 42 }, { value: 42 }),
  total('totals-none-value', 'no function over a value', '', { value: 42 }, { value: 42 }),
]

// --- B6: how many and how long

export const onStyle = (i: number): [string, string, string] => [SHEET1, '<c r="A1"><v>5</v></c>', `<c r="A1" s="${i}"><v>5</v></c>`]
export const FIVE: Extra = { cells: { A1: { value: 5 } } }
const FONTS = /<fonts count="1">.*<\/fonts>/
const FILLS = /<fills count="2">.*<\/fills>/
const BORDERS = /<borders count="1">.*<\/borders>/
export const XFS = /<cellXfs count="1">.*<\/cellXfs>/
const BASE_FONT = '<font><sz val="11"/><name val="Calibri"/></font>'
export const BASE_XF = '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'

/** A font of the name on A1, or as the book's default. */
export const fontCell = (name: string) => subs([STYLES, FONTS, `<fonts count="2">${BASE_FONT}<font><sz val="11"/><name val="${attr(name)}"/></font></fonts>`], [STYLES, XFS, `<cellXfs count="2">${BASE_XF}<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>`], onStyle(1))
const fontBook = (name: string) => subs([STYLES, '<name val="Calibri"/>', `<name val="${attr(name)}"/>`])

/** `n` custom number formats, A1 in the last: the shown text says whether Excel kept it. */
export const numFmts = (n: number) =>
  subs([STYLES, '<fonts ', `<numFmts count="${n}">${list(n, (i) => `<numFmt numFmtId="${164 + i}" formatCode="0.00&quot;i${i + 1}&quot;"/>`)}</numFmts><fonts `], [STYLES, XFS, `<cellXfs count="2">${BASE_XF}<xf numFmtId="${164 + n - 1}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>`], onStyle(1))

/** `n` different cell formats out of 300 fonts and 300 fills, A1 in the last. */
function manyXfs(n: number): Edit {
  const font = (i: number) => `<font><sz val="${(6 + i / 20).toFixed(2)}"/><name val="Calibri"/></font>`
  const fill = (i: number) => (i < 2 ? (i ? '<fill><patternFill patternType="gray125"/></fill>' : '<fill><patternFill patternType="none"/></fill>') : `<fill><patternFill patternType="solid"><fgColor rgb="FF${i.toString(16).padStart(6, '0').toUpperCase()}"/></patternFill></fill>`)
  return subs(
    [STYLES, FONTS, `<fonts count="300">${BASE_FONT}${list(299, (i) => font(i + 1))}</fonts>`],
    [STYLES, FILLS, `<fills count="300">${list(300, fill)}</fills>`],
    [STYLES, XFS, `<cellXfs count="${n}">${BASE_XF}${list(n - 1, (i) => `<xf numFmtId="0" fontId="${(i + 1) % 300}" fillId="${Math.floor((i + 1) / 300)}" borderId="0" xfId="0" applyFont="1" applyFill="1"/>`)}</cellXfs>`],
    onStyle(n - 1),
  )
}
export const manyFonts = (n: number) => subs([STYLES, FONTS, `<fonts count="${n}">${list(n, (i) => `<font><sz val="${(6 + (i % 400) / 20).toFixed(2)}"/><name val="Calibri"/></font>`)}</fonts>`])
const manyFills = (n: number) => subs([STYLES, FILLS, `<fills count="${n}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${list(n - 2, (i) => `<fill><patternFill patternType="solid"><fgColor rgb="FF${i.toString(16).padStart(6, '0').toUpperCase()}"/></patternFill></fill>`)}</fills>`])
const manyBorders = (n: number) => subs([STYLES, BORDERS, `<borders count="${n}">${list(n, () => '<border><left/><right/><top/><bottom/><diagonal/></border>')}</borders>`])

const cellsOdd = (n: number) => list(n, (i) => `A${2 * i + 10}`, ' ')
const onLink = (location: string) => subs([SHEET1, 'location="S!A1"', `location="${attr(location)}"`])
const onUrl = (url: string) => subs(['xl/worksheets/_rels/sheet1.xml.rels', 'Target="https://e.com/?"', `Target="${attr(url)}"`])
const url = (n: number) => `https://e.com/?${rep('q', n - 15)}`
const AUTHOR = (n: number) => subs(['xl/comments1.xml', '<author>x</author>', `<author>${rep('a', n)}</author>`])
export const LABEL = (n: number) => subs(['xl/tables/table1.xml', ANY_COLUMN, `<tableColumn id="1" name="Q" totalsRowLabel="${rep('L', n)}"/></tableColumns>`])
export const labelled = (id: string, about: string, n: number, e: Extra = {}) => one(id, about, {}, { cells: { E1: { value: 'Q' }, E2: { value: 1 }, E3: { value: 'L' } }, sheet: { tables: [{ name: 'Tbl', range: 'E1:E3', totals: true, columns: [{ name: 'Q', totals: { label: 'L' } }] }] }, edit: LABEL(n), ...e })
const SHEET_T = (name: string) => subs(['xl/workbook.xml', '<sheet name="T"', `<sheet name="${attr(name)}"`])

/** A target of a link of `n` characters. */
const target = (n: number) => `https://sp.example.com/${rep('deep/', n - 29)}B.xlsx`
const linkedTo = (b: LinkedBook, e: Extra = {}) => ({ links: [b], ...e })
const extCell = (id: string, about: string, b: LinkedBook, e: Extra = {}) => cell(id, about, `[1]${b.sheets[0].name}!A1`, linkedTo(b, e))

/** The cache and the PivotTable's sheets of a case: the book of `test/pivot-books.ts`, its slicer left out where `slicer` is false. */
export function pivoted(id: string, about: string, e: Extra & { slicer?: boolean; change?: (b: ReturnType<typeof pivotBook>) => void }): LinksCase {
  const { slicer = true, change, ...rest } = e
  const whole = pivotBook()
  if (!slicer) {
    whole.slicerCaches = []
    whole.sheets[1].slicers = []
  }
  change?.(whole)
  const { sheets, ...book } = whole
  return { id, about, sheets, linked: {}, links: [], read: [{ sheet: 'Data', range: 'B2' }], predicted: { 'Data!B2': 10 }, book, ...rest }
}
export /** A slicer as the writer wrote one of no items before it took those of its field. */
const NO_ITEMS: [string, RegExp, string] = ['xl/slicerCaches/slicerCache1.xml', /<items count="\d+">.*<\/items>/, '<items count="0"></items>']
export const CACHE = 'xl/pivotCache/pivotCacheDefinition1.xml'
const PIVOT = 'xl/pivotTables/pivotTable1.xml'

const limits = (): LinksCase[] => [
  one('font-cell-31', 'N6: a cell’s font of 31 characters', {}, { ...FIVE, edit: fontCell(rep('F', 31)), holds: rep('F', 31) }),
  one('font-cell-32', 'N6: a cell’s font of 32 characters', {}, { ...FIVE, edit: fontCell(rep('F', 32)), ...REFUSED }),
  one('font-cell-40', 'N6: a cell’s font of 40 characters', {}, { ...FIVE, edit: fontCell(rep('F', 40)), ...REFUSED }),
  one('font-cell-empty', 'N6: a cell’s font with an empty name', {}, { ...FIVE, edit: fontCell(''), ...REFUSED }),
  one('font-book-31', 'N6: the default font of 31 characters', {}, { edit: fontBook(rep('F', 31)), holds: rep('F', 31) }),
  one('font-book-32', 'N6: the default font of 32 characters', {}, { edit: fontBook(rep('F', 32)), ...REFUSED }),
  one('font-book-empty', 'N6: the default font with an empty name', {}, { edit: fontBook(''), ...REFUSED }),
  // 6 workbooks of users on the corpus have more than 206, which Excel must have opened
  one('numfmt-206', 'N7: 206 number formats', {}, { ...FIVE, edit: numFmts(206), holds: 'i206' }),
  one('numfmt-207', 'N7: 207 number formats', {}, { ...FIVE, edit: numFmts(207), holds: 'i207' }),
  one('numfmt-300', 'N7: 300 number formats', {}, { ...FIVE, edit: numFmts(300), holds: 'i300' }),
  one('numfmt-1000', 'N7: 1,000 number formats', {}, { ...FIVE, edit: numFmts(1000), holds: 'i1000' }),
  one('xf-65430', 'N9: 65,430 cell formats', {}, { ...FIVE, edit: manyXfs(65430) }),
  one('xf-65431', 'N9: 65,431 cell formats', {}, { ...FIVE, edit: manyXfs(65431), ...REFUSED }),
  one('xf-65490', 'N9: 65,490 cell formats, what the dialog’s limit says', {}, { ...FIVE, edit: manyXfs(65490), ...REFUSED }),
  one('xf-65491', 'N9: 65,491 cell formats', {}, { ...FIVE, edit: manyXfs(65491), ...REFUSED }),
  one('xf-70000', 'N9: the audit: 70,000 cell formats', {}, { ...FIVE, edit: manyXfs(70000), ...REFUSED }),
  one('font-65491', 'N9: 65,491 fonts', {}, { edit: manyFonts(65491) }),
  one('font-65492', 'N9: 65,492 fonts', {}, { edit: manyFonts(65492), ...REFUSED }),
  one('fill-65430', 'N9: 65,430 fills', {}, { edit: manyFills(65430) }),
  one('fill-65431', 'N9: 65,431 fills', {}, { edit: manyFills(65431), ...REFUSED }),
  one('border-65430', 'N9: 65,430 borders', {}, { edit: manyBorders(65430) }),
  one('border-65431', 'N9: 65,431 borders', {}, { edit: manyBorders(65431), ...REFUSED }),
  one('link-location-2084', 'N10: a location of 2,084', {}, { cells: { D1: { value: 1, link: { location: 'S!A1' } } }, edit: onLink(`S!A1${rep('x', 2080)}`), holds: 'location="S!A1xxx' }),
  one('link-location-2085', 'N10: a location of 2,085', {}, { cells: { D1: { value: 1, link: { location: 'S!A1' } } }, edit: onLink(`S!A1${rep('x', 2081)}`), ...REFUSED }),
  one('link-url-2083', 'N10: an address of 2,083', {}, { cells: { D1: { value: 1, link: { url: 'https://e.com/?' } } }, edit: onUrl(url(2083)), holds: 'https://e.com/?qqq' }),
  one('link-url-2084', 'N10: an address of 2,084', {}, { cells: { D1: { value: 1, link: { url: 'https://e.com/?' } } }, edit: onUrl(url(2084)), holds: 'https://e.com/?qqq' }),
  one('link-url-2085', 'N10: an address of 2,085', {}, { cells: { D1: { value: 1, link: { url: 'https://e.com/?' } } }, edit: onUrl(url(2085)), ...REFUSED }),
  one('link-url-3000', 'N10: an address of 3,000, a signed URL', {}, { cells: { D1: { value: 1, link: { url: 'https://e.com/?' } } }, edit: onUrl(url(3000)), ...REFUSED }),
  one('cfvo-icon-min', 'N11: min on an icon set', {}, { edit: subs(inSheet(cfBlock('<cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows"><cfvo type="min"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule>'))), holds: 'iconSet' }),
  one('cfvo-icon-max', 'N11: max on an icon set', {}, { edit: subs(inSheet(cfBlock('<cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="max"/></iconSet></cfRule>'))), holds: 'iconSet' }),
  one('cfvo-scale-mid', 'N11: min at a colour scale’s midpoint', {}, { edit: subs(inSheet(cfBlock('<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FFFFFF00"/><color rgb="FF00FF00"/></colorScale></cfRule>'))), holds: 'colorScale' }),
  one('cfvo-bar-swap', 'N11: a bar from max to min', {}, { edit: subs(inSheet(cfBlock('<cfRule type="dataBar" priority="1"><dataBar><cfvo type="max"/><cfvo type="min"/><color rgb="FF638EC6"/></dataBar></cfRule>'))), holds: 'dataBar' }),
  one('count-cf-8192', 'N12: a rule over 8,192 ranges', {}, { edit: subs(inSheet(cfBlock('<cfRule type="duplicateValues" dxfId="0" priority="1"/>', cellsOdd(8192))), DXF), holds: 'duplicateValues' }),
  one('count-cf-8193', 'N12: a rule over 8,193 ranges', {}, { edit: subs(inSheet(cfBlock('<cfRule type="duplicateValues" dxfId="0" priority="1"/>', cellsOdd(8193))), DXF), ...REFUSED }),
  one('count-dv-8193', 'N12: validation over 8,193 ranges', {}, { edit: subs(inSheet(`<dataValidations count="1">${dvRule('custom', 'TRUE', cellsOdd(8193))}</dataValidations>`)), holds: 'type="custom"' }),
  one('count-dv-32767', 'N12: validation over 32,767 ranges', {}, { edit: subs(inSheet(`<dataValidations count="1">${dvRule('custom', 'TRUE', cellsOdd(32767))}</dataValidations>`)), holds: 'type="custom"' }),
  one('count-dv-32768', 'N12: validation over 32,768 ranges', {}, { edit: subs(inSheet(`<dataValidations count="1">${dvRule('custom', 'TRUE', cellsOdd(32768))}</dataValidations>`)), ...REFUSED }),
  one('count-dv-65534', 'N12: 65,534 validation rules', {}, { edit: subs(inSheet(`<dataValidations count="65534">${list(65534, (i) => dvRule('custom', 'TRUE', `A${i + 10}`))}</dataValidations>`)), holds: 'type="custom"' }),
  one('count-dv-65535', 'N12: 65,535 validation rules', {}, { edit: subs(inSheet(`<dataValidations count="65535">${list(65535, (i) => dvRule('custom', 'TRUE', `A${i + 10}`))}</dataValidations>`)), ...REFUSED }),
  one('count-select-8192', 'N12: a selection of 8,192', {}, { edit: subs([SHEET1, '<sheetView workbookViewId="0" tabSelected="1"/>', `<sheetView workbookViewId="0" tabSelected="1"><selection activeCell="A10" sqref="${cellsOdd(8192)}"/></sheetView>`]) }),
  one('count-select-8193', 'N12: a selection of 8,193', {}, { edit: subs([SHEET1, '<sheetView workbookViewId="0" tabSelected="1"/>', `<sheetView workbookViewId="0" tabSelected="1"><selection activeCell="A10" sqref="${cellsOdd(8193)}"/></sheetView>`]), ...REFUSED }),
  extCell('ext-target-218', 'N14: a target of 218', { ...LINK, target: target(218) }, { value: 5 }),
  extCell('ext-target-219', 'N14: a target of 219', { ...LINK, target: target(219) }, { value: 5 }),
  extCell('ext-target-260', 'N14: a target of 260', { ...LINK, target: target(260) }, { value: 5 }),
  extCell('ext-target-261', 'N14: a target of 261', { ...LINK, target: target(261) }, { value: 5 }),
  extCell('ext-target-330', 'N14: the audit: a SharePoint target of 330', { ...LINK, target: target(330) }, { value: 5 }),
  extCell('ext-sheet-31', 'N14: a linked sheet’s name of 31', { target: 'B.xlsx', sheets: [{ name: rep('n', 31), cells: { A1: 5 } }] }, { value: 5 }),
  extCell('ext-sheet-32', 'N14: a linked sheet’s name of 32', { target: 'B.xlsx', sheets: [{ name: rep('n', 32), cells: { A1: 5 } }] }, REFUSED),
  extCell('ext-name-bad', 'N14: a linked book’s name no book takes', { ...LINK, names: [{ name: 'x y', ref: '=A!$A$1' }] }, REFUSED),
  extCell('ext-name-ref', 'N14: a linked book’s name over no formula', { ...LINK, names: [{ name: 'xy', ref: '=garbage(' }] }, REFUSED),
  pivoted('pivot-control', 'N15: the PivotTable alone, no slicer', { slicer: false }),
  pivoted('pivot-slicer', 'N15: a slicer whose cache has no defined name', { edit: subs(NO_ITEMS), ...REFUSED }),
  pivoted('pivot-slicer-name', 'N15: a slicer with the defined name beside it', { names: [{ name: 'Slicer_Region', ref: '#N/A' }], edit: subs(NO_ITEMS), holds: 'Slicer_Region' }),
  pivoted('pivot-slicer-space', 'N15: a slicer cache whose name is no name', { edit: subs(['xl/slicerCaches/slicerCache1.xml', 'name="Slicer_Region"', 'name="Slicer Region"'], NO_ITEMS), ...REFUSED }),
  pivoted('pivot-slicer-taken', 'N15: a slicer cache whose name a defined name holds', { names: [{ name: 'Slicer_Region', ref: 'Data!$A$1' }], edit: subs(NO_ITEMS), ...REFUSED }),
  pivoted('pivot-error-spill', 'N16: #SPILL! among a cache’s items', { slicer: false, edit: subs([CACHE, '<s v="East"/>', '<e v="#SPILL!"/>']), ...REFUSED }),
  pivoted('pivot-error-na', 'N16: #N/A among a cache’s items', { slicer: false, edit: subs([CACHE, '<s v="East"/>', '<e v="#N/A"/>']) }),
  pivoted('pivot-field-dup', 'N16: two fields of a name in another case', { slicer: false, edit: subs([CACHE, 'cacheField name="Month"', 'cacheField name="region"']), ...REFUSED }),
  pivoted('pivot-field-empty', 'N16: a field of no name', { slicer: false, edit: subs([CACHE, 'cacheField name="Month"', 'cacheField name=""']), ...REFUSED }),
  pivoted('pivot-name-255', 'N16: a PivotTable’s name of 255', { slicer: false, edit: subs([PIVOT, 'name="PivotTable1"', `name="${rep('p', 255)}"`]) }),
  pivoted('pivot-name-256', 'N16: a PivotTable’s name of 256', { slicer: false, edit: subs([PIVOT, 'name="PivotTable1"', `name="${rep('p', 256)}"`]), ...REFUSED }),
  pivoted('pivot-caption-255', 'N16: a grand total’s caption of 255', { slicer: false, edit: subs([PIVOT, 'name="PivotTable1"', `name="PivotTable1" grandTotalCaption="${rep('g', 255)}"`]), holds: `grandTotalCaption="ggg` }),
  pivoted('pivot-caption-256', 'N16: a grand total’s caption of 256', { slicer: false, edit: subs([PIVOT, 'name="PivotTable1"', `name="PivotTable1" grandTotalCaption="${rep('g', 256)}"`]), ...REFUSED }),
  pivoted('pivot-source-long', 'N16: a source sheet’s name of 40', { slicer: false, edit: subs([CACHE, 'sheet="Data"', `sheet="${rep('s', 40)}"`]) }),
  one('note-author-65535', 'N18: a note’s author of 65,535', {}, { cells: { D3: { value: 1, note: { author: 'x', text: 't' } } }, edit: AUTHOR(65535), holds: '<author>aaa' }),
  one('note-author-65536', 'N18: a note’s author of 65,536', {}, { cells: { D3: { value: 1, note: { author: 'x', text: 't' } } }, edit: AUTHOR(65536), ...REFUSED }),
  labelled('label-10', 'N18: a label that is not the cell’s text', 10),
  labelled('label-32767', 'N18: a totals label of 32,767', 32767),
  labelled('label-32768', 'N18: a totals label of 32,768', 32768, REFUSED),
  one('sheet-31', 'R4: a sheet’s name of 31', {}, { ...T, edit: SHEET_T(rep('n', 31)) }),
  one('sheet-32', 'R4: a sheet’s name of 32', {}, { ...T, edit: SHEET_T(rep('n', 32)), ...REFUSED }),
  one('sheet-twice', 'R4: two sheets whose names differ in case alone', {}, { ...T, edit: SHEET_T('s'), ...REFUSED }),
  one('sheet-bracket', 'R4: a sheet’s name with a character Excel does not take', {}, { ...T, edit: SHEET_T('a[b'), ...REFUSED }),
]

export function buildLinksCases(): LinksCases {
  return { cases: [...referenceOnly(), ...literals(), ...counts(), ...names(), ...rules(), ...limits()] }
}
