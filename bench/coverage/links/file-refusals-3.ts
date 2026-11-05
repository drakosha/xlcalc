/**
 * The third round of `file-refusals`. The second padded a formula's text with spaces at its end, which Excel drops on load, so its
 * 8,193 characters said nothing: here the padding is a number's leading zeros, which a file holds. Its number formats in conditional
 * styles showed distinct formats count together, and 6 saved books of the corpus hold over 206 `numFmt` elements of few distinct
 * ids in their conditional styles, so those are told apart here. Slicers with their items, the first round's cases having none, and
 * the shapes left of a reference's place.
 *
 * `refused` and `predicted` were committed before Excel ran this set, with the Windows answers of both rounds in hand. The model:
 * a file's text is limited to 8,192 characters as it stands, and 206 distinct custom formats count across `numFmts` and conditional
 * styles, by id.
 */
import { MAX_FORMULA_LENGTH } from '../../../src/contract.js'
import type { LinksCase, LinksCases } from './cases.js'
import { BASE_XF, cell, cfEdit, dvEdit, FIVE, inSheet, list, one, onStyle, pivoted, REFUSED, rep, STYLES, subs, XFS } from './file-refusals.js'

/** `head` then `+` and a number 1 padded with leading zeros, `n` characters in all, `tail` at the end. */
function zeros(head: string, n: number, tail = ''): string {
  const pad = n - head.length - tail.length - 1
  if (pad < 1) throw new Error(`file-refusals-3: ${head.length} characters do not fit ${n}`)
  return `${head}+${rep('0', pad - 1)}1${tail}`
}
/** Calls joined by `+`, a number padding the text to `n` characters; `room` keeps a number's zeros short of Excel's limit on them. */
function chain(unit: string, n: number, room = 200): string {
  const count = Math.floor((n - room) / (unit.length + 1))
  return zeros(list(count, () => unit, '+'), n)
}

const XL = '_xlfn.XLOOKUP(1,B1:B2,C1:C2)'
const QTY = 'Sales[[#This Row],[Qty]]'
const SALES = { cells: { E1: { value: 'Qty' }, F1: { value: 'F' }, E2: { value: 1 } }, sheet: { tables: [{ name: 'Sales', range: 'E1:F2', columns: [{ name: 'Qty' }, { name: 'F' }] }] } }
const REFS = (n: number) => zeros(list(Math.floor((n - 200) / 3), () => 'B1', '+'), n, '>0')

const lengths = (): LinksCase[] => [
  ...[8192, 8193, 8200, 8500, 9000, 9500, 10000].map((n) => cell(`xlfn-${n}`, `${n} characters of XLOOKUP with its prefix, none of them spaces`, chain(XL, n), n > MAX_FORMULA_LENGTH ? REFUSED : {})),
  ...[8192, 8193, 9000, 12000, 16000, 20000, 27000].map((n) => one(`table-${n}`, `${n} characters of a table’s this-row references, none of them spaces`, { F2: chain(QTY, n) }, { ...SALES, ...(n > MAX_FORMULA_LENGTH ? REFUSED : {}) })),
  ...[8192, 8193].flatMap((n) => {
    const over = n > MAX_FORMULA_LENGTH ? REFUSED : {}
    return [
      cell(`plain-${n}`, `${n} characters of references`, zeros(list(2600, () => 'B1', '+'), n), over),
      one(`cf-${n}`, `a rule of ${n} characters`, {}, { edit: cfEdit(REFS(n)), ...over, ...(n > MAX_FORMULA_LENGTH ? {} : { holds: 'type="expression"' }) }),
      one(`dv-${n}`, `a validation rule of ${n} characters`, {}, { edit: dvEdit('custom', REFS(n)), ...over, ...(n > MAX_FORMULA_LENGTH ? {} : { holds: 'type="custom"' }) }),
      cell(`name-${n}`, `a name of ${n} characters`, 'k', { names: [{ name: 'k', ref: zeros(list(1000, () => 'S!$B$1', '+'), n) }], ...over }),
    ]
  }),
  cell('spaces-mid-255', '255 spaces between tokens', `1${rep(' ', 255)}+1`, { value: 2 }),
  cell('spaces-mid-256', '256 spaces between tokens', `1${rep(' ', 256)}+1`, REFUSED),
  cell('zeros-300', 'a number of 300 digits', `${rep('0', 299)}1`, { value: 1 }),
]

/** `n` conditional styles over the ids and codes `f` gives. */
const dxfs = (n: number, f: (i: number) => [id: number, code: string], before = '') =>
  subs(
    inSheet('<conditionalFormatting sqref="D1"><cfRule type="expression" dxfId="0" priority="1"><formula>B1&gt;0</formula></cfRule></conditionalFormatting>'),
    [STYLES, '</styleSheet>', `<dxfs count="${n}">${list(n, (i) => {
      const [id, code] = f(i)
      return `<dxf><numFmt numFmtId="${id}" formatCode="0.00&quot;${code}&quot;"/></dxf>`
    })}</dxfs></styleSheet>`],
    ...(before ? ([[STYLES, '<fonts ', `${before}<fonts `]] as [string, string, string][]) : []),
  )
const fmts = (n: number) => `<numFmts count="${n}">${list(n, (i) => `<numFmt numFmtId="${164 + i}" formatCode="0.00&quot;i${i + 1}&quot;"/>`)}</numFmts>`
/** 207 custom formats where two ids share a code. */
const sharedCode = subs(
  [STYLES, '<fonts ', `<numFmts count="207">${list(207, (i) => `<numFmt numFmtId="${164 + i}" formatCode="0.00&quot;i${Math.min(i + 1, 206)}&quot;"/>`)}</numFmts><fonts `],
  [STYLES, XFS, `<cellXfs count="2">${BASE_XF}<xf numFmtId="370" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>`],
  onStyle(1),
)

const formats = (): LinksCase[] => [
  one('dxf-one-id-600', '600 conditional styles of one id and code, as the saved books of the corpus hold them', {}, { edit: dxfs(600, () => [200, 'a']), holds: 'cfRule' }),
  one('dxf-206', '206 ids and codes in conditional styles', {}, { edit: dxfs(206, (i) => [200 + i, `d${i}`]), holds: 'cfRule' }),
  one('dxf-one-code-207', '207 ids of one code in conditional styles', {}, { edit: dxfs(207, (i) => [200 + i, 'a']), ...REFUSED }),
  one('dxf-one-id-207', '207 codes under one id in conditional styles', {}, { edit: dxfs(207, (i) => [200, `d${i}`]) }),
  one('mixed-206', '100 number formats and 106 others in conditional styles', {}, { edit: dxfs(106, (i) => [400 + i, `d${i}`], fmts(100)), holds: 'cfRule' }),
  one('mixed-207', '100 number formats and 107 others in conditional styles', {}, { edit: dxfs(107, (i) => [400 + i, `d${i}`], fmts(100)), ...REFUSED }),
  one('mixed-same-207', '200 number formats and 7 in conditional styles under the same ids', {}, { edit: dxfs(7, (i) => [164 + i, `i${i + 1}`], fmts(200)), holds: 'cfRule' }),
  one('shared-code-207', '207 number formats where two ids share a code', {}, { ...FIVE, edit: sharedCode, ...REFUSED }),
]

const SLICER_ITEMS = (b: { slicerCaches?: { items?: { item: number }[] }[] }) => {
  b.slicerCaches![0].items = [0, 1, 2, 3].map((item) => ({ item }))
}

const slicers = (): LinksCase[] => [
  pivoted('slicer-space', 'a slicer with its items, a cache whose name is no name', { change: SLICER_ITEMS, edit: subs(['xl/slicerCaches/slicerCache1.xml', 'name="Slicer_Region"', 'name="Slicer Region"']), ...REFUSED }),
  pivoted('slicer-taken', 'a slicer with its items, a cache whose name a defined name holds', { change: SLICER_ITEMS, names: [{ name: 'Slicer_Region', ref: 'Data!$A$1' }], ...REFUSED }),
]

/** What is left of a reference's place: `COUNTIF(@,1)`. */
const SHAPES: [string, string, boolean][] = [
  ['error-na', '#N/A', true],
  ['paren-num', '(1)', true],
  ['text', '"B1"', true],
  ['bool', 'TRUE', true],
  ['neg', '-B1:B3', true],
  ['percent', 'B1:B3%', true],
  ['local', '_xlfn.LET(_xlpm.x,B1:B3,COUNTIF(_xlpm.x,1))', false],
  ['local-num', '_xlfn.LET(_xlpm.x,1,COUNTIF(_xlpm.x,1))', false],
  ['lambda-call', '_xlfn.LAMBDA(_xlpm.x,_xlpm.x)(B1:B3)', false],
  ['name-call', 'nm(1)', false],
  ['single', '_xlfn.SINGLE(B1:B3)', false],
]

const shapes = (): LinksCase[] => [
  ...SHAPES.map(([k, v, refused]) => cell(`shape-${k}`, `COUNTIF over ${v}`, v.startsWith('_xlfn.LET') ? v : `COUNTIF(${v},1)`, { names: [{ name: 'nm', ref: 'S!$B$1:$B$3' }], ...(refused ? REFUSED : {}) })),
  cell('array-text-255', 'a text of 255 in an array constant', `LEN(INDEX({"${rep('a', 255)}"},1))`, { value: 255 }),
  cell('array-text-256', 'a text of 256 in an array constant', `LEN(INDEX({"${rep('a', 256)}"},1))`, REFUSED),
]

export function buildLinksCases(): LinksCases {
  return { cases: [...lengths(), ...formats(), ...slicers(), ...shapes()] }
}
