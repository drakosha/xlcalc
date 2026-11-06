/**
 * What Excel refuses a whole file over, against Excel for Windows and for Mac (`bench/coverage/links/file-refusals*.ts`), and what the
 * writer makes of it. A refusal by either platform counts: Excel for Mac 16.114 takes a text literal over 255 as `_LONGTEXT` where
 * Excel for Windows 16.0.20430 refuses the file, and functions one of them lacks are a user's call there. Where neither refuses, the
 * writer writes on: the specifications' stricter rules (`SUMIF(LET(…))`, ranges in a rule) are not limits.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredLinks } from '../bench/coverage/links.js'
import type { LinksCase, LinksCases } from '../bench/coverage/links/cases.js'
import type { BookInput, CellInput, ConditionalFormat, DataValidation, SheetInput } from '../src/contract.js'
import { parse } from '../src/syntax/parse.js'
import { fromFile } from '../src/xlsx/formulas/from-file.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { writeXlsx, type XlsxIssue } from '../src/xlsx.js'
import { pivotBook } from './pivot-books.js'

const SETS = ['file-refusals', 'file-refusals-2', 'file-refusals-3', 'file-refusals-4', 'file-refusals-5'] as const
type SetName = (typeof SETS)[number]

const stored = (dir: string, set: string) => JSON.parse(readFileSync(new URL(`./excel/${dir}links/${set}.json`, import.meta.url), 'utf8')) as StoredLinks & { cases: LinksCase[] }
const WIN = Object.fromEntries(SETS.map((s) => [s, stored('win/', s)])) as Record<SetName, ReturnType<typeof stored>>
const MAC = Object.fromEntries(SETS.map((s) => [s, stored('', s)])) as Record<SetName, ReturnType<typeof stored>>
const refused = (set: SetName, id: string) => !!WIN[set].rejected?.[id] || !!MAC[set].rejected?.[id]
const opened = (set: SetName, id: string) => !!WIN[set].answers[id] && (!MAC[set].rejected?.[id])
const kept = (dir: string, set: string) => JSON.parse(readFileSync(new URL(`./excel/${dir}file-refusals-kept/${set}.json`, import.meta.url), 'utf8')) as Record<string, boolean>

async function written(book: BookInput) {
  const r = await writeXlsx({ book }, { values: 'omit' })
  const archive = openArchive(r.data)
  const parts = new Map(await Promise.all(archive.names.map(async (n) => [n, (await archive.text(n)) ?? ''] as const)))
  const codes = (code: string) => r.issues.filter((i: XlsxIssue) => i.code === code)
  return { issues: r.issues, parts, codes, sheet: parts.get('xl/worksheets/sheet1.xml')!, styles: parts.get('xl/styles.xml')!, data: r.data }
}
const S = (more: Partial<SheetInput> = {}, cells: Record<string, CellInput> = {}): SheetInput => ({ name: 'S', cells: { A1: { value: 5 }, B1: { value: 1 }, B2: { value: 2 }, B3: { value: 3 }, ...cells }, ...more })

describe('the answers', () => {
  it.each(SETS)('%s: Excel for Windows answered every case, each as a file of its own', (set) => {
    const built = WIN[set].cases
    expect(Object.keys(WIN[set].answers).length + Object.keys(WIN[set].rejected ?? {}).length).toBe(built.length)
  })

  /** Committed before the run: `refused` of each case against the answers, a miss a miss. The rounds on the Mac ran part of the sets 2 to 5. */
  const SCORE: Record<SetName, [winHit: number, winMiss: number, macHit: number, macMiss: number]> = {
    'file-refusals': [487, 99, 476, 110],
    'file-refusals-2': [534, 31, 118, 28],
    'file-refusals-3': [39, 9, 16, 6],
    'file-refusals-4': [987, 3, 23, 11],
    'file-refusals-5': [46, 3, 19, 3],
  }

  it.each(SETS)('%s: the predictions scored what they scored', async (set) => {
    const built = ((await import(`../bench/coverage/links/${set}.ts`)) as { buildLinksCases(): LinksCases }).buildLinksCases().cases
    const score = (a: ReturnType<typeof stored>) => {
      const asked = built.filter((c) => a.answers[c.id] || a.rejected?.[c.id])
      const hit = asked.filter((c) => !!c.refused === !!a.rejected?.[c.id]).length
      return [hit, asked.length - hit]
    }
    expect([...score(WIN[set]), ...score(MAC[set])]).toEqual(SCORE[set])
  })
})

describe('a formula of a cell against Excel', () => {
  /** Files refused the writer writes: formulas of few operators and over 16,384 bytes of tokens, whose count the writer does not know. */
  const TOKENS = new Set(['sum-cells-over'])
  /** Their text is another as the writer prints it: spaces and a number's zeros go. */
  const PRINTED = new Set(['spaces-8195', 'spaces-mid-255', 'spaces-mid-256', 'xlfn-8193', 'xlfn-8200', 'xlfn-8500', 'xlfn-9000', 'xlfn-9500', 'xlfn-10000', 'zeros-300', 'plain-8192', 'plain-8193'])
  const cases = SETS.flatMap((set) => WIN[set].cases.filter((c) => !c.written && !c.names?.length && !c.links.length && Object.keys(c.linked.S ?? {}).join() === 'A1' && !PRINTED.has(c.id)).map((c) => [set, c] as const))

  it.each(cases.map(([set, c]) => [c.id, set, c] as const))('%s', async (id, set, c) => {
    const formula = fromFile(c.linked.S.A1 as string).formula
    let reads = true
    try {
      parse(formula)
    } catch {
      reads = false
    }
    const r = await written({ sheets: [{ ...c.sheets[0], cells: { ...c.sheets[0].cells, A1: { formula, value: 7 } as CellInput } }] })
    const has = /<c r="A1"[^>]*><f/.test(r.sheet)
    if (TOKENS.has(id)) expect(has).toBe(true)
    else if (refused(set, id)) expect(has).toBe(false)
    else if (reads) expect(has).toBe(true)
    if (!has) expect(r.codes('formula').length).toBe(1)
  })
})

describe('what a rule holds', () => {
  const TABLE: Partial<SheetInput> = { tables: [{ name: 'Sales', range: 'E1:E3', columns: [{ name: 'Qty' }] }] }
  const TABLE_CELLS = { E1: { value: 'Qty' }, E2: { value: 1 }, E3: { value: 2 } }
  const others = [{ name: 'T', cells: { B1: { value: 1 } } }]
  const CONDITIONAL: [id: string, formula: string][] = [
    ['cf-plain', 'B1>0'],
    ['cf-array', 'OR(B1={1,2,3})'],
    ['cf-structured', 'B1>SUM(Sales[Qty])'],
    ['cf-union', 'SUM((B1,B2))>0'],
    ['cf-intersect', 'SUM(B1:B3 B2:B3)>0'],
    ['cf-3d', "SUM('S:T'!B1)>0"],
    ['cf-lambda', 'LAMBDA(x,x)(1)>0'],
    ['cf-let', 'LET(x,1,x)>0'],
    ['cf-range-op', 'SUM(B1:INDEX(B1:B3,2))>0'],
    ['cf-own-sheet', 'S!B1>0'],
    ['cf-other-sheet', 'T!B1>0'],
    ['cf-text-255', `LEN("${'a'.repeat(255)}")>0`],
    ['cf-text-256', `LEN("${'a'.repeat(256)}")>0`],
  ]
  it.each(CONDITIONAL)('%s', async (id, formula) => {
    const rule: ConditionalFormat = { type: 'expression', ranges: ['D1'], formula, style: { font: { bold: true } } }
    const r = await written({ sheets: [S({ ...TABLE, conditionalFormats: [rule] }, TABLE_CELLS), ...others] })
    const has = r.sheet.includes('cfRule')
    expect(has).toBe(!refused('file-refusals', id))
    expect(r.codes('conditionalFormat').length).toBe(has ? 0 : 1)
  })

  const VALIDATION: [id: string, type: DataValidation['type'], formula: string][] = [
    ['dv-plain', 'custom', 'B1>0'],
    ['dv-array', 'custom', 'OR(B1={1,2,3})'],
    ['dv-list-array', 'list', '{1,2,3}'],
    ['dv-structured', 'list', 'Sales[Qty]'],
    ['dv-union', 'custom', 'SUM((B1,B2))>0'],
    ['dv-intersect', 'custom', 'SUM(B1:B3 B2:B3)>0'],
    ['dv-3d', 'custom', "SUM('S:T'!B1)>0"],
    ['dv-lambda', 'custom', 'LAMBDA(x,x)(1)>0'],
    ['dv-list-other', 'list', 'T!$B$1:$B$3'],
    ['dv-list-indirect', 'list', 'INDIRECT("Sales[Qty]")'],
    ['dv-list-255', 'list', `"${'abc,'.repeat(64).slice(0, 255)}"`],
    ['dv-list-256', 'list', `"${'abc,'.repeat(64).slice(0, 256)}"`],
    ['dv-custom-256', 'custom', `LEN("${'a'.repeat(256)}")>0`],
  ]
  it.each(VALIDATION)('%s', async (id, type, formula1) => {
    const r = await written({ sheets: [S({ ...TABLE, validations: [{ range: 'D1', type, formula1 }] }, TABLE_CELLS), ...others] })
    const has = r.sheet.includes('<dataValidation ')
    expect(has).toBe(!refused('file-refusals', id))
    expect(r.codes('dataValidation').length).toBe(has ? 0 : 1)
  })

  it('a dynamic list by the range operator is refused, by a defined name or OFFSET it is not', async () => {
    expect(refused('file-refusals-2', 'dv-list-range-op')).toBe(true)
    const off = await written({ sheets: [S({ validations: [{ range: 'D1', type: 'list', formula1: '$B$1:INDEX($B$1:$B$3,2)' }] })] })
    expect(off.codes('dataValidation').length).toBe(1)
    expect(opened('file-refusals-2', 'cf-range-offset')).toBe(true)
  })

  it('rules beyond 65,534 are left out with their issue', async () => {
    expect(refused('file-refusals', 'count-dv-65535')).toBe(true)
    const rules: DataValidation[] = Array.from({ length: 65_535 }, (_, i) => ({ range: `D${i + 1}`, type: 'custom', formula1: 'TRUE' }))
    const r = await written({ sheets: [S({ validations: rules })] })
    expect(r.sheet).toContain('<dataValidations count="65534"')
    expect(r.codes('dataValidation').length).toBe(1)
  })
})

describe('names', () => {
  it.each([
    ['name-rel', 'C1'],
    ['name-abs', '$C$2'],
    ['name-calc', 'C1*2'],
    ['name-sum', 'SUM(C1:C2)'],
  ])('%s: a reference without its sheet costs the file, and the name is left out', async (id, ref) => {
    expect(refused('file-refusals', id)).toBe(true)
    const r = await written({ sheets: [S()], names: [{ name: 'x', ref }] })
    expect(r.parts.get('xl/workbook.xml')).not.toContain('definedName')
    expect(r.codes('name').length).toBe(1)
  })

  it('a calculation over constants, or with sheets, is written', async () => {
    const r = await written({ sheets: [S()], names: [{ name: 'k', ref: '1+1' }, { name: 'ok', ref: 'S!$A$1' }] })
    expect(r.parts.get('xl/workbook.xml')).toContain('name="k"')
    expect(r.parts.get('xl/workbook.xml')).toContain('name="ok"')
    expect(r.codes('name')).toEqual([])
  })

  it('names the specification reserves open: no limit on them', () => {
    for (const id of ['name-xl-foo', 'name-xlnm-foo', 'name-xlcalc', 'name-backslash-a', 'name-backslash', 'name-backslash-cyr']) expect(opened('file-refusals', id)).toBe(true)
  })
})

describe('styles', () => {
  const formats = (n: number) => Object.fromEntries(Array.from({ length: n }, (_, i) => [`A${i + 1}`, { value: i, numberFormat: `0.00"${i}"` }]))

  it('206 different codes stay, the others are General in the order of first appearance, and no value changes', async () => {
    expect(refused('file-refusals', 'numfmt-207')).toBe(true)
    expect(opened('file-refusals', 'numfmt-206')).toBe(true)
    const r = await written({ sheets: [{ name: 'S', cells: formats(208) }] })
    expect(r.styles).toContain('<numFmts count="206">')
    expect(r.styles).toContain('0.00&quot;205&quot;')
    expect(r.styles).not.toContain('0.00&quot;206&quot;')
    expect(r.codes('numberFormat').length).toBe(1)
    for (let i = 0; i < 208; i++) expect(r.sheet).toMatch(new RegExp(`<c r="A${i + 1}"[^>]*><v>${i}</v></c>`))
    expect(r.sheet).toContain('<c r="A207"><v>206</v></c>')
  })

  it('a code in a conditional style counts with the others, a copy does not, and its rule stays without the format', async () => {
    expect(refused('file-refusals-5', 'numfmt-206-dxf-new')).toBe(true)
    expect(opened('file-refusals-5', 'numfmt-206-dxf-same')).toBe(true)
    expect(refused('file-refusals-3', 'mixed-207')).toBe(true)
    const rule = (code: string): ConditionalFormat => ({ type: 'expression', ranges: ['D1'], formula: 'A1>0', style: { numberFormat: code, font: { bold: true } } })
    const over = await written({ sheets: [{ name: 'S', cells: formats(206), conditionalFormats: [rule('0.00"new"')] }] })
    expect(over.styles).not.toContain('new')
    expect(over.sheet).toContain('cfRule')
    expect(over.codes('numberFormat').length).toBe(1)
    const same = await written({ sheets: [{ name: 'S', cells: formats(206), conditionalFormats: [rule('0.00"0"')] }] })
    expect(same.codes('numberFormat')).toEqual([])
  })

  it('cell formats beyond 65,430 are format 0, no value changes', async () => {
    expect(refused('file-refusals', 'xf-65431')).toBe(true)
    expect(opened('file-refusals', 'xf-65430')).toBe(true)
    const n = 65_432
    const styles = Array.from({ length: n }, (_, i) => ({ fill: { color: `#${i.toString(16).padStart(6, '0')}` } }))
    const r = await written({ styles, sheets: [{ name: 'S', cells: Object.fromEntries(styles.map((_, i) => [`A${i + 1}`, { value: i, style: i }])) }] })
    expect(r.styles).toContain('<cellXfs count="65430">')
    expect(r.sheet).toContain(`<c r="A${n}"><v>${n - 1}</v></c>`)
    expect(r.codes('style').length).toBe(1)
  })

  it('a font name is cut at 31 and an empty one is the workbook’s', async () => {
    for (const id of ['font-cell-32', 'font-cell-empty', 'font-book-32', 'font-book-empty']) expect(refused('file-refusals', id)).toBe(true)
    for (const id of ['font-cell-31', 'font-book-31']) expect(opened('file-refusals', id)).toBe(true)
    const cut = await written({ styles: [{ font: { name: 'F'.repeat(40) } }], sheets: [S({}, { A1: { value: 5, style: 0 } })] })
    expect(cut.styles).toContain(`<name val="${'F'.repeat(31)}"/>`)
    expect(cut.codes('textLength').length).toBe(1)
    const empty = await written({ styles: [{ font: { name: '', bold: true } }], sheets: [S({}, { A1: { value: 5, style: 0 } })] })
    expect(empty.styles).not.toContain('<name val=""/>')
    expect(empty.codes('style').length).toBe(1)
    const book = await written({ font: { name: 'F'.repeat(40), size: 11 }, sheets: [S()] })
    expect(book.styles).toContain(`<name val="${'F'.repeat(31)}"/>`)
  })
})

describe('hyperlinks', () => {
  it('a location over 2,084 leaves its link out, the cell kept; an address has no such limit', async () => {
    expect(refused('file-refusals', 'link-location-2085')).toBe(true)
    expect(opened('file-refusals', 'link-location-2084')).toBe(true)
    expect(opened('file-refusals', 'link-url-3000')).toBe(true)
    const long = await written({ sheets: [S({}, { A1: { value: 5, link: { location: `S!A1${'x'.repeat(2081)}` } } })] })
    expect(long.sheet).not.toContain('<hyperlink')
    expect(long.sheet).toContain('<c r="A1"><v>5</v></c>')
    expect(long.codes('hyperlink').length).toBe(1)
    const edge = await written({ sheets: [S({}, { A1: { value: 5, link: { location: `S!A1${'x'.repeat(2080)}` } } })] })
    expect(edge.sheet).toContain('<hyperlink')
    const url = await written({ sheets: [S({}, { A1: { value: 5, link: { url: `https://e.com/?${'q'.repeat(3000)}` } } })] })
    expect(url.sheet).toContain('<hyperlink')
    expect(url.issues).toEqual([])
  })
})

describe('a table’s totals', () => {
  const book = (foot: CellInput | undefined, totals?: { function: 'sum' } | { label: string }) => ({ sheets: [S({ tables: [{ name: 'Tbl', range: 'E1:E3', totals: true, columns: [{ name: 'Q', ...(totals ? { totals } : {}) }] }] }, { E1: { value: 'Q' }, E2: { value: 1 }, ...(foot ? { E3: foot } : {}) })] }) as BookInput
  const column = async (b: BookInput) => /<tableColumn .*<\/tableColumns>/.exec((await written(b)).parts.get('xl/tables/table1.xml')!)![0]

  it.each([
    ['totals-none-value', { value: 42 }, undefined, 'totalsRowFunction="custom"/>'],
    ['label-10', { value: 'L' }, { label: 'LLLLLLLLLL' }, 'name="Q"/>'],
    ['totals-sum-other', { formula: 'E2*2' }, { function: 'sum' }, '<totalsRowFormula>E2*2</totalsRowFormula>'],
    ['totals-sum-sum', { formula: 'SUM(E2:E2)' }, { function: 'sum' }, '<totalsRowFormula>SUM(E2:E2)</totalsRowFormula>'],
  ] as const)('%s, which Excel refuses, is written as Excel opens it', async (id, foot, totals, expected) => {
    expect(refused('file-refusals', id)).toBe(true)
    expect(await column(book(foot, totals))).toContain(expected)
  })

  it.each([
    ['totals-sum-subtotal', { formula: 'SUBTOTAL(109,Tbl[Q])' }, { function: 'sum' }, 'totalsRowFunction="sum"/>'],
    ['totals-sum-value', { value: 42 }, { function: 'sum' }, 'totalsRowFunction="sum"/>'],
    ['totals-sum-empty', undefined, { function: 'sum' }, 'totalsRowFunction="sum"/>'],
  ] as const)('%s, which Excel opens, stays', async (id, foot, totals, expected) => {
    expect(opened('file-refusals', id)).toBe(true)
    expect(await column(book(foot, totals))).toContain(expected)
  })

  it('a label that is the cell’s text stays, over an empty cell too', async () => {
    expect(opened('file-refusals-2', 'totals-label-same')).toBe(true)
    expect(opened('file-refusals-2', 'totals-label-empty')).toBe(true)
    expect(await column(book({ value: 'Total' }, { label: 'Total' }))).toContain('totalsRowLabel="Total"')
    expect(await column(book(undefined, { label: 'Total' }))).toContain('totalsRowLabel="Total"')
  })
})

describe('PivotTables and slicers', () => {
  const noSlicer = (b: BookInput) => {
    b.slicerCaches = []
    b.sheets[1].slicers = []
    return b
  }
  const CACHE = 'xl/pivotCache/pivotCacheDefinition1.xml'

  it('a slicer of no items takes its field’s, which opens', async () => {
    expect(refused('file-refusals', 'pivot-slicer')).toBe(true)
    expect(opened('file-refusals-2', 'slicer-items')).toBe(true)
    // the frame it needs to stay on saving: test/xlsx-slicer-names.test.ts
    for (const dir of ['', 'win/']) expect(kept(dir, 'file-refusals-2')['slicer-items']).toBe(false)
    const b = pivotBook()
    b.sheets[1].slicers![0].place = { from: 'H2', to: 'K12' }
    const r = await written(b)
    expect(r.parts.get('xl/slicerCaches/slicerCache1.xml')).toContain('<items count="4"><i x="0" s="1"/>')
  })

  it('a slicer cache whose name is no name is left out', async () => {
    expect(refused('file-refusals-3', 'slicer-space')).toBe(true)
    const b = pivotBook()
    b.slicerCaches![0].name = 'Slicer Region'
    b.sheets[1].slicers![0].cache = 'Slicer Region'
    const r = await written(b)
    expect([...r.parts.keys()].some((p) => p.startsWith('xl/slicerCaches/'))).toBe(false)
    expect(r.codes('pivotTable').length).toBe(1)
  })

  it('an error a cache does not take is #VALUE!, as measured error by error', async () => {
    for (const id of ['pivot-error-calc', 'pivot-error-cycle', 'pivot-error-connect', 'pivot-error-blocked', 'pivot-error-error']) expect(refused('file-refusals-5', id)).toBe(true)
    for (const id of ['pivot-error-busy', 'pivot-error-getting']) expect(opened('file-refusals-5', id)).toBe(true)
    expect(refused('file-refusals', 'pivot-error-spill')).toBe(true)
    const b = noSlicer(pivotBook())
    b.pivotCaches![0].fields[0].items = ['North', 'South', 'East', { error: 'SPILL' }]
    const r = await written(b)
    expect(r.parts.get(CACHE)).toContain('<e v="#VALUE!"/>')
    expect(r.codes('errorValue').length).toBe(1)
  })

  it.each([
    ['pivot-field-dup', 'region'],
    ['pivot-field-empty', ''],
  ])('%s: the PivotTable is left out with its cache, its cells kept', async (id, name) => {
    expect(refused('file-refusals', id)).toBe(true)
    const b = noSlicer(pivotBook())
    b.pivotCaches![0].fields[2].name = name
    const r = await written(b)
    expect([...r.parts.keys()].some((p) => p.includes('pivot'))).toBe(false)
    expect(r.parts.get('xl/worksheets/sheet2.xml')).toContain('<c r="B8"><v>360</v></c>')
    expect(r.codes('pivotTable').length).toBe(1)
  })

  it('a name over 255 is cut, and two cut alike stay apart', async () => {
    expect(refused('file-refusals', 'pivot-name-256')).toBe(true)
    expect(opened('file-refusals', 'pivot-name-255')).toBe(true)
    // Excel opens two of a name on a sheet (the set's `pivot-twin-sheet`); still, a slicer names one by its name
    expect(opened('file-refusals-5', 'pivot-twin-sheet')).toBe(true)
    const b = noSlicer(pivotBook())
    const p = b.sheets[1].pivotTables![0]
    b.sheets[1].pivotTables = [{ ...p, name: `${'p'.repeat(299)}a` }, { ...p, name: `${'p'.repeat(299)}b`, location: { ...p.location, range: 'D3:E8' } }]
    const r = await written(b)
    const names = ['1', '2'].map((n) => /<pivotTableDefinition [^>]*?name="([^"]*)"/.exec(r.parts.get(`xl/pivotTables/pivotTable${n}.xml`)!)![1])
    expect(names.map((n) => n.length)).toEqual([255, 255])
    expect(names[0]).not.toBe(names[1])
    expect(r.codes('textLength').length).toBe(1)
  })

  it('a caption over 255 is cut', async () => {
    expect(refused('file-refusals', 'pivot-caption-256')).toBe(true)
    const b = noSlicer(pivotBook())
    b.sheets[1].pivotTables![0].captions = { grandTotal: 'g'.repeat(300) }
    const r = await written(b)
    expect(r.parts.get('xl/pivotTables/pivotTable1.xml')).toContain(`grandTotalCaption="${'g'.repeat(255)}"`)
  })
})

describe('the writer’s own files open', () => {
  /** Each file `writeXlsx` writes of what the other rounds showed refused; the Mac ran part of them. */
  const OWN = WIN['file-refusals-5'].cases.filter((c) => c.written).map((c) => c.id)

  it.each(OWN)('%s', (id) => {
    expect(WIN['file-refusals-5'].rejected?.[id]).toBeUndefined()
    expect(MAC['file-refusals-5'].rejected?.[id]).toBeUndefined()
    expect(WIN['file-refusals-5'].answers[id]).toBeDefined()
  })
})

describe('where the specifications are stricter than Excel', () => {
  it.each([
    ['file-refusals', 'call-sumif-3-let'],
    ['file-refusals', 'call-countif-xlookup'],
    ['file-refusals', 'call-row-ifs'],
    ['file-refusals', 'count-dv-32768'],
    ['file-refusals', 'count-cf-8193'],
    ['file-refusals', 'count-select-8193'],
    ['file-refusals', 'cf-3d'],
    ['file-refusals', 'cfvo-icon-min'],
    ['file-refusals', 'fill-65431'],
    ['file-refusals', 'border-65431'],
    ['file-refusals', 'link-url-3000'],
    ['file-refusals', 'note-author-65536'],
    ['file-refusals', 'ext-name-bad'],
    ['file-refusals-3', 'dxf-one-code-207'],
  ] as const)('%s %s opens on both', (set, id) => {
    expect(WIN[set].answers[id]).toBeDefined()
    expect(MAC[set].answers[id]).toBeDefined()
  })
})
