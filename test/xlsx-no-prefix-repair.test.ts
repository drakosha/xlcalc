/**
 * `writeKept`'s `prefixes: 'repair'`: the formulas of a file that call newer functions without their prefix (`noPrefix`),
 * as openpyxl writes them, written again with it, and nothing else; the engine reads the file repaired as it read it
 * before. Against Excel for Mac and for Windows over openpyxl's files (`bench/coverage/keep/prefixes.ts`).
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { PREFIX_BOOKS, prefixMade, readCells, repaired, type Shown } from '../bench/coverage/keep/prefixes.js'
import { compute, createSession, InputError } from '../src/index.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { repairPrefixes } from '../src/xlsx/formulas/repair.js'
import { readXlsx, writeXlsx, type XlsxIssue } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { changedParts, partsOf } from './keep-files.js'
import { ruRU } from './locales.js'

const noPrefix = (issues: XlsxIssue[]) => issues.filter((i) => i.code === 'noPrefix').map(({ fn, sheet, name, cells }) => ({ fn, sheet, name, cells }))

/** The values the engine gives every cell of a file read. */
async function values(data: Uint8Array) {
  const r = await readXlsx(data)
  return compute(r.book, { cached: r.cached }).sheets.map((s) => Object.fromEntries(Object.entries(s.cells).map(([k, c]) => [k, c.value])))
}

/** A file written over itself with the engine's values, as a host writes it, with `opts`. */
async function over(original: Uint8Array, opts?: Parameters<typeof writeKept>[2]) {
  const { book, cached } = await openKept(original)
  return writeKept(original, { book, result: compute(book, { cached }) }, opts)
}

describe('repairPrefixes', () => {
  it.each([
    ['XLOOKUP(2,A1:A2,B1:B2)', '_xlfn.XLOOKUP(2,A1:A2,B1:B2)'],
    ['IFERROR(XLOOKUP(2,A1:A2,B1:B2),"fallback")', 'IFERROR(_xlfn.XLOOKUP(2,A1:A2,B1:B2),"fallback")'],
    ['FILTER(A1:A3,B1:B3>1)', '_xlfn._xlws.FILTER(A1:A3,B1:B3>1)'],
    ['xlookup(1,{1,2},{3,4})+_xlfn.IFNA(1,2)+DAYS(10,1)', '_xlfn.XLOOKUP(1,{1,2},{3,4})+_xlfn.IFNA(1,2)+_xlfn.DAYS(10,1)'],
    ['"XLOOKUP(" & DAYS(2,1)', '"XLOOKUP(" & _xlfn.DAYS(2,1)'],
    ['MAKEARRAY(2,2,LAMBDA(r,c,r))', '_xlfn.MAKEARRAY(2,2,_xlfn.LAMBDA(_xlpm.r,_xlpm.c,_xlpm.r))'],
    ['LET(a, A1, a + A$1 + SUM($A:A))', '_xlfn.LET(_xlpm.a, A1, _xlpm.a + A$1 + SUM($A:A))'],
    ['S!A1+LET(x,S!x,x)', 'S!A1+_xlfn.LET(_xlpm.x,S!x,_xlpm.x)'],
    ['LET(f,LAMBDA(x,[y],IF(ISOMITTED(y),x,x+y)),f(2))', '_xlfn.LET(_xlpm.f,_xlfn.LAMBDA(_xlpm.x,_xlop.y,IF(_xlfn.ISOMITTED(_xlpm.y),_xlpm.x,_xlpm.x+_xlpm.y)),_xlpm.f(2))'],
  ])('%s is written %s', (f, out) => {
    expect(repairPrefixes(f)?.formula).toBe(out)
  })

  it.each([['_xlfn._xlws.SORT(A1:A2)'], ['SUM(A1)'], ['"XLOOKUP(1)"'], ['NETWORKDAYS.INTL(1,2)'], ['Sales[XLOOKUP(1)]']])('leaves %s alone', (f) => {
    expect(repairPrefixes(f)).toBeUndefined()
  })

  it('names the functions as noPrefix does, once, in the order of their calls', () => {
    expect(repairPrefixes('DAYS(1,2)+XLOOKUP(1,A1,B1)+DAYS(3,4)')?.fns).toEqual(['DAYS', 'XLOOKUP'])
  })
})

describe('writeKept with prefixes: repair', () => {
  it.each(PREFIX_BOOKS.map((b) => b.id))("%s: the reader finds no call without its prefix after, the engine's values stay, the repaired calls are the read ones", async (id) => {
    const original = prefixMade(id)
    const read = await readXlsx(original)
    expect(noPrefix(read.issues).length).toBeGreaterThan(0)
    const w = await over(original, { prefixes: 'repair' })
    const back = await readXlsx(w.data)
    expect(noPrefix(back.issues)).toEqual([])
    expect(await values(w.data)).toEqual(await values(original))
    expect(noPrefix(w.issues)).toEqual(noPrefix(read.issues))
    // the same formulas, of the same kind, for the engine: the case of a name typed in lower case aside
    const formulas = (r: typeof read) => JSON.stringify([r.book.sheets.map((s) => s.cells), r.book.names, r.book.sheets.map((s) => [s.conditionalFormats, s.validations, s.tables])]).toUpperCase()
    expect(formulas(back)).toBe(formulas(read))
  })

  it('writes the worksheets alone: the charts, the drawing, the pivot table and its cache stay byte for byte', async () => {
    const original = prefixMade('pivot')
    const w = await over(original, { prefixes: 'repair' })
    // the sheets take the result's values, as for any change; openpyxl's workbook asks for a full calculation already
    expect(await changedParts(original, w.data)).toEqual(['xl/worksheets/sheet1.xml', 'xl/worksheets/sheet3.xml'])
    const [a, b] = [openArchive(original), openArchive(w.data)]
    expect(a.names.filter((n) => /chart|drawing|pivot/.test(n)).length).toBe(7)
    for (const name of a.names) if (!/worksheets\/sheet/.test(name)) expect(await b.bytes(name)).toEqual(await a.bytes(name))
  })

  it("rewrites the formulas' text alone", async () => {
    const original = prefixMade('lookup')
    const w = await over(original, { prefixes: 'repair' })
    const f = async (data: Uint8Array) => [...(await partsOf(data)).get('xl/worksheets/sheet1.xml')!.matchAll(/<f>[^<]*<\/f>/g)].map((m) => m[0])
    expect(await f(w.data)).toEqual(['<f>_xlfn.XLOOKUP(2,A1:A2,B1:B2)</f>', '<f>_xlfn._xlws.SORT(A1:A2)</f>', '<f>IFERROR(_xlfn.XLOOKUP(2,A1:A2,B1:B2),"fallback")</f>', '<f>_xlfn.XLOOKUP(1,A1:A2,B1:B2)</f>'])
    expect(await f(original)).toEqual(['<f>XLOOKUP(2,A1:A2,B1:B2)</f>', '<f>_xlfn._xlws.SORT(A1:A2)</f>', '<f>IFERROR(XLOOKUP(2,A1:A2,B1:B2),"fallback")</f>', '<f>xlookup(1,A1:A2,B1:B2)</f>'])
  })

  it("keeps a shared formula shared: the master's text repaired, its copies as they were", async () => {
    const w = await over(prefixMade('shared'), { prefixes: 'repair' })
    const xml = (await partsOf(w.data)).get('xl/worksheets/sheet1.xml')!
    expect(xml).toContain('<c r="E1"><f t="shared" ref="E1:E4" si="0">_xlfn.XLOOKUP(D1,$A$1:$A$4,$B$1:$B$4)</f>')
    expect(xml).toContain('<c r="E2"><f t="shared" si="0"/>')
    expect(noPrefix(w.issues)).toEqual([{ fn: 'XLOOKUP', sheet: 'S', cells: ['E1', 'F1', 'E2', 'E3', 'E4', 'F2', 'F3', 'F4'] }])
  })

  it('keeps the kind of a formula: FILTER of the old kind stays one, an array formula of the old kind keeps its range', async () => {
    const filter = (await partsOf((await over(prefixMade('filter'), { prefixes: 'repair' })).data)).get('xl/worksheets/sheet1.xml')!
    expect(filter).toContain('<c r="C1"><f>_xlfn._xlws.FILTER(B1:B4,A1:A4&gt;2)</f>')
    expect(filter).not.toContain('cm=')
    const array = (await partsOf((await over(prefixMade('array'), { prefixes: 'repair' })).data)).get('xl/worksheets/sheet1.xml')!
    expect(array).toContain('<f t="array" ref="C1:C3">_xlfn._xlws.SORT(A1:A3)</f>')
    expect((await readXlsx((await over(prefixMade('filter'), { prefixes: 'repair' })).data)).book.sheets[0].cells.C1.formula).toBe('@FILTER(B1:B4,A1:A4>2)')
  })

  it('repairs defined names, rules, validation and a table column in their parts', async () => {
    const names = (await partsOf((await over(prefixMade('name'), { prefixes: 'repair' })).data)).get('xl/workbook.xml')!
    expect(names).toContain('<definedName name="Pick">_xlfn.XLOOKUP(2,S!$A$1:$A$2,S!$B$1:$B$2)</definedName><definedName name="Local" localSheetId="0">_xlfn.IFNA(_xlfn.XLOOKUP(9,')
    const rules = (await partsOf((await over(prefixMade('rules'), { prefixes: 'repair' })).data)).get('xl/worksheets/sheet1.xml')!
    expect(rules).toContain('<formula>_xlfn.XLOOKUP(A1,$A$1:$A$2,$B$1:$B$2)&gt;15</formula>')
    expect(rules).toContain('<formula1>_xlfn.XLOOKUP(B1,$B$1:$B$2,$A$1:$A$2)&gt;0</formula1>')
    const table = (await partsOf((await over(prefixMade('table'), { prefixes: 'repair' })).data)).get('xl/tables/table1.xml')!
    expect(table).toContain('<calculatedColumnFormula>_xlfn.XLOOKUP(T[[#This Row],[a]],$E$1:$E$2,$F$1:$F$2)</calculatedColumnFormula>')
  })

  it("repairs Excel 2010's rules in the sheet's extensions and a threshold in an attribute", async () => {
    const original = prefixMade('rules')
    const parts = await partsOf(original)
    const ext = '<extLst><ext uri="{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list" allowBlank="1"><x14:formula1><xm:f>SORT(S!$A$1:$A$2)</xm:f></x14:formula1><xm:sqref>D1</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>'
    const cfvo = '<conditionalFormatting sqref="B1:B2"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="formula" val="DAYS(10,1)"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule></conditionalFormatting>'
    const sheet = parts.get('xl/worksheets/sheet1.xml')!.replace('<dataValidations', `${cfvo}<dataValidations`).replace('</worksheet>', `${ext}</worksheet>`)
    const { zip } = await import('../src/xlsx/zip/write.js')
    const data = await zip([...parts].map(([name, text]) => ({ name, data: new TextEncoder().encode(name === 'xl/worksheets/sheet1.xml' ? sheet : text) })))
    const w = await over(data, { prefixes: 'repair' })
    const xml = (await partsOf(w.data)).get('xl/worksheets/sheet1.xml')!
    expect(xml).toContain('<xm:f>_xlfn._xlws.SORT(S!$A$1:$A$2)</xm:f>')
    expect(xml).toContain('<cfvo type="formula" val="_xlfn.DAYS(10,1)"/>')
    expect(noPrefix((await readXlsx(w.data)).issues)).toEqual([])
    expect(noPrefix(w.issues)).toEqual(noPrefix((await readXlsx(data)).issues))
  })

  it('with nothing to repair gives the original bytes, and without the option the file keeps its formulas', async () => {
    const book = { sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { formula: 'XLOOKUP(2,{1,2},{3,4})' } } }] }
    const made = (await writeXlsx({ book, result: compute(book) })).data
    expect((await over(made, { prefixes: 'repair' })).data).toBe(made)
    for (const b of PREFIX_BOOKS) {
      const original = prefixMade(b.id)
      const read = noPrefix((await readXlsx(original)).issues)
      for (const opts of [undefined, { prefixes: 'keep' as const }]) expect(noPrefix((await readXlsx((await over(original, opts)).data)).issues)).toEqual(read)
    }
    const lookup = prefixMade('lookup')
    expect((await over(lookup)).data).toBe(lookup)
  })

  it('re-entering the same formula changes nothing; with the option the formula is repaired and its value stored', async () => {
    const original = prefixMade('lookup')
    const { book, cached } = await openKept(original)
    const s = createSession(book, { cached })
    s.setCells([{ sheet: 'S', cell: 'C1', input: { formula: 'XLOOKUP(2,A1:A2,B1:B2)' } }])
    const kept = await writeKept(original, s.getSnapshot(), { writeValues: true })
    expect(noPrefix((await readXlsx(kept.data)).issues).length).toBe(1)
    const fixed = await writeKept(original, s.getSnapshot(), { writeValues: true, prefixes: 'repair' })
    const back = await readXlsx(fixed.data)
    expect(noPrefix(back.issues)).toEqual([])
    expect(back.cached.sheets[0].cells.C1).toEqual({ value: 20 })
    expect((await partsOf(fixed.data)).get('xl/worksheets/sheet1.xml')).toContain('<c r="C1"><f>_xlfn.XLOOKUP(2,A1:A2,B1:B2)</f><v>20</v></c>')
    s.close()
  })

  it('repairs the formulas moved by an operation of the structure and those an edit kept', async () => {
    const original = prefixMade('shared')
    const { book, cached } = await openKept(original)
    const s = createSession(book, { cached })
    s.editStructure({ op: 'insertRows', sheet: 'S', row: 2 })
    s.setCells([{ sheet: 'S', cell: 'D1', input: { value: 3 } }])
    const w = await writeKept(original, s.getSnapshot(), { prefixes: 'repair' })
    const back = await readXlsx(w.data)
    expect(noPrefix(back.issues)).toEqual([])
    const got = compute(back.book, { cached: back.cached }).sheets[0].cells
    const want = s.getSnapshot().result.sheets[0].cells
    for (const a of ['E1', 'E3', 'E4', 'E5', 'F1', 'F3', 'F4', 'F5']) expect(got[a].value).toEqual(want[a].value)
    s.close()
  })

  it('reports in the locale of the call', async () => {
    const w = await over(prefixMade('lookup'), { prefixes: 'repair', locale: ruRU })
    expect(w.issues.find((i) => i.code === 'noPrefix')?.message).toContain('_xlfn.XLOOKUP')
    expect(w.issues.find((i) => i.code === 'noPrefix')?.message).toMatch(/записана с ним/)
  })

  it("refuses an option it does not take", async () => {
    await expect(over(prefixMade('lookup'), { prefixes: 'fix' as 'repair' })).rejects.toThrow(InputError)
  })
})

/** The answers of a platform: Excel's refusal of a case, or the file it saved; a platform not run yet is named as skipped. */
const PLATFORMS = [
  ['Excel for Mac', 'keep/prefixes'],
  ['Excel for Windows', 'win/keep/prefixes'],
] as const
const answered = (dir: string) => existsSync(new URL(`./excel/${dir}.json`, import.meta.url))
const predicted = JSON.parse(readFileSync(new URL('../bench/coverage/keep/prefixes.predicted.json', import.meta.url), 'utf8')).cases as Record<string, { opens: boolean; opensOnWindows?: false; values?: Record<string, Shown>; formulas?: Record<string, string | null> }>

describe.each(PLATFORMS)('openpyxl files repaired, against %s', (platform, dir) => {
  if (!answered(dir)) {
    it.skip(`${platform} has not run bench/coverage/keep/prefixes.ts yet`)
    return
  }
  type Stored = { rejected?: Record<string, string>; cases?: Record<string, { rejected?: string }> }
  const stored = () => JSON.parse(readFileSync(new URL(`./excel/${dir}.json`, import.meta.url), 'utf8')) as Stored
  // the Mac's answers by case, Windows's refusals apart
  const refused = (id: string) => {
    const s = stored()
    return s.cases ? s.cases[id]?.rejected !== undefined : s.rejected?.[id] !== undefined
  }
  const saved = (id: string) => new Uint8Array(readFileSync(new URL(`./excel/${dir}/${id}.xlsx`, import.meta.url)))
  const windows = platform === 'Excel for Windows'

  it.each(PREFIX_BOOKS.map((b) => b.id))('%s: Excel opens the repaired file, shows the engine values and keeps the formulas of the old kind', async (id) => {
    const want = predicted[`${id}-repaired`]
    expect(refused(`${id}-repaired`)).toBe(false)
    const b = PREFIX_BOOKS.find((x) => x.id === id)!
    const got = await readCells(saved(`${id}-repaired`), b.read)
    // The committed prediction remains 3: Windows Excel gives 1 for SUM(XLOOKUP({1;2},...)) in a legacy CSE cell.
    // Only this measured miss is accepted; every other value and the formula text must match the prediction.
    const misses = windows && id === 'array' ? { G1: 1 } : {}
    expect(got.values).toEqual({ ...want.values, ...misses })
    expect(got.formulas).toEqual(want.formulas)
    expect(noPrefix((await readXlsx(saved(`${id}-repaired`))).issues)).toEqual([])
  })

  it.each(PREFIX_BOOKS.map((b) => b.id))('%s: the original as predicted', async (id) => {
    const want = predicted[`${id}-original`]
    const opens = windows && want.opensOnWindows === false ? false : want.opens
    expect(refused(`${id}-original`)).toBe(!opens)
    if (opens) expect((await readCells(saved(`${id}-original`), PREFIX_BOOKS.find((x) => x.id === id)!.read)).values).toEqual(want.values)
  })

  it('keeps the rules and the table repaired', async () => {
    const [a, b] = [await readXlsx(await repaired(prefixMade('rules'))), await readXlsx(saved('rules-repaired'))]
    const unstyled = (r: typeof a) => r.book.sheets[0].conditionalFormats?.map((rule) => ({ ...rule, style: undefined }))
    expect(unstyled(b)).toEqual(unstyled(a))
    expect(b.book.sheets[0].validations).toEqual(a.book.sheets[0].validations)
    expect((await readXlsx(saved('table-repaired'))).book.sheets[0].tables?.[0].columns[1].formula).toBe((await readXlsx(await repaired(prefixMade('table')))).book.sheets[0].tables?.[0].columns[1].formula)
  })
})
