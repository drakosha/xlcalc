/**
 * `writeKept` over a sheet whose conditional formats the model does not all read: the rules it does not have are
 * carried at their places among the workbook's, their ranges and formulas moved by operations of the structure as
 * Excel moves a rule's; a rule read without a part of its look keeps the file's differential format; what cannot go as
 * Excel would is refused (`rules`, `moved`). The books are the Excel step's (`bench/coverage/keep/rules.ts`).
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { createSession, type BookInput, type StructureOp } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { RULES_SETS, rulesFile, rulesOf, rulesWritten } from '../bench/coverage/keep/rules.js'
import { editParts } from '../bench/coverage/xlsx/sent.js'

const sheetXml = async (data: Uint8Array) => (await openArchive(data).text('xl/worksheets/sheet1.xml'))!
const written = async (data: Uint8Array, id: string) => {
  const w = await rulesWritten(data, RULES_SETS.rules.find((c) => c.id === id)!)
  if (!('data' in w)) throw new Error(w.refused)
  return w.data
}

/** The rules of the book `mixed` in their order, as `rulesOf` gives them. */
const MIXED = [
  'cellIs A2:A21 =100',
  'colorScale B2:B21 =$G$2 unread',
  'expression C2:C21 =C2>$G$1',
  'expression D2:D21 =MOD(D2,2)=0 gradient',
  'dataBar E2:E21 unread bar E2:E21',
  'dataBar F2:F21 bar F2:F21',
  "x14 colorScale H2:H21 ='P'!$A$1 unread",
]

describe('rules the model does not have, carried', () => {
  it('stay where they were among the rules a host adds, deletes and moves, their priorities counted with them', async () => {
    const data = await rulesFile('mixed')
    expect((await rulesOf(data)).rules).toEqual(MIXED)
    const after = async (id: string) => (await rulesOf(await written(data, id))).rules
    expect(await after('add')).toEqual([...MIXED, 'cellIs B2:B21 =150 stop'])
    expect(await after('delete-1')).toEqual(MIXED.slice(1))
    expect(await after('delete-3')).toEqual(MIXED.filter((_, i) => i !== 2))
    expect(await after('first-6')).toEqual([MIXED[5], ...MIXED.filter((_, i) => i !== 5)])
    expect(await after('last-1')).toEqual([...MIXED.slice(1), MIXED[0]])
    const out = await written(data, 'add')
    expect((await rulesOf(out)).priorities).toEqual([1, 2, 3, 4, 5, 6, 7, 8])
    // read back, the workbook has its rules and the added one; the file keeps the others for the next writer
    const back = await readXlsx(out)
    expect(back.book.sheets[0].conditionalFormats?.map((r) => r.ranges[0])).toEqual(['A2:A21', 'C2:C21', 'D2:D21', 'F2:F21', 'B2:B21'])
    expect(back.issues.filter((i) => i.code === 'conditionalFormat').map((i) => i.cells)).toEqual([['B2:B21', 'H2:H21'], ['D2:D21'], ['E2:E21', 'E2:E21']])
  })

  it('move with rows and columns as Excel moves a rule: inserted lines right after join it, its formulas follow, deleted cells go', async () => {
    const data = await rulesFile('mixed')
    const after = async (id: string) => (await rulesOf(await written(data, id))).rules
    expect(await after('insert-row-1')).toEqual(MIXED.map((r) => r.replace(/(\D)2:(\w)21/g, '$13:$222').replace('=$G$2', '=$G$3').replace('C2>$G$1', 'C3>$G$2').replace('MOD(D2', 'MOD(D3')))
    expect(await after('insert-row-22')).toEqual(MIXED.map((r) => r.replace(/21\b/g, '23')))
    // a colour scale and the data bars that lost rows go after the others, the read data bar among them in its order
    expect(await after('delete-rows-5')).toEqual([MIXED[0], MIXED[2], MIXED[3], MIXED[1], MIXED[4], MIXED[5], MIXED[6]].map((r) => r.replace(/21\b/g, '19')))
    expect((await after('delete-row-2'))[3]).toBe('colorScale B2:B20 =#REF! unread')
    expect(await after('delete-col-b')).not.toContainEqual(expect.stringContaining('colorScale B'))
    expect(await after('delete-rows-all')).toEqual([])
    expect((await after('rename-p')).at(-1)).toBe("x14 colorScale H2:H21 ='Q 2'!$A$1 unread")
    expect((await after('delete-sheet-p')).at(-1)).toBe('x14 colorScale H2:H21 =#REF! unread')
    expect((await after('move-b'))[1]).toBe('colorScale J2:J21 =$G$2 unread')
  })

  it('a value changed, or another sheet, leaves the sheet as the file has it', async () => {
    const data = await rulesFile('mixed')
    const s = createSession((await openKept(data)).book)
    s.setCells([{ sheet: 'S', cell: 'A2', input: { value: 1 } }])
    s.editStructure({ op: 'insertRows', sheet: 'P', row: 5 })
    const out = await writeKept(data, s.getSnapshot())
    const cf = (x: string) => x.slice(x.indexOf('<conditionalFormatting'))
    expect(cf(await sheetXml(out.data))).toBe(cf(await sheetXml(data)))
  })

  it("an Excel 2010 data bar's id the workbook's rules write is renamed in both its parts", async () => {
    // the id the writer gives the read data bar F (the sheet's rule 4) where E's carried one had it
    const taken = '{00000001-0000-4000-8000-000000000004}'
    const data = await editParts(await rulesFile('mixed'), (name, text) => (name === 'xl/worksheets/sheet1.xml' ? text.replaceAll('{6D2B1D5A-0B8C-4C1E-9A43-2F4E0D7A1B01}', taken) : undefined))
    const xml = await sheetXml(await written(data, 'add'))
    expect(xml.split(taken).length - 1).toBe(2)
    const fresh = /\{FFFFFFFF-0000-4000-8000-0{11}1\}/g
    expect(xml.match(fresh)).toHaveLength(2)
  })
})

/** A sheet S with a rule the model reads on A1:A5 and, put in by hand, `cf` and the extension list `ext`. */
async function handMade(cf: string, ext = '', book: Partial<BookInput> = {}): Promise<Uint8Array> {
  const cells = Object.fromEntries(Array.from({ length: 6 }, (_, r) => [`A${r + 1}`, { value: r }]))
  const { data } = await writeXlsx({ book: { sheets: [{ name: 'S', cells, conditionalFormats: [{ type: 'cellIs', ranges: ['A1:A5'], operator: 'greaterThan', formulas: ['2'], style: { font: { bold: true } } }] }, { name: 'T', cells: {} }], ...book } })
  return editParts(data, (name, text) => (name === 'xl/worksheets/sheet1.xml' ? text.replace('</conditionalFormatting>', `</conditionalFormatting>${cf}`).replace('</worksheet>', `${ext}</worksheet>`) : undefined))
}

const SCALE = (sqref: string, val: string) => `<conditionalFormatting sqref="${sqref}"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="num" val="${val}"/><color auto="1"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>`

async function refusal(data: Uint8Array, op: StructureOp | 'rule'): Promise<unknown> {
  const s = createSession((await openKept(data)).book)
  if (op === 'rule') s.setLayout([{ sheet: 'S', conditionalFormats: [] }])
  else expect(s.editStructure(op).ok).toBe(true)
  return writeKept(data, s.getSnapshot()).then(
    () => undefined,
    (e: Error & { code?: string }) => [e.name, e.code, e.message],
  )
}

describe('what a carried rule cannot do is refused', () => {
  it("a formula relative to the rule's first cell, which a deletion took: rules; one with the first cell kept moves", async () => {
    const data = await handMade(SCALE('B2:B5', 'B2*2'))
    expect(await refusal(data, { op: 'deleteRows', sheet: 'S', row: 2 })).toEqual(['KeptRefusal', 'rules', expect.stringContaining('lost its first cell')])
    const s = createSession((await openKept(data)).book)
    s.editStructure({ op: 'insertRows', sheet: 'S', row: 1 })
    expect(await sheetXml((await writeKept(data, s.getSnapshot())).data)).toContain('<conditionalFormatting sqref="B3:B6"><cfRule type="colorScale" priority="2"><colorScale><cfvo type="min"/><cfvo type="num" val="B3*2"/>')
  })

  it('a rule moved to another sheet with its cells: moved', async () => {
    const data = await handMade(SCALE('B2:B5', '$C$1'))
    expect(await refusal(data, { op: 'moveRange', from: { sheet: 'S', range: 'B1:B6' }, to: { sheet: 'T', cell: 'A1' } })).toEqual(['KeptRefusal', 'moved', expect.stringContaining('another sheet')])
  })

  it("a pivot table's rules, and a rule read over part of its ranges: rules, where the sheet's rules change", async () => {
    const pivot = await handMade('<conditionalFormatting sqref="C1:C3" pivot="1"><cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>')
    expect(await refusal(pivot, 'rule')).toEqual(['KeptRefusal', 'rules', expect.stringContaining("pivot table's rules")])
    const outside = await handMade('<conditionalFormatting sqref="C1:C3 XFE1"><cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan"><formula>1</formula></cfRule></conditionalFormatting>')
    expect(await refusal(outside, 'rule')).toEqual(['KeptRefusal', 'rules', expect.stringContaining('condRangeOutside')])
  })

  it('an unread rule whose references are not read refuses rows inserted on its sheet: unread', async () => {
    const data = await handMade(SCALE('B2:B5', 'SUM('))
    expect(await refusal(data, { op: 'insertRows', sheet: 'S', row: 1 })).toEqual(['KeptRefusal', 'unread', expect.stringContaining('unread conditional format')])
    expect(await refusal(data, 'rule')).toBeUndefined()
  })
})

describe("a rule read without a part of its look keeps the file's", () => {
  const FILLS = 'test/excel/cond/fills-held/fills.xlsx'

  it('its dxf, while its look in the workbook is the same; a new look is the workbook\'s', async () => {
    const data = new Uint8Array(readFileSync(FILLS))
    const dxfs = (xml: string) => [...xml.matchAll(/sqref="(\w+):\w+"><cfRule[^>]*?dxfId="(\d+)"/g)].map((m) => `${m[1]} ${m[2]}`)
    expect(dxfs(await sheetXml(data))).toEqual(['B1 6', 'B2 5', 'B3 4', 'B4 3', 'B5 2', 'B6 1', 'B7 0'])
    const s = createSession((await openKept(data)).book)
    s.editStructure({ op: 'insertRows', sheet: 'S', row: 1 })
    const rules = s.getBook().sheets[0].conditionalFormats!
    s.setLayout([{ sheet: 'S', conditionalFormats: rules.map((r, i) => (i === 1 ? ({ ...r, style: { fill: { color: '#00FF00' } } } as typeof r) : r)) }])
    const out = await writeKept(data, s.getSnapshot())
    // the rules over B2 to B4, B5 and B7 read their fills without a part; B3's look is new, B6's is read whole
    expect(dxfs(await sheetXml(out.data))).toEqual(['B2 6', 'B3 8', 'B4 4', 'B5 3', 'B6 7', 'B7 1', 'B8 0'])
  })
})
