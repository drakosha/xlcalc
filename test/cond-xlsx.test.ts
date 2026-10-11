/**
 * Conditional formats in xlsx: every kind written and read back as the book holds it, Excel 2010's rules in the sheet's
 * extension list, differential formats as Excel keeps them, what the model cannot hold reported, and a file Excel for
 * Windows saved with the rules it made itself.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { BookInput, ConditionalFormat, DataValidation } from '../src/contract.js'
import { editStructure, validateFormula } from '../src/index.js'
import { buildCondCases as buildBoth } from '../bench/coverage/cond/both.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { codes, read } from './xlsx-fixtures.js'

const RED = { fill: { color: '#FFC7CE' }, font: { color: '#9C0006', bold: false } }

const RULES: ConditionalFormat[] = [
  { type: 'cellIs', ranges: ['A1:A10', 'C1:C10'], operator: 'between', formulas: ['1', '$B$1'], style: { ...RED, format: '0.0%' }, stopIfTrue: true },
  { type: 'expression', ranges: ['B2:B10'], formula: 'XLOOKUP(1,A1:A3,A1:A3)>5', style: { font: { bold: true, underline: 'double', strike: true }, border: { bottom: { style: 'thin', color: '#FF0000' } } } },
  { type: 'containsText', ranges: ['D1:D10'], text: 'a"b', style: { fill: { color: '#00FF00' } } },
  { type: 'notContainsText', ranges: ['D1:D10'], text: 'q', formula: 'ISERROR(SEARCH("z",D1))' },
  { type: 'timePeriod', ranges: ['E1:E10'], period: 'lastWeek', style: { format: 14 } },
  { type: 'containsBlanks', ranges: ['E1:E10'], style: RED },
  { type: 'top10', ranges: ['F1:F10'], rank: 3, percent: true, bottom: true, style: RED },
  { type: 'aboveAverage', ranges: ['F1:F10'], below: true, stdDev: 2, style: RED },
  { type: 'duplicateValues', ranges: ['G:G'], style: RED },
  { type: 'colorScale', ranges: ['H1:H10'], points: [{ value: { type: 'min' }, color: '#F8696B' }, { value: { type: 'percentile', value: '50' }, color: '#FFEB84' }, { value: { type: 'formula', value: '$B$1*2' }, color: '#63BE7B' }] },
  { type: 'dataBar', ranges: ['I1:I10'], min: { type: 'autoMin' }, max: { type: 'num', value: '100' }, color: '#638EC6', border: '#638EC6', negativeColor: '#FF0000', negativeBorder: '#C00000', axis: 'middle', axisColor: '#7F7F7F', direction: 'rightToLeft', gradient: false, showValue: false },
  { type: 'dataBar', ranges: ['J1:J10'], min: { type: 'min' }, max: { type: 'max' }, color: '#FF0000', minLength: 20, excel2007: true },
  { type: 'iconSet', ranges: ['K1:K10'], icons: '3Stars', points: [{ value: { type: 'percent', value: '0' } }, { value: { type: 'percent', value: '33' }, gte: false }, { value: { type: 'percent', value: '67' } }], reverse: true },
  { type: 'iconSet', ranges: ['L1:L10'], icons: '3Arrows', points: [{ value: { type: 'percent', value: '0' }, icon: { set: '3Flags', index: 0 } }, { value: { type: 'num', value: '5' }, icon: null }, { value: { type: 'percent', value: '67' } }], showValue: false },
  { type: 'expression', ranges: ['M1'], formula: "'T T'!$A$1>1", style: { font: { italic: true } } },
  { type: 'cellIs', ranges: ['N1'], operator: 'lessThan', formulas: ["'T T'!$A$2"], style: RED },
]

const BOOK: BookInput = { sheets: [{ name: 'S', cells: { A1: { value: 1 } }, conditionalFormats: RULES }, { name: 'T T', cells: {} }] }

async function partOf(data: Uint8Array, name: string): Promise<string> {
  return (await openArchive(data).text(name))!
}

describe('conditional formats in xlsx', () => {
  it('every kind and option comes back as written, without issues', async () => {
    const { data, issues } = await writeXlsx(BOOK)
    expect(issues).toEqual([])
    const back = await readXlsx(data)
    expect(back.issues).toEqual([])
    expect(back.book.sheets[0].conditionalFormats).toEqual(RULES)
    // and again: the file the reader's book writes reads the same
    expect((await readXlsx((await writeXlsx(back.book)).data)).book).toEqual(back.book)
  })

  it("Excel 2010's rules go into the extension list: a bar's part joined by its id, icon sets 2007 lacks, custom icons, other sheets' references", async () => {
    const xml = await partOf((await writeXlsx(BOOK)).data, 'xl/worksheets/sheet1.xml')
    const ext = xml.slice(xml.indexOf('<extLst><ext uri="{78C0D931'))
    expect(ext).toMatch(/<x14:cfRule type="dataBar" id="(\{[0-9A-F-]+\})">/)
    const id = /<x14:cfRule type="dataBar" id="(\{[0-9A-F-]+\})">/.exec(ext)![1]
    expect(xml).toContain(`<x14:id>${id}</x14:id>`)
    expect(ext).toContain('<x14:iconSet iconSet="3Stars" reverse="1">')
    expect(ext).toContain('<x14:cfIcon iconSet="NoIcons" iconId="0"/>')
    expect(ext).toContain("<xm:f>'T T'!$A$1&gt;1</xm:f><x14:dxf><font><i/></font></x14:dxf>")
    expect(xml.indexOf('<conditionalFormatting')).toBeLessThan(xml.indexOf('<extLst>'))
    expect(await partOf((await writeXlsx(BOOK)).data, 'xl/styles.xml')).toContain('<dxf><font><b val="0"/><color rgb="FF9C0006"/></font><numFmt numFmtId="164" formatCode="0.0%"/><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>')
    // a code every locale shows as a built-in format is written by its number and comes back as that number, as a cell's
    const pct: BookInput = { sheets: [{ name: 'S', cells: {}, conditionalFormats: [{ type: 'expression', ranges: ['A1'], formula: 'TRUE', style: { format: '0.00%' } }] }] }
    expect((await readXlsx((await writeXlsx(pct)).data)).book.sheets[0].conditionalFormats![0]).toMatchObject({ style: { format: 10 } })
  })

  it("a sheet with conditional formats and data validation keeps both: the extension list of a data bar is not the sheet's", async () => {
    const book: BookInput = {
      sheets: [
        {
          name: 'S',
          cells: { A1: { value: 1 } },
          conditionalFormats: [RULES[10], RULES[14]],
          validations: [{ range: 'D1:D10', type: 'list', formula1: "'T T'!$A$1:$A$3" }, { range: 'A1:A10', type: 'whole', formula1: '1', formula2: '100' }],
        },
        { name: 'T T', cells: {} },
      ],
    }
    const { data, issues } = await writeXlsx(book)
    expect(issues).toEqual([])
    const xml = await partOf(data, 'xl/worksheets/sheet1.xml')
    expect(xml.indexOf('<conditionalFormatting')).toBeLessThan(xml.indexOf('<dataValidations'))
    const back = await readXlsx(data)
    expect(back.issues).toEqual([])
    expect(back.book.sheets[0].conditionalFormats).toEqual(book.sheets[0].conditionalFormats)
    expect(back.book.sheets[0].validations).toEqual(book.sheets[0].validations)
  })

  it('a rule whose formula the file cannot hold is left out and reported', async () => {
    const book: BookInput = { sheets: [{ name: 'S', cells: {}, conditionalFormats: [{ type: 'expression', ranges: ['A1'], formula: 'SUM(', style: RED }, { type: 'expression', ranges: ['A2'], formula: "'Nope'!A1>1" }] }] }
    const { data, issues } = await writeXlsx(book)
    expect(codes(issues)).toEqual(['conditionalFormat A1 A2'])
    expect((await readXlsx(data)).book.sheets[0].conditionalFormats).toBeUndefined()
  })

  it('reads dxfs as Excel saves them, sqref with $ and whole columns, and reports what the model cannot hold', async () => {
    const styles = '<dxfs count="3"><dxf><numFmt numFmtId="14" formatCode="0.00%"/><fill><patternFill patternType="solid"><fgColor rgb="FF00FF00"/><bgColor rgb="FFFF0000"/></patternFill></fill></dxf><dxf><fill><patternFill patternType="gray125"><bgColor rgb="FFFF0000"/></patternFill></fill><font><sz val="14"/></font></dxf><dxf><border><left style="thin"><color auto="1"/></left></border></dxf></dxfs>'
    const body =
      '<sheetData/>' +
      '<conditionalFormatting sqref="$A$1:$A$1048576 B2"><cfRule type="cellIs" dxfId="0" priority="2" operator="equal"><formula>1</formula></cfRule><cfRule type="expression" dxfId="1" priority="1"><formula>TRUE</formula></cfRule></conditionalFormatting>' +
      '<conditionalFormatting sqref="C1"><cfRule type="aboveAverage" dxfId="2" priority="3" aboveAverage="0" equalAverage="1" stdDev="1"/><cfRule type="unknownKind" priority="4"/></conditionalFormatting>' +
      '<conditionalFormatting sqref="D1" pivot="1"><cfRule type="expression" priority="5"><formula>1</formula></cfRule></conditionalFormatting>'
    const r = await read({ sheets: [{ name: 'S', body }], styles })
    expect(r.book.sheets[0].conditionalFormats).toEqual([
      { ranges: ['A:A', 'B2'], type: 'expression', formula: 'TRUE' },
      { ranges: ['A:A', 'B2'], type: 'cellIs', operator: 'equal', formulas: ['1'], style: { format: '0.00%', fill: { color: '#FF0000' } } },
      { ranges: ['C1'], type: 'aboveAverage', below: true, equal: true, style: { border: { left: { style: 'thin' } } } },
    ])
    expect(r.issues.map((i) => i.message)).toEqual([
      "a conditional format's font sz is not in the model",
      "a conditional format's pattern fill is not in the model",
      'an average rule with both equal and standard deviations is read with equal alone, as Excel keeps it',
      'a conditional format of type unknownKind is not in the model',
      "a pivot table's conditional formatting is not in the model",
    ])
  })

  it('the files both Excels saved of a sheet with conditional formats, data validation and grouped rows hold what the engine holds', async () => {
    const built = buildBoth()
    const canon = (f: string | undefined) => (f === undefined ? f : (validateFormula(f) as { canonical: string }).canonical)
    const norm = (vs: DataValidation[] = []) => vs.map((x) => ({ ...x, operator: x.operator ?? 'between', formula1: canon(x.formula1), formula2: canon(x.formula2) })).sort((a, b) => (a.range < b.range ? -1 : 1))
    for (const dir of ['./excel/cond/both/', './excel/win/cond/both/'])
      for (const c of built.cases) {
        const r = await readXlsx(new Uint8Array(readFileSync(new URL(`${dir}${c.id}.xlsx`, import.meta.url))))
        const r2 = c.op ? editStructure(built.books.both, c.op) : { ok: true as const, book: built.books.both }
        if (!r2.ok) throw new Error(r2.code)
        expect(r.issues, `${dir}${c.id}`).toEqual([])
        expect(norm(r.book.sheets[0].validations), `${dir}${c.id}`).toEqual(norm(r2.book.sheets[0].validations))
        expect(r.book.sheets[0].rows, `${dir}${c.id}`).toEqual(r2.book.sheets[0].rows)
      }
  })

  it("the rules Excel for Windows made through COM read from the file it saved, with Excel's defaults", async () => {
    const r = await readXlsx(new Uint8Array(readFileSync(new URL('./excel/win/cond/made/made.xlsx', import.meta.url))))
    const cf = r.book.sheets[0].conditionalFormats!
    expect(cf.map((x) => x.type)).toEqual(['cellIs', 'cellIs', 'expression', 'containsText', 'containsBlanks', 'containsErrors', 'timePeriod', 'top10', 'aboveAverage', 'uniqueValues', 'colorScale', 'colorScale', 'dataBar', 'iconSet', 'expression'])
    expect(cf.find((x) => x.type === 'timePeriod')).toEqual({ ranges: ['H1:H12'], stopIfTrue: true, type: 'timePeriod', period: 'yesterday', style: { font: { bold: true, italic: false, color: '#9C0006' }, fill: { color: '#FFC7CE' } } })
    expect(cf.at(-1)).toMatchObject({ ranges: ['P1:P5', 'R1:R5'], formula: 'P1>2' })
    expect(r.issues.filter((i) => i.code === 'conditionalFormat')).toEqual([])
  })
})
