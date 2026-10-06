/**
 * The measurement's synthetic book as xlsx (ExcelJS, no cached values): the data sheet `D` and the case sheet `C`.
 * A case's row: A the case formula, B the answer's type code, C the number, D the text, E the case id. B, C and D
 * never give an error, so Excel returns them as a range in one call without shifting positions (errors drop out
 * of a range's `value` list).
 *
 * Type code in B: 0 number, -1 TRUE, -2 FALSE, -3 text, -4 other, >0 the ERROR.TYPE number.
 *
 * Case formulas are dynamic array formulas, as Excel 365 saves them: `<c cm="1"><f t="array" ref="A1">` and the
 * XLDAPR metadata with `fDynamic="1"` in `xl/metadata.xml`. Such a formula is computed without implicit
 * intersection and an array answer spills; the neighbouring cases and columns B…E are taken and a stub stands
 * under the last case, so an answer larger than one cell is always `#SPILL!`.
 */
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { readFileSync, writeFileSync } from 'node:fs'
import type { Case } from './cases.js'
import { DATA_SHEET, dataCells } from './data.js'
import type { SheetCells } from './probes.js'

export const CASE_SHEET = 'C'

/** `sheets`: further data sheets by name, values or formulas by A1 address (the probes of `probes.ts`, the case sets). */
export async function writeBook(cases: Pick<Case, 'id' | 'formula'>[], file: string, sheets: SheetCells = {}): Promise<void> {
  const wb = new ExcelJS.Workbook()
  for (const [name, cells] of Object.entries(sheets)) {
    const ws = wb.addWorksheet(name)
    for (const [addr, v] of Object.entries(cells)) ws.getCell(addr).value = (typeof v === 'object' ? { formula: v.formula } : v) as ExcelJS.CellValue
  }
  const d = wb.addWorksheet(DATA_SHEET)
  for (const [addr, v] of dataCells()) {
    const cell = d.getCell(addr)
    if (typeof v !== 'object') cell.value = v
    else if ('date' in v) {
      cell.value = v.date
      cell.numFmt = 'yyyy-mm-dd'
    } else cell.value = { formula: v.formula } as ExcelJS.CellFormulaValue
  }
  const c = wb.addWorksheet(CASE_SHEET)
  cases.forEach((x, i) => {
    const r = i + 1
    c.getCell(`A${r}`).value = { formula: x.formula, shareType: 'array', ref: `A${r}` } as ExcelJS.CellFormulaValue
    c.getCell(`B${r}`).value = { formula: `IF(ISERROR(A${r}),ERROR.TYPE(A${r}),IF(ISNUMBER(A${r}),0,IF(ISLOGICAL(A${r}),IF(A${r},-1,-2),IF(ISTEXT(A${r}),-3,-4))))` } as ExcelJS.CellFormulaValue
    c.getCell(`C${r}`).value = { formula: `IF(ISNUMBER(A${r}),A${r},0)` } as ExcelJS.CellFormulaValue
    c.getCell(`D${r}`).value = { formula: `IF(ISTEXT(A${r}),A${r},"")` } as ExcelJS.CellFormulaValue
    c.getCell(`E${r}`).value = x.id
  })
  // a stub under the last case: an array answer does not spill there either
  c.getCell(`A${cases.length + 1}`).value = 'end'
  await wb.xlsx.writeFile(file)
  await markDynamic(file, wb.worksheets.length)
}

const METADATA = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`

/**
 * The array formulas of the first `sheets` sheets (`sheetN.xml` in ExcelJS's order) become dynamic array formulas:
 * `cm="1"` points at the XLDAPR metadata, as Excel 365 (and XlsxWriter) write them.
 */
export async function markDynamic(file: string, sheets: number): Promise<void> {
  const zip = await JSZip.loadAsync(readFileSync(file))
  for (let i = 1; i <= sheets; i++) {
    const path = `xl/worksheets/sheet${i}.xml`
    const xml = await zip.file(path)!.async('string')
    zip.file(path, xml.replace(/<c r="([A-Z]+\d+)"([^>]*)><f t="array"/g, '<c r="$1"$2 cm="1"><f t="array"'))
  }
  zip.file('xl/metadata.xml', METADATA)
  const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string')
  zip.file('xl/_rels/workbook.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdMeta" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>'))
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>'))
  writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}
