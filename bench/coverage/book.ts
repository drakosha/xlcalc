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

/** Spill references (`E2#`) as a file writes them, `_xlfn.ANCHORARRAY(E2)`. */
export const anchorArray = (f: string) => f.replace(/((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?[A-Z]{1,3}\$?\d+#/g, (m) => `_xlfn.ANCHORARRAY(${m.slice(0, -1)})`)

/** The formulas of the read columns over a cell `x`, which never give an error: its type code, its number, its text. */
export const READ_COLUMNS = {
  type: (x: string) => `IF(ISERROR(${x}),ERROR.TYPE(${x}),IF(ISNUMBER(${x}),0,IF(ISLOGICAL(${x}),IF(${x},-1,-2),IF(ISTEXT(${x}),-3,-4))))`,
  num: (x: string) => `IF(ISNUMBER(${x}),${x},0)`,
  text: (x: string) => `IF(ISTEXT(${x}),${x},"")`,
}

/**
 * `sheets`: further data sheets by name, values or formulas by A1 address (the probes of `probes.ts`, the case sets).
 * `legacy`: the case formulas are written as formulas of the old kind (no array, no dynamic metadata), which Excel
 * 365 computes with implicit intersection (the sets of `legacy/`).
 */
export async function writeBook(cases: Pick<Case, 'id' | 'formula'>[], file: string, sheets: SheetCells = {}, legacy = false): Promise<void> {
  const wb = new ExcelJS.Workbook()
  for (const [name, cells] of Object.entries(sheets)) {
    const ws = wb.addWorksheet(name)
    for (const [addr, v] of Object.entries(cells)) ws.getCell(addr).value = (typeof v === 'object' ? { formula: anchorArray(v.formula), ...(v.dynamic ? { shareType: 'array', ref: addr } : {}) } : v) as ExcelJS.CellValue
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
    c.getCell(`A${r}`).value = (legacy ? { formula: anchorArray(x.formula) } : { formula: anchorArray(x.formula), shareType: 'array', ref: `A${r}` }) as ExcelJS.CellFormulaValue
    c.getCell(`B${r}`).value = { formula: READ_COLUMNS.type(`A${r}`) } as ExcelJS.CellFormulaValue
    c.getCell(`C${r}`).value = { formula: READ_COLUMNS.num(`A${r}`) } as ExcelJS.CellFormulaValue
    c.getCell(`D${r}`).value = { formula: READ_COLUMNS.text(`A${r}`) } as ExcelJS.CellFormulaValue
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
 * `cm="1"` points at the XLDAPR metadata, as Excel 365 (and XlsxWriter) write them; but those `old` says are array
 * formulas of the old kind (sheet index from 0, address), which stay as they are.
 */
export async function markDynamic(file: string, sheets: number, old: (sheet: number, addr: string) => boolean = () => false): Promise<void> {
  const zip = await JSZip.loadAsync(readFileSync(file))
  for (let i = 1; i <= sheets; i++) {
    const path = `xl/worksheets/sheet${i}.xml`
    const xml = await zip.file(path)!.async('string')
    zip.file(path, xml.replace(/<c r="([A-Z]+\d+)"([^>]*)><f t="array"/g, (m, addr: string, rest: string) => (old(i - 1, addr) ? m : `<c r="${addr}"${rest} cm="1"><f t="array"`)))
  }
  zip.file('xl/metadata.xml', METADATA)
  const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string')
  zip.file('xl/_rels/workbook.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdMeta" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>'))
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>'))
  writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}
