/**
 * Синтетическая книга замера в xlsx (ExcelJS, без кэша значений): лист данных `D` и лист случаев `C`.
 * В строке случая: A формула случая (у ответа-массива формула массива), B код типа ответа, C число,
 * D текст, E id случая. B, C и D никогда не дают ошибку, поэтому Excel отдаёт их диапазоном за один
 * вызов без потери позиций (ошибки из `value` диапазона выпадают из списка).
 *
 * Код типа в B: 0 число, -1 TRUE, -2 FALSE, -3 текст, -4 прочее, >0 номер ошибки ERROR.TYPE.
 *
 * С `dynamic` формулы случаев пишутся формулами динамического массива (как их сохраняет Excel 365):
 * `<c cm="1"><f t="array" ref="A1">` и метаданные XLDAPR с `fDynamic="1"` в `xl/metadata.xml`. Такая
 * формула считается без неявного пересечения, а ответ-массив растекается; соседние случаи и колонки
 * B…E заняты, и под последним случаем стоит заглушка, так что ответ больше одной ячейки всегда `#SPILL!`.
 */
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { readFileSync, writeFileSync } from 'node:fs'
import type { Case } from './cases.js'
import { DATA_SHEET, dataCells } from './data.js'
import type { SheetCells } from './probes.js'

export const CASE_SHEET = 'C'

export interface BookOptions {
  /** Формулы случаев как формулы динамического массива (Excel 365), а не старого вида. */
  dynamic?: boolean
}

/** `sheets`: further data sheets by name, values or formulas by A1 address (the probes of `probes.ts`, the case sets). */
export async function writeBook(cases: Pick<Case, 'id' | 'formula' | 'pick'>[], file: string, sheets: SheetCells = {}, opts: BookOptions = {}): Promise<void> {
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
    // ответ-массив (INDEX/ROWS/COLUMNS над функцией-массивом): формула массива, как ввод Ctrl+Shift+Enter;
    // в обычной формуле старого вида Excel применяет к аргументу TRANSPOSE неявное пересечение (#VALUE!)
    const array = x.pick !== undefined || opts.dynamic
    c.getCell(`A${r}`).value = (array ? { formula: x.formula, shareType: 'array', ref: `A${r}` } : { formula: x.formula }) as ExcelJS.CellFormulaValue
    c.getCell(`B${r}`).value = { formula: `IF(ISERROR(A${r}),ERROR.TYPE(A${r}),IF(ISNUMBER(A${r}),0,IF(ISLOGICAL(A${r}),IF(A${r},-1,-2),IF(ISTEXT(A${r}),-3,-4))))` } as ExcelJS.CellFormulaValue
    c.getCell(`C${r}`).value = { formula: `IF(ISNUMBER(A${r}),A${r},0)` } as ExcelJS.CellFormulaValue
    c.getCell(`D${r}`).value = { formula: `IF(ISTEXT(A${r}),A${r},"")` } as ExcelJS.CellFormulaValue
    c.getCell(`E${r}`).value = x.id
  })
  // заглушка под последним случаем: ответ-массив не растекается и там
  if (opts.dynamic) c.getCell(`A${cases.length + 1}`).value = 'end'
  await wb.xlsx.writeFile(file)
  if (opts.dynamic) await markDynamic(file, wb.worksheets.length)
}

const METADATA = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`

/**
 * Формулы массива колонки A листа случаев (последний лист, `sheetN.xml` у ExcelJS) становятся формулами
 * динамического массива: `cm="1"` ссылается на метаданные XLDAPR, как пишет их Excel 365 (и XlsxWriter).
 */
async function markDynamic(file: string, sheetCount: number): Promise<void> {
  const zip = await JSZip.loadAsync(readFileSync(file))
  const path = `xl/worksheets/sheet${sheetCount}.xml`
  const sheet = await zip.file(path)!.async('string')
  if (!sheet.includes('<f t="array"')) throw new Error(`${path}: нет формул массива, лист случаев не найден`)
  zip.file(path, sheet.replace(/<c r="(A\d+)"([^>]*)><f t="array"/g, '<c r="$1"$2 cm="1"><f t="array"'))
  zip.file('xl/metadata.xml', METADATA)
  const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('string')
  zip.file('xl/_rels/workbook.xml.rels', rels.replace('</Relationships>', '<Relationship Id="rIdMeta" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>'))
  const types = await zip.file('[Content_Types].xml')!.async('string')
  zip.file('[Content_Types].xml', types.replace('</Types>', '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>'))
  writeFileSync(file, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }))
}
