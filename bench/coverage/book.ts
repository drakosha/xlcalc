/**
 * Синтетическая книга замера в xlsx (ExcelJS, без кэша значений): лист данных `D` и лист случаев `C`.
 * В строке случая: A формула случая, B код типа ответа, C число, D текст, E id случая. B, C и D
 * никогда не дают ошибку, поэтому Excel отдаёт их диапазоном за один вызов без потери позиций
 * (ошибки из `value` диапазона выпадают из списка).
 *
 * Код типа в B: 0 число, -1 TRUE, -2 FALSE, -3 текст, -4 прочее, >0 номер ошибки ERROR.TYPE.
 */
import ExcelJS from 'exceljs'
import type { Case } from './cases.js'
import { DATA_SHEET, dataCells } from './data.js'

export const CASE_SHEET = 'C'

export async function writeBook(cases: Case[], file: string): Promise<void> {
  const wb = new ExcelJS.Workbook()
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
    c.getCell(`A${r}`).value = { formula: x.formula } as ExcelJS.CellFormulaValue
    c.getCell(`B${r}`).value = { formula: `IF(ISERROR(A${r}),ERROR.TYPE(A${r}),IF(ISNUMBER(A${r}),0,IF(ISLOGICAL(A${r}),IF(A${r},-1,-2),IF(ISTEXT(A${r}),-3,-4))))` } as ExcelJS.CellFormulaValue
    c.getCell(`C${r}`).value = { formula: `IF(ISNUMBER(A${r}),A${r},0)` } as ExcelJS.CellFormulaValue
    c.getCell(`D${r}`).value = { formula: `IF(ISTEXT(A${r}),A${r},"")` } as ExcelJS.CellFormulaValue
    c.getCell(`E${r}`).value = x.id
  })
  await wb.xlsx.writeFile(file)
}
