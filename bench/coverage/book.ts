/**
 * Синтетическая книга замера в xlsx (ExcelJS, без кэша значений): лист данных `D` и лист случаев `C`.
 * В строке случая: A формула случая (у ответа-массива формула массива), B код типа ответа, C число,
 * D текст, E id случая. B, C и D никогда не дают ошибку, поэтому Excel отдаёт их диапазоном за один
 * вызов без потери позиций (ошибки из `value` диапазона выпадают из списка).
 *
 * Код типа в B: 0 число, -1 TRUE, -2 FALSE, -3 текст, -4 прочее, >0 номер ошибки ERROR.TYPE.
 */
import ExcelJS from 'exceljs'
import type { Case } from './cases.js'
import { DATA_SHEET, dataCells } from './data.js'
import type { SheetCells } from './probes.js'

export const CASE_SHEET = 'C'

/** `sheets`: further data sheets by name, values or formulas by A1 address (the probes of `probes.ts`, the case sets). */
export async function writeBook(cases: Pick<Case, 'id' | 'formula' | 'pick'>[], file: string, sheets: SheetCells = {}): Promise<void> {
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
    c.getCell(`A${r}`).value = (x.pick === undefined ? { formula: x.formula } : { formula: x.formula, shareType: 'array', ref: `A${r}` }) as ExcelJS.CellFormulaValue
    c.getCell(`B${r}`).value = { formula: `IF(ISERROR(A${r}),ERROR.TYPE(A${r}),IF(ISNUMBER(A${r}),0,IF(ISLOGICAL(A${r}),IF(A${r},-1,-2),IF(ISTEXT(A${r}),-3,-4))))` } as ExcelJS.CellFormulaValue
    c.getCell(`C${r}`).value = { formula: `IF(ISNUMBER(A${r}),A${r},0)` } as ExcelJS.CellFormulaValue
    c.getCell(`D${r}`).value = { formula: `IF(ISTEXT(A${r}),A${r},"")` } as ExcelJS.CellFormulaValue
    c.getCell(`E${r}`).value = x.id
  })
  await wb.xlsx.writeFile(file)
}
