/** Record the 'empty' case after accepting the repair prompt and saving it again in Excel for Mac. */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { readXlsx } from '../../../src/xlsx.js'
import { excelVersion } from '../excel.js'
import { ROOT } from '../paths.js'
import { writeAnswers } from '../stored.js'

const file = process.argv[2]
if (!file) throw new Error('Give the file Excel saved after repairing the empty case')
const { book, cached } = await readXlsx(readFileSync(file))
const value = cached.sheets[0]?.cells.B1?.value
if (value === undefined) throw new Error('The repaired workbook did not save its calculated value at B1')
writeAnswers(join(ROOT, 'test/excel/xlsx/empty-defined-names-repair.json'), {
  meta: { version: excelVersion(), locale: 'ru-RU', mode: 'interactive repair' },
  answers: { empty: { names: book.names ?? [], value } },
})
