import type { BookInput } from '../../../src/contract.js'
import { createSession } from '../../../src/index.js'
import { bookGeometry, columnPoints } from '../../../src/column-width.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import { editParts } from './sent.js'

interface DefaultWidthCase {
  id: string
  font?: BookInput['font']
  defaultColWidth?: number
  baseColWidth?: number
}

export const DEFAULT_WIDTH_CASES: readonly DefaultWidthCase[] = [
  { id: 'none' },
  { id: 'narrow', defaultColWidth: 4.5 },
  { id: 'fractional', defaultColWidth: 12.25 },
  { id: 'wide', defaultColWidth: 18.75 },
  { id: 'base', baseColWidth: 20 },
  { id: 'both', defaultColWidth: 4.5, baseColWidth: 20 },
  { id: 'arial-8-default', font: { name: 'Arial', size: 8 }, defaultColWidth: 4.5 },
  { id: 'arial-8-base', font: { name: 'Arial', size: 8 }, baseColWidth: 20 },
  { id: 'times-10-default', font: { name: 'Times New Roman', size: 10 }, defaultColWidth: 12.25 },
  { id: 'times-10-base', font: { name: 'Times New Roman', size: 10 }, baseColWidth: 20 },
]

const number = { value: 12345.67, numberFormat: '0.00' }
const BOOK: BookInput = { sheets: [{ name: 'S', cells: { A1: number, B1: number }, columns: { B: { width: 12 } } }] }

export async function defaultWidthFile({ id: _, font, ...attrs }: DefaultWidthCase): Promise<Uint8Array> {
  const { data } = await writeXlsx({ book: { ...BOOK, font } })
  if (!Object.keys(attrs).length) return data
  const props = Object.entries(attrs).map(([name, value]) => ` ${name}="${value}"`).join('')
  return editParts(data, (part, xml) => part === 'xl/worksheets/sheet1.xml' ? xml.replace('<sheetFormatPr ', `<sheetFormatPr${props} `) : undefined)
}

export async function defaultWidthPrediction(c: DefaultWidthCase) {
  const { book, issues } = await readXlsx(await defaultWidthFile(c))
  const [digit] = bookGeometry(book.font)
  const session = createSession(book)
  const result = session.getViewport('S', 'A1:B1')
  session.close()
  return { points: result.columns.map(({ width }) => columnPoints(width, digit)), shown: [result.cells.A1.text, result.cells.B1.text], issues: issues.map(({ code }) => code) }
}
