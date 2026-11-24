/** Tables without an app part, with a namespace prefix on the table part. */
import type { BookInput } from '../../../src/contract.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'
import { editParts } from './sent.js'

export const TABLE_STYLE_CASES = [
  { id: 'no-style' },
  { id: 'default', name: 'TableStyleMedium2' },
  { id: 'light-1', name: 'TableStyleLight1' },
  { id: 'light', name: 'TableStyleLight' },
  { id: 'medium', name: 'TableStyleMedium' },
  { id: 'dark', name: 'TableStyleDark' },
  { id: 'custom-missing', name: 'UnregisteredStyle' },
  { id: 'custom-light', name: 'TableStyleLight', defined: true },
] as const

export type TableStyleCase = (typeof TABLE_STYLE_CASES)[number]

const BOOK: BookInput = {
  sheets: [{
    name: 'S',
    cells: {
      A1: { value: 'Name' }, B1: { value: 'Amount' },
      A2: { value: 'One' }, B2: { value: 1 },
      A3: { value: 'Two' }, B3: { value: 2 },
    },
    tables: [{ name: 'Amounts', range: 'A1:B3', columns: [{ name: 'Name' }, { name: 'Amount' }], style: { rowStripes: true } }],
  }],
}

const CUSTOM_STYLE = '<dxfs count="1"><dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>'
  + '<tableStyles count="1"><tableStyle name="TableStyleLight" pivot="0" table="1" count="1"><tableStyleElement type="headerRow" dxfId="0"/></tableStyle></tableStyles>'

export async function tableStyleFile(c: TableStyleCase): Promise<Uint8Array> {
  const { data } = await writeXlsx({ book: BOOK }, { values: 'omit' })
  return editParts(data, (part, xml) => {
    if (part === 'xl/styles.xml' && 'defined' in c) return xml.replace('</styleSheet>', CUSTOM_STYLE + '</styleSheet>')
    if (part !== 'xl/tables/table1.xml') return undefined
    const withName = 'name' in c ? xml.replace('<tableStyleInfo', `<tableStyleInfo name="${c.name}"`) : xml
    return withName.replace(' xmlns="', ' xmlns:x="').replace(/(<\/?)([A-Za-z]\w*)(?=[\s/>])/g, '$1x:$2')
  })
}

/** What readXlsx currently takes of a case. */
export async function tableStylePrediction(c: TableStyleCase) {
  const { book, issues } = await readXlsx(await tableStyleFile(c))
  return { style: book.sheets[0].tables?.[0]?.style ?? null, issues: issues.map(({ code, message }) => ({ code, message })) }
}
