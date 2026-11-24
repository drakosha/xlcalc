/**
 * Custom table styles with what a cell style cannot hold: inside borders of a region, patterned fills and number
 * formats. The book is written with marker fills, then each marker's differential format is replaced by hand, as a
 * file from Excel holds it.
 */
import type { BookInput } from '../../../src/contract.js'
import { registerTableStyles } from '../../../src/table-styles.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { editParts } from '../xlsx/sent.js'

export const INSIDE_CASES = [{ id: 'inside', sheet: 'S', range: 'A1:I10' }] as const

const black = '<color rgb="FF000000"/>'
/** Marker fill colour (its last byte) and the differential format put in its place. */
const FORMATS: Record<string, string> = {
  '01': `<dxf><border><left style="thin">${black}</left><right style="thin">${black}</right><top style="thin">${black}</top><bottom style="thin">${black}</bottom><vertical style="dashed"><color rgb="FF0000FF"/></vertical><horizontal style="thin"><color rgb="FFFF0000"/></horizontal></border></dxf>`,
  '02': '<dxf><fill><patternFill patternType="lightUp"><fgColor rgb="FFFF0000"/><bgColor rgb="FFFFFF00"/></patternFill></fill></dxf>',
  '03': '<dxf><fill><patternFill patternType="darkGrid"><fgColor rgb="FF0000FF"/></patternFill></fill></dxf>',
  '04': '<dxf><fill><patternFill patternType="gray125"><bgColor rgb="FF00FF00"/></patternFill></fill></dxf>',
  '05': '<dxf><numFmt numFmtId="164" formatCode="0.00%"/><fill><patternFill patternType="solid"><bgColor rgb="FFDDDDDD"/></patternFill></fill></dxf>',
  '06': '<dxf><font><b/></font><numFmt numFmtId="165" formatCode="#,##0.000"/></dxf>',
  '07': '<dxf><border><horizontal style="thin"><color rgb="FF00B050"/></horizontal><vertical style="medium"><color rgb="FF7030A0"/></vertical></border></dxf>',
  '08': '<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFFC000"/><bgColor rgb="FFFFC000"/></patternFill></fill></dxf>',
}
const marker = (n: string) => ({ fill: { color: `#0101${n}` } })

const BOOK: BookInput = {
  tableStyles: {
    custom: {
      Probe: {
        table: true, pivot: false, elements: [
          { type: 'wholeTable', style: marker('01') },
          { type: 'headerRow', style: marker('02') },
          { type: 'firstRowStripe', style: marker('03') },
          { type: 'secondRowStripe', style: marker('04') },
          { type: 'totalRow', style: marker('05') },
          { type: 'firstColumn', style: marker('06') },
        ],
      },
      Pairs: {
        table: true, pivot: false, elements: [
          { type: 'firstRowStripe', size: 2, style: marker('07') },
          { type: 'secondRowStripe', size: 2, style: marker('08') },
        ],
      },
    },
  },
  sheets: [{
    name: 'S',
    cells: {
      B2: { value: 'Name' }, C2: { value: 'Qty' }, D2: { value: 'Share' }, E2: { value: 'Note' },
      B3: { value: 1234.5 }, C3: { value: 3 }, D3: { value: 0.25 }, E3: { value: 'a' },
      B4: { value: 2 }, C4: { value: 4 }, D4: { value: 0.5 },
      B5: { value: 3 }, C5: { value: 5 }, D5: { value: 0.75 },
      B6: { value: 4 }, C6: { value: 6 },
      B7: { value: 5 }, C7: { value: 7 },
      B8: { value: 'Total' }, C8: { value: 25 }, D8: { value: 0.5 },
      G2: { value: 'K' }, H2: { value: 'V' },
      G3: { value: 1 }, H3: { value: 1 }, G4: { value: 2 }, G5: { value: 3 }, G6: { value: 4 }, G7: { value: 5 }, G8: { value: 6 }, G9: { value: 7 },
    },
    tables: [
      {
        name: 'Probe1', range: 'B2:E8', totals: true,
        columns: [{ name: 'Name' }, { name: 'Qty' }, { name: 'Share' }, { name: 'Note' }],
        style: { name: 'Probe', rowStripes: true, firstColumn: true },
      },
      { name: 'Pairs1', range: 'G2:H9', columns: [{ name: 'K' }, { name: 'V' }], style: { name: 'Pairs', rowStripes: true } },
    ],
  }],
}

export async function insideFile(): Promise<Uint8Array> {
  registerTableStyles()
  const { data } = await writeXlsx({ book: BOOK })
  return editParts(data, (name, text) => {
    if (name !== 'xl/styles.xml') return undefined
    let edited = text
    for (const [n, dxf] of Object.entries(FORMATS)) {
      const written = `<dxf><fill><patternFill><bgColor rgb="FF0101${n}"/></patternFill></fill></dxf>`
      if (!edited.includes(written)) throw new Error(`inside-held: marker ${n} is not in styles.xml`)
      edited = edited.replace(written, dxf)
    }
    return edited
  })
}
