/**
 * A one-sheet xlsx package written by hand, for the harness's books that the writer must not make (built-in
 * formats by number, theme colours, formulas of the old kind): the worksheet's content, and styles.xml, a theme
 * and what goes after the workbook's sheets when given.
 */
import { zip } from '../../../src/xlsx/zip/write.js'

/** The main namespace with the relationships prefix, for the parts' root elements. */
export const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const TYPE = 'application/vnd.openxmlformats-officedocument'
const RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'

export interface HandBook {
  sheet: string
  /** The worksheet's content: `<sheetData>…</sheetData>` and what goes around it. */
  body: string
  /** The content of `<styleSheet>`. */
  styles?: string
  theme?: string
  /** After `<sheets>` in the workbook: names, calculation settings. */
  workbook?: string
}

const enc = new TextEncoder()

export function handBook(b: HandBook): Promise<Uint8Array> {
  const parts: [name: string, type: string, rel: string, text: string][] = [['xl/worksheets/sheet1.xml', `${TYPE}.spreadsheetml.worksheet+xml`, 'worksheet', `<worksheet ${NS}>${b.body}</worksheet>`]]
  if (b.styles) parts.push(['xl/styles.xml', `${TYPE}.spreadsheetml.styles+xml`, 'styles', `<styleSheet ${NS}>${b.styles}</styleSheet>`])
  if (b.theme) parts.push(['xl/theme/theme1.xml', `${TYPE}.theme+xml`, 'theme', b.theme])
  const overrides = parts.map(([name, type]) => `<Override PartName="/${name}" ContentType="${type}"/>`).join('')
  const rels = parts.map(([name, , rel], i) => `<Relationship Id="rId${i + 1}" Type="${REL}/${rel}" Target="${name.slice(3)}"/>`).join('')
  const files: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="${TYPE}.spreadsheetml.sheet.main+xml"/>${overrides}</Types>`,
    '_rels/.rels': `<Relationships xmlns="${RELS}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    'xl/workbook.xml': `<workbook ${NS}><sheets><sheet name="${b.sheet}" sheetId="1" r:id="rId1"/></sheets>${b.workbook ?? ''}</workbook>`,
    'xl/_rels/workbook.xml.rels': `<Relationships xmlns="${RELS}">${rels}</Relationships>`,
    ...Object.fromEntries(parts.map(([name, , , text]) => [name, text])),
  }
  return zip(Object.entries(files).map(([name, text]) => ({ name, data: enc.encode(text) })))
}
