/**
 * Built-in number formats as Excel shows them: a book whose cell A<n> has the built-in format n (1 to 58), written
 * here as XML. `npx tsx bench/coverage/run.ts xlsx formats` reads the codes Excel en-US gives them into
 * `test/excel/xlsx/formats.json`; `test/xlsx-excel.test.ts` checks the reader's table against them.
 */
import { handBook } from './hand.js'

export const FORMAT_IDS = Array.from({ length: 58 }, (_, i) => i + 1)

export function formatBook(): Promise<Uint8Array> {
  const xfs = FORMAT_IDS.map((id) => `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`)
  const styles = `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>${xfs.join('')}</cellXfs>`
  const rows = FORMAT_IDS.map((id, i) => `<row r="${id}"><c r="A${id}" s="${i + 1}"><v>1</v></c></row>`)
  return handBook({ sheet: 'Formats', body: `<sheetFormatPr defaultRowHeight="15"/><sheetData>${rows.join('')}</sheetData>`, styles })
}
