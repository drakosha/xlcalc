/**
 * Hand-made xlsx packages for the reader's tests: parts as text, zipped by the package's own writer. `getBook()` makes
 * a minimal workbook around sheet bodies; any part can be replaced or left out to test what a file may lack.
 */
import type { CellValue, ComputeResult } from '../src/contract.js'
import { readXlsx, type XlsxCached, type XlsxRead } from '../src/xlsx.js'
import { zip } from '../src/xlsx/zip/write.js'

const enc = new TextEncoder()

export const pack = (parts: Record<string, string | undefined>): Promise<Uint8Array> =>
  zip(Object.entries(parts).flatMap(([name, text]) => (text === undefined ? [] : [{ name, data: enc.encode(text) }])))

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

export interface FixtureSheet {
  name: string
  /** The worksheet's content: `<sheetData>…</sheetData>` and what goes around it. */
  body: string
  /** Attributes of `<sheet>` beyond name, id and relationship (`state="hidden"`). */
  attrs?: string
}

export interface Fixture {
  sheets: FixtureSheet[]
  strings?: string[]
  styles?: string
  workbook?: { before?: string; after?: string }
  /** Parts added or replaced (undefined removes one). */
  parts?: Record<string, string | undefined>
}

/** The parts of a minimal workbook: sheets in order, shared strings and styles when given. */
export function parts(f: Fixture): Record<string, string | undefined> {
  const out: Record<string, string | undefined> = {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  }
  const rels: string[] = []
  f.sheets.forEach((s, i) => {
    out[`xl/worksheets/sheet${i + 1}.xml`] = `<worksheet ${NS}>${s.body}</worksheet>`
    rels.push(`<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
  })
  if (f.strings) {
    out['xl/sharedStrings.xml'] = `<sst ${NS}>${f.strings.map((s) => (s.startsWith('<') ? `<si>${s}</si>` : `<si><t>${s}</t></si>`)).join('')}</sst>`
    rels.push(`<Relationship Id="rIdS" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`)
  }
  if (f.styles) {
    out['xl/styles.xml'] = `<styleSheet ${NS}>${f.styles}</styleSheet>`
    rels.push(`<Relationship Id="rIdY" Type="${REL}/styles" Target="styles.xml"/>`)
  }
  const sheets = f.sheets.map((s, i) => `<sheet name="${s.name}" sheetId="${i + 1}" r:id="rId${i + 1}"${s.attrs ? ` ${s.attrs}` : ''}/>`).join('')
  out['xl/workbook.xml'] = `<workbook ${NS}>${f.workbook?.before ?? ''}<sheets>${sheets}</sheets>${f.workbook?.after ?? ''}</workbook>`
  out['xl/_rels/workbook.xml.rels'] = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels.join('')}</Relationships>`
  return { ...out, ...f.parts }
}

export const read = async (f: Fixture) => readXlsx(await pack(parts(f)))

/** A one-sheet book (`S`) from its sheetData rows. */
export const sheet = (rows: string, extra: Omit<Fixture, 'sheets'> = {}) => read({ sheets: [{ name: 'S', body: `<sheetData>${rows}</sheetData>` }], ...extra })

/** The issue codes of a read, with the cells each is about. */
export const codes = (issues: { code: string; cells?: string[] }[]) => issues.map((i) => (i.cells ? `${i.code} ${i.cells.join(' ')}` : i.code))

/** The values a file stored on one of its sheets (`XlsxRead.cached`), by address. */
export const storedValues = (r: XlsxRead, sheet = 0): Record<string, CellValue> => Object.fromEntries(Object.entries(r.cached.sheets[sheet].cells).map(([a, c]) => [a, c.value]))

/** What a file written from `result` stores of it: the formulas' values and the cells of their arrays, with `spill` and `spillOf`. */
export const storedOf = (result: ComputeResult): XlsxCached => ({
  sheets: result.sheets.map((s) => ({
    name: s.name,
    cells: Object.fromEntries(
      Object.entries(s.cells).flatMap(([a, c]) => (c.formula !== undefined || c.spillOf !== undefined ? [[a, { value: c.value, ...(c.spill !== undefined ? { spill: c.spill } : {}), ...(c.spillOf !== undefined ? { spillOf: c.spillOf } : {}) }]] : [])),
    ),
  })),
})
