/**
 * All Using Source Theme beside All and Formats between two workbooks of different themes, in both Excels: the source
 * open beside the target (`fromBook`), its theme another than the target's Office theme (accent 1 and 2 and the minor
 * font changed). The source's cells: a fill of accent 1, a font of accent 2, a bold font of the minor scheme, a fill of
 * an explicit color, a fill of accent 1 lightened 40 %, a font named without a scheme. The engine has no paste between
 * workbooks: what a host gets that carries the cells over by reading the source (`readXlsx` resolves a theme's colors
 * and keeps a font's name) is the source's colors and font names, which this set compares with what each paste gives.
 * Both books are written by hand (`handBook`), as the writer writes no theme.
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import { readXlsx } from '../../../src/xlsx.js'
import { internStyle } from '../../../src/style-table.js'
import { THEME } from '../xlsx/colors.js'
import { handBook } from '../xlsx/hand.js'
import type { PasteCase, PasteCases, PasteWhat } from './cases.js'
import { shownBook, type Shown } from './shown.js'

/** The source's theme: Office with accent 1 purple, accent 2 teal and Georgia as the minor font. */
export const OTHER_THEME = THEME.replaceAll('name="Office"', 'name="Other"').replace('4472C4', '8E44AD').replace('ED7D31', '16A085').replace('<a:minorFont><a:latin typeface="Calibri"/>', '<a:minorFont><a:latin typeface="Georgia"/>')

/** A sheet `S` of values and the styles of `xfs` (font, fill) over fonts and fills given as XML, the default font of the theme's minor scheme named `minor`. */
function themed(minor: string, cells: [string, string | number, number][], fonts: string[], fills: string[], xfs: [number, number][], theme: string): Promise<Uint8Array> {
  const font0 = `<font><sz val="11"/><color theme="1"/><name val="${minor}"/><family val="2"/><scheme val="minor"/></font>`
  const fillsXml = `<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${fills.map((f) => `<fill><patternFill patternType="solid">${f}<bgColor indexed="64"/></patternFill></fill>`).join('')}`
  const xfsXml = `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>${xfs.map(([font, fill]) => `<xf numFmtId="0" fontId="${font}" fillId="${fill}" borderId="0" xfId="0"${font ? ' applyFont="1"' : ''}${fill ? ' applyFill="1"' : ''}/>`).join('')}`
  const styles = `<fonts count="${fonts.length + 1}">${font0}${fonts.join('')}</fonts><fills count="${fills.length + 2}">${fillsXml}</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length + 1}">${xfsXml}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>`
  const value = (v: string | number) => (typeof v === 'number' ? `<v>${v}</v>` : `<is><t>${v}</t></is>`)
  const rows = cells.map(([a, v, s]) => `<row r="${a.slice(1)}"><c r="${a}"${s ? ` s="${s}"` : ''}${typeof v === 'string' ? ' t="inlineStr"' : ''}>${value(v)}</c></row>`)
  return handBook({ sheet: 'S', body: `<sheetData>${rows.join('')}</sheetData>`, styles, theme })
}

/** The source: A1:A6 as the set's comment lists them, in the other theme. */
const sourceFile = () =>
  themed(
    'Georgia',
    [['A1', 1, 1], ['A2', 2, 2], ['A3', 'minor', 3], ['A4', 4, 4], ['A5', 5, 5], ['A6', 'named', 6]],
    ['<font><sz val="11"/><color theme="5"/><name val="Georgia"/><family val="2"/><scheme val="minor"/></font>', '<font><b/><sz val="11"/><color theme="1"/><name val="Georgia"/><family val="2"/><scheme val="minor"/></font>', '<font><sz val="11"/><color theme="1"/><name val="Georgia"/><family val="1"/></font>'],
    ['<fgColor theme="4"/>', '<fgColor rgb="FF00B050"/>', '<fgColor theme="4" tint="0.39997558519241921"/>'],
    [[0, 2], [1, 0], [2, 0], [0, 3], [0, 4], [3, 0]],
    OTHER_THEME,
  )

/** The target: values in C1:C6, the Office theme. */
const targetFile = () => themed('Calibri', ['C1', 'C2', 'C3', 'C4', 'C5', 'C6'].map((a) => [a, 'old', 0] as [string, string, number]), [], [], [], THEME)

/** The books as `BookInput` for the stored file; the files sent are written by hand (`made`). */
const book = (cells: string[], value: string | number): BookInput => ({ sheets: [{ name: 'S', cells: Object.fromEntries(cells.map((a) => [a, { value }])) }] })

export function buildPasteCases(): PasteCases {
  const cases: PasteCase[] = (['allUsingSourceTheme', 'all', 'formats'] as PasteWhat[]).map((what) => ({ id: what, book: 'target', fromBook: 'source', from: { sheet: 'S', range: 'A1:A6' }, to: { sheet: 'S', range: 'C1' }, what }))
  return {
    books: { target: book(['C1', 'C2', 'C3', 'C4', 'C5', 'C6'], 'old'), source: book(['A1', 'A2', 'A3', 'A4', 'A5', 'A6'], 1) },
    cases,
    made: { target: targetFile, source: sourceFile },
  }
}

/**
 * What a host gets that carries a case's cells over from the source to the target by reading both files: the target's
 * book with each source cell's value (but for Formats) and style at its place in C1:C6, the style as the reader gives it
 * (a theme's colors resolved by the source's theme, a font's name kept, the default font left out).
 */
export async function carried(c: PasteCase): Promise<Shown> {
  const [source, target] = await Promise.all([sourceFile(), targetFile()].map(async (f) => (await readXlsx(await f)).book))
  let styles = target.styles ?? []
  const cells: Record<string, CellInput> = { ...target.sheets[0].cells }
  for (let r = 1; r <= 6; r++) {
    const x = source.sheets[0].cells[`A${r}`]
    const at = `C${r}`
    const cell: CellInput = c.what === 'formats' ? { value: cells[at].value } : { value: x.value }
    if (x.style !== undefined) {
      const put = internStyle(styles, source.styles![x.style])
      styles = put.styles
      cell.style = put.index
    }
    cells[at] = cell
  }
  return shownBook({ ...target, styles, sheets: [{ ...target.sheets[0], cells }] })
}
