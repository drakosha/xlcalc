/**
 * Fill colours as Excel shows them: a book whose cells are filled with each theme colour at Excel's tints and with
 * each colour of the legacy palette, written here as XML. `npx tsx bench/coverage/run.ts xlsx colors` reads the
 * colours Excel shows into `test/excel/xlsx/colors.json`; `test/xlsx-excel.test.ts` checks the reader's against them.
 */
import { formatAddress } from '../../../src/model.js'
import { handBook } from './hand.js'

/** The Office theme's colours (Office 2013 to 2022). */
export const THEME =
  '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>'

/** The tints Excel's colour picker gives (lighter 80, 60, 40 %, darker 25, 50 %, for black and white 5 to 50 %) and rounder ones. */
const TINTS = [0.79998168889431442, 0.59999389629810485, 0.39997558519241921, -0.249977111117893, -0.499984740745262, -4.9989318521683403e-2, -0.14999847407452621, -0.34998626667073579, 0.499984740745262, 0.34998626667073579, 0.1, -0.1, 0.9]

export interface ColorCase {
  cell: string
  /** The attributes of the fill's `<fgColor>`. */
  color: string
}

export function colorCases(): ColorCase[] {
  const colors = [
    ...Array.from({ length: 12 }, (_, t) => [`theme="${t}"`, ...TINTS.map((x) => `theme="${t}" tint="${x}"`)]).flat(),
    ...Array.from({ length: 64 }, (_, i) => `indexed="${i}"`),
    'rgb="00FFC7CE"',
  ]
  return colors.map((color, i) => ({ cell: formatAddress(i % 50, Math.floor(i / 50)), color }))
}

/** The book: one cell per case, filled with its colour, the theme as a part of its own. */
export function colorBook(cases: ColorCase[]): Promise<Uint8Array> {
  const fills = cases.map((c) => `<fill><patternFill patternType="solid"><fgColor ${c.color}/><bgColor indexed="64"/></patternFill></fill>`)
  const xfs = cases.map((_, i) => `<xf numFmtId="0" fontId="0" fillId="${i + 2}" borderId="0" xfId="0" applyFill="1"/>`)
  const styles = `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="${fills.length + 2}"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>${fills.join('')}</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>${xfs.join('')}</cellXfs>`
  const byRow = new Map<number, string[]>()
  cases.forEach((c, i) => {
    const row = Number(/\d+/.exec(c.cell)![0])
    byRow.set(row, [...(byRow.get(row) ?? []), `<c r="${c.cell}" s="${i + 1}"/>`])
  })
  const rows = [...byRow].sort((a, b) => a[0] - b[0]).map(([r, cs]) => `<row r="${r}">${cs.join('')}</row>`)
  return handBook({ sheet: 'Colors', body: `<sheetFormatPr defaultRowHeight="15"/><sheetData>${rows.join('')}</sheetData>`, styles, theme: THEME })
}
