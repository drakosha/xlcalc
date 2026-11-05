/**
 * What the reader reads of a file another writer made, against what Excel shows: a colour with more than one kind (`auto`, `indexed`,
 * `rgb`, `theme`; [MS-OI29500] 2.1.694, 2.1.704 name the order auto < indexed < rgb < theme), and the first two fills, which 2.1.705,
 * 2.1.1084 say Excel reserves and does not read from the file. A book per `FILE`, a cell per case. `npx tsx bench/coverage/run.ts xlsx
 * fills-read` and `npx tsx bench/windows/run.ts xlsx fills-read` read each cell's fill and font into `xlsx/fills-read.json` of each
 * Excel's answers. `fills-read.predicted.json` was committed before Excel ran the set: the notes' rules, the pattern for each cell, the
 * colour where the pattern is solid, the font's colour where a case gives one.
 */
import { handBook } from './hand.js'
import { THEME } from './colors.js'

export interface FillCell {
  id: string
  cell: string
  /** The attributes of the fill's `<fgColor>` and of the font's `<color>`. */
  fg?: string
  font?: string
  /** A cell of a `reserved` file: the `fillId` of its format. */
  fillId?: number
}

export interface FillFile {
  id: string
  /** The first two `<fill>` of the file; without them `none` and `gray125`, as Excel writes them. */
  first?: [string, string]
  cells: FillCell[]
}

const solid = (rgb: string) => `<fill><patternFill patternType="solid"><fgColor rgb="${rgb}"/><bgColor indexed="64"/></patternFill></fill>`
const NONE = '<fill><patternFill patternType="none"/></fill>'
const GRAY = '<fill><patternFill patternType="gray125"/></fill>'

const KINDS: [string, string][] = [
  ['theme-rgb', 'theme="4" rgb="FFFF0000"'],
  ['auto-rgb', 'auto="1" rgb="FFFF0000"'],
  ['indexed-rgb', 'indexed="2" rgb="FF00FF00"'],
  ['indexed-theme', 'indexed="2" theme="4"'],
  ['auto-theme', 'auto="1" theme="4"'],
  ['auto-indexed', 'auto="1" indexed="2"'],
  ['all-three', 'indexed="2" rgb="FF00FF00" theme="4"'],
  ['theme-tint-rgb', 'theme="4" tint="0.39997558519241921" rgb="FFFF0000"'],
]

const reserved = (id: string, first: [string, string]): FillFile => ({ id, first, cells: [{ id: `${id}-0`, cell: 'A1', fillId: 0 }, { id: `${id}-1`, cell: 'B1', fillId: 1 }, { id: `${id}-2`, cell: 'C1', fg: 'rgb="FF0000FF"' }] })

export const FILL_FILES: FillFile[] = [
  {
    id: 'kinds',
    cells: [
      ...KINDS.map(([id, fg], i): FillCell => ({ id: `fill-${id}`, cell: `A${i + 1}`, fg })),
      ...KINDS.slice(0, 7).map(([id, font], i): FillCell => ({ id: `font-${id}`, cell: `B${i + 1}`, font })),
    ],
  },
  reserved('control', [NONE, GRAY]),
  reserved('red-green', [solid('FFFF0000'), solid('FF00FF00')]),
  reserved('swapped', [GRAY, NONE]),
]

/** A file's styles: fonts and fills of its cases, one format per cell, the format 0 plain. */
export function fillBook(f: FillFile): Promise<Uint8Array> {
  const first = f.first ?? [NONE, GRAY]
  const fills = [...first]
  const fonts = ['<font><sz val="11"/><name val="Calibri"/></font>']
  const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
  const cells = f.cells.map((c, i) => {
    let fillId = c.fillId ?? 0
    if (c.fg) fills.push(`<fill><patternFill patternType="solid"><fgColor ${c.fg}/><bgColor indexed="64"/></patternFill></fill>`), (fillId = fills.length - 1)
    let fontId = 0
    if (c.font) fonts.push(`<font><sz val="11"/><color ${c.font}/><name val="Calibri"/></font>`), (fontId = fonts.length - 1)
    xfs.push(`<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="0" xfId="0" applyFont="1" applyFill="1"/>`)
    return `<c r="${c.cell}" s="${i + 1}"><v>1</v></c>`
  })
  const rows = new Map<number, string[]>()
  f.cells.forEach((c, i) => {
    const r = Number(/\d+/.exec(c.cell)![0])
    rows.set(r, [...(rows.get(r) ?? []), cells[i]])
  })
  const styles = `<fonts count="${fonts.length}">${fonts.join('')}</fonts><fills count="${fills.length}">${fills.join('')}</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${xfs.length}">${xfs.join('')}</cellXfs>`
  const sheet = [...rows].sort((a, b) => a[0] - b[0]).map(([r, cs]) => `<row r="${r}">${cs.join('')}</row>`).join('')
  return handBook({ sheet: 'S', body: `<sheetData>${sheet}</sheetData>`, styles, theme: THEME })
}
