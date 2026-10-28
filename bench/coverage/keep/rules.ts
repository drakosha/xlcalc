/**
 * Conditional formats the model does not read, carried by `writeKept`, against Excel for Windows
 * (`bench/windows/run.ts keep rules-made|rules`). Each book of `RULES_BOOKS` is written with rules the model reads and
 * rules it does not (a colour scale with an automatic colour, Excel 2010's colour scale over another sheet, a data bar
 * whose thresholds lack their values with its 2010 part, and a rule whose fill is a gradient), and Excel opens and saves
 * it (`rules-made`). Over the saved `mixed` each case of `RULES_CASES` is done twice: by Excel through COM (a rule
 * added, deleted, given another priority or another look, or an operation of the structure) and by a host in a session
 * whose book `writeKept` writes over the same file; `rulesOf` gives what both files hold of the rules, in priority
 * order. `keep rules-open` (`keep.ts`) has Excel for Mac open each file `writeKept` wrote. Test:
 * `test/xlsx-keep-rules-excel.test.ts`.
 */
import type { ConditionalFormat, Session, StructureOp } from '../../../src/contract.js'
import { createSession } from '../../../src/index.js'
import { writeXlsx } from '../../../src/xlsx.js'
import { openKept, writeKept } from '../../../src/xlsx-keep.js'
import { openSource } from '../../../src/xlsx/keep/source.js'
import { fileRules } from '../../../src/xlsx/keep/rules.js'
import { plainArea, sqrefRects } from '../../../src/xlsx/keep/refs.js'
import { partText } from '../../../src/xlsx/read/package.js'
import { parseFormula } from '../../../src/xlsx/read/formulas.js'
import { palette } from '../../../src/xlsx/styles/color.js'
import { readDxfs } from '../../../src/xlsx/styles/dxf.js'
import { canonical } from '../../../src/syntax/print.js'
import { attributes, element, elements, placed } from '../../../src/xlsx/xml/scan.js'
import { readText } from '../../../src/xlsx/xml/text.js'
import { editParts } from '../xlsx/sent.js'

const X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main'
const XM = 'http://schemas.microsoft.com/office/excel/2006/main'
const BAR_EXT = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}'
const CF_EXT = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}'
const UNREAD_BAR = '{6D2B1D5A-0B8C-4C1E-9A43-2F4E0D7A1B01}'
const READ_BAR = '{3F1C7E22-5A4B-4D8E-B1C0-7E9D2A6F4C02}'
const SCALE_ID = '{9A7E3B10-2C4D-4F6A-8B1E-5D3C7A9F0E03}'

/**
 * The rules of the books in their order, by key: `r1` a cell value rule on A, `u1` a colour scale on B with an automatic
 * colour (unread), `r2` a formula rule on C relative to its cells, `p1` a formula rule on D whose fill is a gradient
 * (read without it), `u3` a data bar on E whose thresholds lack their values, with its 2010 part (both unread), `r3` a
 * data bar on F with its 2010 part, `u2` Excel 2010's colour scale on H with a threshold on the sheet P and an
 * automatic colour (unread). Rows 2 to 21; G1 and G2 hold the thresholds.
 */
const RULES: Record<string, { main?: (p: number) => string; ext?: (p: number) => string }> = {
  r1: { main: (p) => `<conditionalFormatting sqref="A2:A21"><cfRule type="cellIs" dxfId="0" priority="${p}" operator="greaterThan"><formula>100</formula></cfRule></conditionalFormatting>` },
  u1: { main: (p) => `<conditionalFormatting sqref="B2:B21"><cfRule type="colorScale" priority="${p}"><colorScale><cfvo type="min"/><cfvo type="num" val="$G$2"/><color auto="1"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>` },
  r2: { main: (p) => `<conditionalFormatting sqref="C2:C21"><cfRule type="expression" dxfId="1" priority="${p}"><formula>C2&gt;$G$1</formula></cfRule></conditionalFormatting>` },
  p1: { main: (p) => `<conditionalFormatting sqref="D2:D21"><cfRule type="expression" dxfId="2" priority="${p}"><formula>MOD(D2,2)=0</formula></cfRule></conditionalFormatting>` },
  u3: {
    main: (p) => `<conditionalFormatting sqref="E2:E21"><cfRule type="dataBar" priority="${p}"><dataBar><cfvo type="min"/><cfvo type="num"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="${BAR_EXT}" xmlns:x14="${X14}"><x14:id>${UNREAD_BAR}</x14:id></ext></extLst></cfRule></conditionalFormatting>`,
    ext: () => `<x14:conditionalFormatting xmlns:xm="${XM}"><x14:cfRule type="dataBar" id="${UNREAD_BAR}"><x14:dataBar minLength="0" maxLength="100" gradient="0"><x14:cfvo type="autoMin"/><x14:cfvo type="num"/><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>E2:E21</xm:sqref></x14:conditionalFormatting>`,
  },
  r3: {
    main: (p) => `<conditionalFormatting sqref="F2:F21"><cfRule type="dataBar" priority="${p}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF63C384"/></dataBar><extLst><ext uri="${BAR_EXT}" xmlns:x14="${X14}"><x14:id>${READ_BAR}</x14:id></ext></extLst></cfRule></conditionalFormatting>`,
    ext: () => `<x14:conditionalFormatting xmlns:xm="${XM}"><x14:cfRule type="dataBar" id="${READ_BAR}"><x14:dataBar minLength="0" maxLength="100" gradient="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:negativeFillColor rgb="FFFF0000"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>F2:F21</xm:sqref></x14:conditionalFormatting>`,
  },
  u2: { ext: (p) => `<x14:conditionalFormatting xmlns:xm="${XM}"><x14:cfRule type="colorScale" priority="${p}" id="${SCALE_ID}"><x14:colorScale><x14:cfvo type="min"/><x14:cfvo type="num"><xm:f>P!$A$1</xm:f></x14:cfvo><x14:color auto="1"/><x14:color rgb="FFF8696B"/></x14:colorScale></x14:cfRule><xm:sqref>H2:H21</xm:sqref></x14:conditionalFormatting>` },
}

/** The books: each unread kind alone beside the read rules, and all of them (`mixed`), the cases' book. */
export const RULES_BOOKS: Record<string, string[]> = {
  u1: ['r1', 'u1', 'r2'],
  u2: ['r1', 'r2', 'u2'],
  u3: ['r1', 'u3', 'r3'],
  p1: ['r1', 'p1', 'r2'],
  mixed: ['r1', 'u1', 'r2', 'p1', 'u3', 'r3', 'u2'],
}

/** The differential formats of the books: a red fill and dark red font (`r1`), bold (`r2`), a gradient fill (`p1`). */
const DXFS =
  '<dxfs count="3"><dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf><dxf><font><b/></font></dxf>' +
  '<dxf><fill><gradientFill degree="90"><stop position="0"><color rgb="FFFF0000"/></stop><stop position="1"><color rgb="FF0070C0"/></stop></gradientFill></fill></dxf></dxfs>'

/** A book's file as the step sends it to Excel: S with numbers in A2:H21 and the thresholds in G1:G2, P with a number in A1, and the book's rules put in by hand. */
export async function rulesFile(id: string): Promise<Uint8Array> {
  const cells: Record<string, { value: number | string }> = { G1: { value: 60 }, G2: { value: 90 } }
  for (const [k, c] of ['A', 'B', 'C', 'D', 'E', 'F', 'H'].entries()) {
    cells[`${c}1`] = { value: c }
    for (let r = 2; r <= 21; r++) cells[`${c}${r}`] = { value: ((r * 7 + k * 13) % 120) + 1 }
  }
  const { data } = await writeXlsx({ book: { sheets: [{ name: 'S', cells }, { name: 'P', cells: { A1: { value: 30 } } }] } })
  const keys = RULES_BOOKS[id]
  const main = keys.map((k, i) => RULES[k].main?.(i + 1) ?? '').join('')
  const ext = keys.map((k, i) => RULES[k].ext?.(i + 1) ?? '').join('')
  return editParts(data, (name, text) => {
    if (name === 'xl/styles.xml') return text.replace('</cellStyles>', `</cellStyles>${DXFS}`)
    if (name !== 'xl/worksheets/sheet1.xml') return undefined
    const list = ext ? `<extLst><ext uri="${CF_EXT}" xmlns:x14="${X14}"><x14:conditionalFormattings>${ext}</x14:conditionalFormattings></ext></extLst>` : ''
    return text.replace('</sheetData>', `</sheetData>${main}`).replace('</worksheet>', `${list}</worksheet>`)
  })
}

/**
 * What a case does to the rules of S, by priority among all the sheet's rules in the saved `mixed`: a cell value rule
 * over B2:B21 added (Excel's `FormatConditions.Add`, last, with the red style and stop if true as COM gives it), a rule
 * deleted, a rule given another priority, a rule's font made italic.
 */
export type RulesEdit = { add: true } | { remove: number } | { priority: [from: number, to: number] } | { italic: number }

export interface RulesCase {
  id: string
  op?: StructureOp
  edit?: RulesEdit
}

const rows = (op: 'insertRows' | 'deleteRows', row: number, count = 1, sheet = 'S'): StructureOp => ({ op, sheet, row, count })

export const RULES_CASES: RulesCase[] = [
  { id: 'add', edit: { add: true } },
  { id: 'delete-1', edit: { remove: 1 } },
  { id: 'delete-3', edit: { remove: 3 } },
  { id: 'delete-6', edit: { remove: 6 } },
  { id: 'first-6', edit: { priority: [6, 1] } },
  { id: 'last-1', edit: { priority: [1, 7] } },
  { id: 'italic-3', edit: { italic: 3 } },
  { id: 'insert-row-1', op: rows('insertRows', 1) },
  { id: 'insert-row-10', op: rows('insertRows', 10, 2) },
  { id: 'insert-row-22', op: rows('insertRows', 22, 2) },
  { id: 'delete-rows-5', op: rows('deleteRows', 5, 2) },
  { id: 'delete-row-2', op: rows('deleteRows', 2) },
  { id: 'delete-rows-all', op: rows('deleteRows', 2, 20) },
  { id: 'insert-col-b', op: { op: 'insertColumns', sheet: 'S', column: 'B' } },
  { id: 'delete-col-b', op: { op: 'deleteColumns', sheet: 'S', column: 'B' } },
  { id: 'delete-col-h', op: { op: 'deleteColumns', sheet: 'S', column: 'H' } },
  { id: 'insert-row-p', op: rows('insertRows', 1, 1, 'P') },
  { id: 'rename-p', op: { op: 'renameSheet', sheet: 'P', name: 'Q 2' } },
  { id: 'delete-sheet-p', op: { op: 'deleteSheet', sheet: 'P' } },
  { id: 'move-b', op: { op: 'moveRange', from: { sheet: 'S', range: 'B2:B21' }, to: { sheet: 'S', cell: 'J2' } } },
]

/** Excel's added rule as the host adds it: last, with the red style and stop if true. */
const ADDED: ConditionalFormat = { type: 'cellIs', ranges: ['B2:B21'], operator: 'greaterThan', formulas: ['150'], style: { fill: { color: '#FFC7CE' }, font: { color: '#9C0006', bold: true } }, stopIfTrue: true }

/** The rules of S of a file in priority order, the model's by their index (none for one it does not have). */
async function order(data: Uint8Array): Promise<(number | undefined)[]> {
  const src = await openSource(data)
  const [styles, theme] = await Promise.all(['styles', 'theme'].map((k) => partText(src.archive, src.wbRels, k)))
  const colors = palette(theme, styles)
  return fileRules(src.sheets[0].xml, 'S', readDxfs(styles, colors), colors).filter((s) => !s.joined).map((s) => s.index)
}

/** The case done by a host in a session over the file's book. */
export async function hostEdit(s: Session, data: Uint8Array, c: RulesCase): Promise<void> {
  if (c.op) {
    const r = s.editStructure(c.op)
    if (r.ok === false) throw new Error(`${c.id}: the session refused the operation: ${r.message}`)
  }
  const e = c.edit
  if (!e) return
  const list = s.getBook().sheets[0].conditionalFormats ?? []
  const all = await order(data)
  const model = (p: number) => {
    const i = all[p - 1]
    if (i === undefined) throw new Error(`${c.id}: the rule of priority ${p} is not the model's`)
    return i
  }
  let next = list
  if ('add' in e) next = [...list, ADDED]
  else if ('remove' in e) next = list.filter((_, i) => i !== model(e.remove))
  else if ('italic' in e) next = list.map((r, i) => (i === model(e.italic) && 'style' in r ? { ...r, style: { ...r.style, font: { ...r.style?.font, italic: true } } } : r))
  else {
    // Excel's order with the rule moved, of the model's rules
    const [from, to] = e.priority
    const moved = [...all]
    moved.splice(to - 1, 0, ...moved.splice(from - 1, 1))
    next = moved.flatMap((i) => (i === undefined ? [] : [list[i]]))
  }
  s.setLayout([{ sheet: 'S', conditionalFormats: next }])
}

/** The case written by `writeKept` over the file: the file, or the refusal. */
export async function rulesWritten(data: Uint8Array, c: RulesCase): Promise<{ data: Uint8Array } | { refused: string }> {
  const read = await openKept(data)
  const s = createSession(read.book, { cached: read.cached })
  try {
    await hostEdit(s, data, c)
    return { data: (await writeKept(data, s.getSnapshot())).data }
  } catch (e) {
    if ((e as Error).name !== 'KeptRefusal') throw e
    return { refused: (e as Error).message }
  } finally {
    s.close()
  }
}

/** The steps of a case in Excel for Windows (`xl.ps1`, `cond.ps1`), before the rules are listed and the book saved. */
export function rulesSteps(c: RulesCase, struct: (op: StructureOp) => Record<string, unknown>[]): Record<string, unknown>[] {
  const e = c.edit
  const steps = c.op ? struct(c.op) : []
  if (!e) return steps
  if ('add' in e) return [...steps, { do: 'cond', kind: 'make', sheet: 'S', range: 'B2:B21', add: 'cellIs', operator: 5, f1: '150' }]
  if ('remove' in e) return [...steps, { do: 'cond', kind: 'pick', pick: 'delete', sheet: 'S', priority: e.remove }]
  if ('italic' in e) return [...steps, { do: 'cond', kind: 'pick', pick: 'italic', sheet: 'S', priority: e.italic }]
  return [...steps, { do: 'cond', kind: 'pick', pick: 'priority', sheet: 'S', priority: e.priority[0], to: e.priority[1] }]
}

/** A formula as it compares: canonical where it parses. */
function canon(text: string): string {
  const f = parseFormula(text)
  return f.ast ? canonical(f.ast) : text
}

/** An `sqref` as it compares: its areas without `$`, sorted. */
const areas = (sqref: string) => (sqrefRects(sqref) ?? []).map(plainArea).sort().join(' ') || sqref

/**
 * What a file holds of the rules of its sheet S, in priority order: each rule as `[x14 ]type sqref` with its formulas
 * and thresholds' values canonical, `unread` for a rule the model does not have, `gradient` for a differential format
 * with a gradient fill, a data bar's 2010 part by its `sqref`; and the priorities as written.
 */
export async function rulesOf(data: Uint8Array): Promise<{ rules: string[]; priorities: number[] }> {
  const src = await openSource(data)
  const [styles, theme] = await Promise.all(['styles', 'theme'].map((k) => partText(src.archive, src.wbRels, k)))
  const colors = palette(theme, styles)
  const gradient = elements(element(styles ?? '', 'dxfs')?.body ?? '', 'dxf').map((d) => /gradientFill/.test(d.body ?? ''))
  const xml = src.sheets[0].xml
  const sources = fileRules(xml, 'S', readDxfs(styles, colors), colors)
  const rules: string[] = []
  const priorities: number[] = []
  for (const s of sources) {
    if (s.joined) continue
    const text = xml.slice(s.start, s.end)
    const a = attributes(text.slice(0, text.indexOf('>')))
    const f = [...placed(text, 'formula'), ...placed(text, 'f')].map((p) => canon(readText(p.body ?? '')))
    const vals = placed(text, 'cfvo').flatMap((p) => {
      const v = attributes(p.attrs).val
      return v === undefined ? [] : [canon(readText(v))]
    })
    const pair = s.id === undefined ? undefined : sources.find((o) => o !== s && o.id === s.id)
    const parts = [
      `${s.x14 ? 'x14 ' : ''}${a.type} ${areas(s.block.sqref)}`,
      ...[...f, ...vals].map((x) => `=${x}`),
      ...(s.index === undefined ? ['unread'] : []),
      ...(a.dxfId !== undefined && gradient[Number(a.dxfId)] ? ['gradient'] : []),
      ...(a.stopIfTrue === '1' ? ['stop'] : []),
      ...(pair && !s.x14 ? [`bar ${areas(pair.block.sqref)}`] : []),
    ]
    rules.push(parts.join(' '))
    priorities.push(s.priority)
  }
  return { rules, priorities }
}
