import type { PrintBreak, PrintHeaderFooter, WorksheetPrintSettings } from '../../contract.js'
import { MAX_COLS, MAX_ROWS, indexToCol } from '../../model.js'
import { parseRange } from '../../structure/area.js'
import { areaText, sheetPrefix } from '../keep/refs.js'
import type { WorksheetPrintParts } from '../write/book.js'
import { escapeText } from '../xml/text.js'
import { BREAK, HEADER_ATTRS, HEADER_TEXT, MARGINS, OPTIONS, PAGE_FIELD_WRAP, PROPERTIES, SETUP, leaf, writeFields } from './xml.js'

function namesFor(given: WorksheetPrintSettings, sheet: string, index: number): string[] {
  const prefix = sheetPrefix(sheet) + '!'
  const names: string[] = []
  const name = (kind: string, refs: string[]) => {
    if (refs.length) names.push(`<definedName name="_xlnm.${kind}" localSheetId="${index}">${escapeText(refs.map((ref) => prefix + ref).join(','))}</definedName>`)
  }
  if (given.areas) name('Print_Area', given.areas.map((range) => areaText(parseRange(range.replace(/\$/g, '').toUpperCase())!)))
  const titles: string[] = []
  const rows = given.repeatRows
  const columns = given.repeatColumns
  if (rows) titles.push(`$${rows.first}:$${rows.last}`)
  if (columns) titles.push(`$${indexToCol(columns.first - 1)}:$${indexToCol(columns.last - 1)}`)
  name('Print_Titles', titles)
  return names
}

function headerFooter(given: PrintHeaderFooter | undefined): string {
  if (!given) return ''
  const attrs = writeFields(HEADER_ATTRS, { ...given })
  const content = HEADER_TEXT.flatMap((tag) => given[tag] === undefined ? [] : [`<${tag}>${escapeText(given[tag]!)}</${tag}>`]).join('')
  return content ? `<headerFooter${attrs}>${content}</headerFooter>` : leaf('headerFooter', attrs)
}

function breaks(items: readonly PrintBreak[] | undefined, tag: string, across: number): string {
  if (!items?.length) return ''
  let manualCount = 0
  const content = items.map(({ before, first, last, manual, pivot }) => {
    if (manual === true) manualCount++
    const attrs = writeFields(BREAK, { id: before - 1, min: first === undefined ? undefined : first - 1, max: (last ?? across) - 1, man: manual, pt: pivot })
    return leaf('brk', attrs)
  }).join('')
  return `<${tag} count="${items.length}" manualBreakCount="${manualCount}">${content}</${tag}>`
}

export function worksheetPrintParts(print: WorksheetPrintSettings, sheet: string, index: number): WorksheetPrintParts {
  const first = print.setup?.firstPageNumber
  const firstPageNumber = first !== undefined && first < 0 ? PAGE_FIELD_WRAP + first : first
  return {
    properties: leaf('pageSetUpPr', writeFields(PROPERTIES, { ...print.setup })),
    settings: leaf('printOptions', writeFields(OPTIONS, { ...print.options })) +
      leaf('pageMargins', writeFields(MARGINS, { ...print.margins })) +
      leaf('pageSetup', writeFields(SETUP, { ...print.setup, firstPageNumber })) +
      headerFooter(print.headerFooter) +
      breaks(print.rowBreaks, 'rowBreaks', MAX_COLS) +
      breaks(print.columnBreaks, 'colBreaks', MAX_ROWS),
    names: namesFor(print, sheet, index),
  }
}
