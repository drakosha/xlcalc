/**
 * Excel for Windows as the print oracle. The PDF supplies page members and positions. COM supplies the nominal sheet
 * measurements and the print configuration. A run writes its synthetic XLSX inputs, PDFs and summary into the chosen
 * directory. Run with XLWIN_HOST=your-excel-host and the current locale, for example:
 * XLWIN_HOST=your-excel-host XLWIN_CODEPAGE=any npx tsx bench/windows/print.ts ru-RU /tmp/xlcalc-print-matrix
 */
import { execFileSync } from 'node:child_process'
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import type { BookInput, CellInput } from '../../src/contract.js'
import { indexToCol } from '../../src/model.js'
import { layoutPrintPages, printSheetGeometry, writeXlsxWithPrint, type PrintBreak, type PrintHeaderFooter, type PrintSpan, type SheetPrint } from '../../src/xlsx-print.js'
import { metaOf, runJob, withWindowsExcel } from './remote.js'

interface Case {
  name: string
  height: number
  zoom: number | 'fit'
  columnWidth?: number
  top?: number
  bottom?: number
  paper?: number
  firstNumber?: number
  headerFooter?: PrintHeaderFooter
  titles?: PrintSpan
  repeatColumns?: PrintSpan
  repeats?: number
  areas?: string[]
  rowBreaks?: PrintBreak[]
  columnBreaks?: PrintBreak[]
  pageOrder?: 'downThenOver' | 'overThenDown'
  fitHeight?: number
  hiddenRows?: number[]
  markAllColumns?: boolean
  font?: BookInput['font']
}

const heights = [9, 12, 14.5, 15, 15.25, 15.5, 15.75, 16, 18, 24, 30]
const scales = [25, 50, 65, 67, 68, 69, 70, 72, 73, 74, 75, 90, 100, 125, 150, 200]
const cases: Case[] = [
  ...heights.map((height): Case => ({ name: `height-${height}`, height, zoom: 100 })),
  ...scales.filter((zoom) => zoom !== 100).map((zoom): Case => ({ name: `zoom-${zoom}`, height: 15, zoom })),
  ...[65, 67, 68, 69, 70, 72, 73, 75, 100, 'fit' as const].map((zoom): Case => ({ name: `wide-${zoom}`, height: 15, zoom, columnWidth: 50 })),
  ...[0.45, 0.48, 0.51, 0.52, 0.53, 0.55].map((bottom): Case => ({ name: `bottom-${bottom}`, height: 15, zoom: 100, bottom })),
  ...[0.492, 0.494, 0.495, 0.496, 0.497, 0.498, 0.499].map((bottom): Case => ({ name: `threshold-${bottom}`, height: 15, zoom: 100, bottom })),
  { name: 'threshold-shift', height: 15, zoom: 100, top: 0.54, bottom: 0.456 },
  { name: 'threshold-letter', height: 15, zoom: 100, paper: 1 },
  { name: 'threshold-a3', height: 15, zoom: 100, paper: 8 },
  {
    name: 'footer-parity', height: 15, zoom: 100, firstNumber: 4,
    headerFooter: {
      differentFirst: true, differentOddEven: true,
      firstFooter: '&CFIRST &P / &N / &P+2 / &N+1',
      evenFooter: '&CEVEN &P / &N / &P+2 / &N+1',
      oddFooter: '&CODD &P / &N / &P+2 / &N+1',
    },
  },
  {
    name: 'footer-codes', height: 15, zoom: 100,
    headerFooter: {
      oddFooter: '&LPLUS[&N+12] MINUS[&N-12] SIGN[&N+] NUM[&N3]&CPPLUS[&P+2] PMINUS[&P-2] PSIGN[&P+]&RPEMPTY[&P-] NCASE[&n]',
    },
  },
  {
    name: 'footer-terminal', height: 15, zoom: 100,
    headerFooter: { oddFooter: '&LBRACKET[&N+1] P[&P+1]&CNEND[&N+1&R PEND[&P+1' },
  },
  {
    name: 'footer-terminal-start4', height: 15, zoom: 100, firstNumber: 4,
    headerFooter: { oddFooter: '&LBRACKET[&N+1] P[&P+1]&CNEND[&N+1&R PEND[&P+1' },
  },
  { name: 'footer-implicit', height: 15, zoom: 100, headerFooter: { oddFooter: 'IMPLICIT' } },
  { name: 'ranges-disjoint', height: 15, zoom: 100, areas: ['A1:A48', 'A70:A125'] },
  { name: 'ranges-reversed', height: 15, zoom: 100, areas: ['A70:A125', 'A1:A48'] },
  { name: 'ranges-inner-titles', height: 15, zoom: 100, areas: ['A1:A125'], titles: { first: 20, last: 22 } },
  { name: 'ranges-below-titles', height: 15, zoom: 100, areas: ['A30:A100'], titles: { first: 5, last: 6 } },
  { name: 'ranges-above-titles', height: 15, zoom: 100, areas: ['A1:A25'], titles: { first: 30, last: 32 } },
  { name: 'ranges-fit-height', height: 15, zoom: 'fit', fitHeight: 1, areas: ['A1:A48', 'A70:A125'] },
  { name: 'ranges-fit-width', height: 15, zoom: 'fit', columnWidth: 50, areas: ['A1:B48', 'A70:D125'] },
  { name: 'breaks-rows', height: 15, zoom: 100, rowBreaks: [{ before: 12, manual: true }, { before: 46, manual: true }] },
  { name: 'breaks-partial-column', height: 15, zoom: 100, columnWidth: 50, rowBreaks: [{ before: 12, first: 2, last: 2, manual: true }] },
  { name: 'breaks-partial-area', height: 15, zoom: 100, columnWidth: 50, areas: ['A1:B125', 'C1:D125'], rowBreaks: [{ before: 12, first: 1, last: 2, manual: true }] },
  { name: 'breaks-columns-order', height: 15, zoom: 100, columnWidth: 50, pageOrder: 'overThenDown', columnBreaks: [{ before: 2, manual: true }] },
  { name: 'breaks-column-titles', height: 15, zoom: 100, columnWidth: 50, repeatColumns: { first: 1, last: 1 } },
  { name: 'breaks-hidden', height: 15, zoom: 100, hiddenRows: [2, 3, 34, 35, 36, 70, 71], rowBreaks: [{ before: 36, manual: true }] },
  { name: 'cross-titles', height: 52, zoom: 100, areas: ['A1:A39'], titles: { first: 20, last: 22 } },
  { name: 'cross-title-first-line', height: 52, zoom: 100, areas: ['A20:A39'], titles: { first: 20, last: 22 } },
  { name: 'cross-title-middle-line', height: 52, zoom: 100, areas: ['A21:A39'], titles: { first: 20, last: 22 } },
  { name: 'cross-partial-row-order', height: 15, zoom: 100, columnWidth: 50, pageOrder: 'overThenDown', markAllColumns: true,
    rowBreaks: [{ before: 12, first: 2, last: 2, manual: true }] },
  { name: 'cross-partial-column', height: 15, zoom: 100, columnWidth: 50, markAllColumns: true,
    columnBreaks: [{ before: 2, first: 30, last: 50, manual: true }] },
  { name: 'cross-both', height: 15, zoom: 100, columnWidth: 50, markAllColumns: true,
    rowBreaks: [{ before: 12, first: 2, last: 2, manual: true }], columnBreaks: [{ before: 2, first: 30, last: 50, manual: true }] },
  { name: 'cross-both-over', height: 15, zoom: 100, columnWidth: 50, pageOrder: 'overThenDown', markAllColumns: true,
    rowBreaks: [{ before: 12, first: 2, last: 2, manual: true }], columnBreaks: [{ before: 2, first: 30, last: 50, manual: true }] },
  { name: 'footer-terminal-long', height: 15, zoom: 100, headerFooter: { oddFooter: '&RTERMINAL[&P+12' } },
  { name: 'footer-terminal-minus', height: 15, zoom: 100, headerFooter: { oddFooter: '&RTERMINAL[&P-12' } },
  { name: 'footer-negative-page', height: 15, zoom: 100, firstNumber: -2, headerFooter: { oddFooter: '&RPage &P of &N' } },
  { name: 'no-repeat', height: 15, zoom: 100, repeats: 0 },
  { name: 'height-18-zoom-90', height: 18, zoom: 90 },
  { name: 'arial', height: 15, zoom: 100, font: { name: 'Arial', size: 10 } },
  { name: 'times', height: 15, zoom: 100, font: { name: 'Times New Roman', size: 11 } },
]

const LAST_ROW = 125

/** Scale 100 labels both axes, in cells wide enough to print with their neighbors filled. */
function make(c: Case) {
  const columns = c.columnWidth ? 4 : 1
  const lastColumn = indexToCol(columns - 1)
  const range = `A1:${lastColumn}${LAST_ROW}`
  const cells: Record<string, CellInput> = {}
  for (let row = 1; row <= LAST_ROW; row++) cells[`A${row}`] = { value: 'R' + String(row).padStart(3, '0') }
  for (let column = 1; column < columns; column++) {
    const name = indexToCol(column)
    cells[`${name}1`] = { value: 'C' + String(column + 1).padStart(2, '0') }
    if (c.markAllColumns) for (let row = 2; row <= LAST_ROW; row++) {
      cells[`${name}${row}`] = { value: 'R' + String(row).padStart(3, '0') + '-' + name }
    }
  }
  const rows = c.hiddenRows
    ? Object.fromEntries(Array.from({ length: LAST_ROW }, (_, i) => [String(i + 1), { height: c.height, ...(c.hiddenRows?.includes(i + 1) ? { hidden: true as const } : {}) }]))
    : { [`1:${LAST_ROW}`]: { height: c.height } }
  const sheet = {
    name: 'Report', cells, rows,
    ...(c.columnWidth ? { columns: { [`A:${lastColumn}`]: { width: c.columnWidth } } } : {}),
  }
  const book: BookInput = { sheets: [sheet], ...(c.font ? { font: c.font } : {}) }
  const repeats = c.repeats ?? 2
  const settings: SheetPrint = {
    sheet: sheet.name, areas: c.areas ?? [range],
    ...(c.titles ? { repeatRows: c.titles } : repeats ? { repeatRows: { first: 1, last: repeats } } : {}),
    repeatColumns: c.repeatColumns,
    rowBreaks: c.rowBreaks,
    columnBreaks: c.columnBreaks,
    setup: {
      orientation: 'landscape', paperSize: c.paper ?? 9, pageOrder: c.pageOrder,
      ...(c.firstNumber === undefined ? {} : { useFirstPageNumber: true, firstPageNumber: c.firstNumber }),
      ...(c.zoom === 'fit' ? { fitToPage: true, fitToWidth: 1, fitToHeight: c.fitHeight ?? 0 } : { scale: c.zoom }),
    },
    margins: { left: 0.5, right: 0.5, top: c.top ?? 0.5, bottom: c.bottom ?? 0.5, header: 0.2, footer: 0.2 },
    headerFooter: c.headerFooter ?? { oddFooter: '&RPage &P of &N' },
  }
  return { book, settings, range }
}

interface PdfAnswer {
  printer: string
  quality: number[] | null
  margins: number[]
  zoom: number | boolean
  rangeWidth: number
  rangeHeight: number
  rowBreaks: number[]
  columnBreaks: number[]
  pdf: string
}

function printFacts(file: string) {
  const xml = execFileSync('pdftotext', ['-bbox', file, '-'], { encoding: 'utf8', maxBuffer: 64 << 20 })
  return [...xml.matchAll(/<page width="([^"]+)" height="([^"]+)">([\s\S]*?)<\/page>/g)].map(([, width, height, page]) => {
    const words = [...page.matchAll(/<word xMin="([^"]+)" yMin="([^"]+)" xMax="[^"]+" yMax="[^"]+">([^<]+)<\/word>/g)]
    const rows = words.filter((word) => /^R\d{3}$/.test(word[3])).map(([, x, y, label]) => ({ row: Number(label.slice(1)), x: Number(x), y: Number(y) }))
    const sideRows = words.flatMap(([, x, y, word]) => {
      const marker = /^R(\d{3})-([B-D])$/.exec(word)
      return marker ? [{ row: Number(marker[1]), column: marker[2], x: Number(x), y: Number(y) }] : []
    })
    const columns = words.filter((word) => /^C\d{2}$/.test(word[3])).map(([, x, , label]) => ({ column: Number(label.slice(1)), x: Number(x) }))
    const band = words.filter((word) => Number(word[2]) > Number(height) - 40).map(([, x, y, text]) => ({ x: Number(x), y: Number(y), text }))
    const pitch = rows.length >= 2 ? (rows.at(-1)!.y - rows[0].y) / (rows.length - 1) : undefined
    return {
      width: Number(width), height: Number(height),
      ...(rows.length ? { first: rows.slice(0, 3), last: rows.at(-1), count: rows.length } : {}),
      ...(pitch === undefined ? {} : { pitch: Number(pitch.toFixed(6)) }),
      columns,
      ...(sideRows.length ? { side: Object.fromEntries([...new Set(sideRows.map((row) => row.column))].map((column) => {
        const labels = sideRows.filter((row) => row.column === column)
        return [column, { first: labels.slice(0, 3), last: labels.at(-1), count: labels.length }]
      })) } : {}),
      band,
    }
  })
}

const [locale, directory, ...names] = process.argv.slice(2)
if (!locale || !directory) throw new Error('usage: print.ts <Excel locale> <output directory> [case names or prefixes]')
const selected = names.length ? cases.filter((c) => names.some((name) => c.name === name || c.name.startsWith(name))) : cases
if (!selected.length) throw new Error('no matching print cases')
mkdirSync(directory, { recursive: true })
const entries = await Promise.all(selected.map(async (c) => {
  const data = make(c)
  const file = join(directory, c.name + '.xlsx')
  const written = await writeXlsxWithPrint({ book: data.book }, [data.settings])
  if (written.issues.length) throw new Error(`${c.name}: ${JSON.stringify(written.issues)}`)
  writeFileSync(file, written.data)
  return { c, ...data, job: { file, steps: [{ do: 'printPdf', sheet: data.settings.sheet, range: data.range }] } }
}))

await withWindowsExcel(locale, async (env) => {
  const answers = await runJob('print-pages', entries.map((entry) => entry.job), 120_000 + 15_000 * entries.length)
  const results = entries.map(({ c, book, settings, range }, i) => {
    const answer = answers[i]
    if (!answer || !('answers' in answer)) throw new Error(`${c.name}: ${answer?.rejected ?? 'no Excel answer'}`)
    const { pdf, ...setup } = answer.answers[0] as PdfAnswer
    const file = join(directory, c.name + '.pdf')
    writeFileSync(file, Buffer.from(pdf, 'base64'))
    const pages = printFacts(file)
    if (!pages.length) throw new Error(`${c.name}: no PDF pages`)
    const geometry = printSheetGeometry(book.sheets[0], { font: book.font, usedRange: range })
    const plan = layoutPrintPages(settings, geometry, pages[0])
    const nominal = plan.pages.map((page) => ({ scale: page.scale, ranges: page.regions.map((region) => region.range) }))
    console.log(c.name, 'Excel:', pages.map((page) => page.last?.row ?? '-').join(','), 'xlcalc:', nominal.map((page) => page.ranges.at(-1)).join(','))
    return { case: c, setup, pages, nominal }
  })
  writeFileSync(join(directory, 'print-pages.json'), JSON.stringify({ meta: metaOf(env, locale), results }, null, 2) + '\n')
})
