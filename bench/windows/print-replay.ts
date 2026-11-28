import { readXlsx } from '../../src/xlsx.js'
/** Replays the print matrix on the measured Windows Excel 16.0.20430 Microsoft Print to PDF profile. */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { advance } from '../../src/numfmt/glyphs.js'
import { formatPrintHeaderFooter, layoutPrintPages, layoutPrintWorkbook, printPaperSize, printSheetGeometry, type PrintPage, type PrintNotes } from '../../src/xlsx-print.js'

interface Word {
  x: number
  y: number
  text: string
}

interface Trace {
  case: { name: string; height: number; columnWidth?: number; paper?: number; hiddenRows?: number[]; workbook?: object; noteMode?: string; font?: { name: string } }
  setup: { printer: string }
  pages: { width: number; height: number; pitch?: number; last?: { row: number }; band?: Word[]; words?: Word[] }[]
}

function traces(directory: string): Trace[] {
  return (JSON.parse(readFileSync(join(directory, 'print-pages.json'), 'utf8')) as { results: Trace[] }).results
}

const [controlDirectory, ...also] = process.argv.slice(2)
if (!controlDirectory) throw new Error('usage: print-replay.ts <control matrix directory> [additional trace directories]')
const controls = traces(controlDirectory)
const controlByName = new Map(controls.map((run) => [run.case.name, run]))
const pointsPerUnit = 0.120023
const trailingClearance = 0.72
const wideColumnUnits = 2350
const noteLabels = { address: 'Ячейка:', body: 'Примечание.' }
// Excel quantizes both the font size and the line pitch on this device, independently of the grid.
const noteFont = new Map([
  [0.42, { pitch: 54, size: 4.56, inkTop: 0.9 }],
  [0.5, { pitch: 64, size: 5.52, inkTop: 1.14 }],
  [1, { pitch: 121, size: 11.04, inkTop: 2.28 }],
])

function checkNotes(pages: Trace['pages'], plan: PrintPage[]): string | undefined {
  if (pages.length !== plan.length) return 'page count: Excel ' + pages.length + ', xlcalc ' + plan.length
  const noteStart = pages.findIndex((page) => page.words?.some((word) => word.text === noteLabels.address))
  if (noteStart !== plan.findIndex((page) => page.noteLines !== undefined)) return 'first end-notes page differs'

  for (const [i, page] of plan.entries()) {
    const observed = new Map<number, Word[]>()
    for (const word of pages[i].words ?? []) {
      const row = observed.get(word.y)
      if (row) row.push(word)
      else observed.set(word.y, [word])
    }
    const pageText = [...observed.values()].map((row) => row.map((word) => word.text).join(' '))
    if (!pageText.includes('Page ' + page.number + ' of ' + plan.length)) return 'footer on page ' + (i + 1)
    if (!page.noteLines) continue
    const metric = noteFont.get(page.scale)
    if (!metric) return 'unmeasured note scale ' + page.scale
    const actual = [...observed.values()].sort((a, b) => a[0].y - b[0].y).slice(0, -1)
    const expected = page.noteLines.filter((line) => line.label || line.text)
    if (actual.length !== expected.length) return 'line count on page ' + (i + 1) + ': Excel ' + actual.length + ', xlcalc ' + expected.length
    for (const [j, line] of expected.entries()) {
      const words = actual[j].sort((a, b) => a.x - b.x)
      const text = [line.label?.text, line.text].filter(Boolean).join(' ').replace(/\s+/g, ' ').trim()
      if (text !== words.map((word) => word.text).join(' ')) {
        return 'text on page ' + (i + 1) + ', line ' + (j + 1) + ': Excel ' + words.map((word) => word.text).join(' ') + ', xlcalc ' + text
      }
      const column = words[line.label ? 1 : 0]
      if (line.text && Math.abs(column.x - line.bounds.x) > 0.4) return 'text x on page ' + (i + 1) + ', line ' + (j + 1)
      if (line.label && Math.abs(words[0].x - line.label.bounds.x) > 0.4) return 'label x on page ' + (i + 1) + ', line ' + (j + 1)
      if (Math.abs(words[0].y - line.bounds.y - metric.inkTop) > 0.05) return 'text y on page ' + (i + 1) + ', line ' + (j + 1)
    }
  }
}

let total = 0
let unknown = 0
let different = 0
for (const directory of [controlDirectory, ...also]) {
  for (const run of traces(directory)) {
    const { case: c, setup, pages } = run
    const reference = c.font?.name === 'Times New Roman' ? 'times' : `height-${c.height}`
    const measuredPitch = controlByName.get(reference)?.pages[0].pitch
    if (c.noteMode === 'displayed') {
      console.log(c.name, 'VML reference PDF')
      continue
    }
    if (measuredPitch === undefined || !setup.printer.startsWith('Microsoft Print to PDF')) {
      console.log(c.name, 'missing calibration')
      unknown++
      continue
    }
    const { book } = await readXlsx(readFileSync(join(directory, c.name + '.xlsx')))
    const settings = { sheet: book.sheets[0].name, ...book.sheets[0].print }
    const usedRange = c.columnWidth ? 'A1:D125' : 'A1:A125'
    // The driver printed Letter through an A4 PDF. Excel laid the breaks out on the requested stock before that adaptation.
    const requested = c.paper === 1 ? printPaperSize(settings.setup) : undefined
    const { width, height } = requested ?? pages[0]
    const baseRowUnits = Math.round(measuredPitch / pointsPerUnit)
    const hidden = new Set(c.noteMode === 'sample' ? [4] : c.hiddenRows)
    const device = {
      rowEdge: (line: number, scale: number) => {
        let before = line - 1
        for (const row of hidden) if (row < line) before--
        return before * Math.round(baseRowUnits * scale) * pointsPerUnit
      },
      ...(c.columnWidth === 50 ? {
        columnEdge: (line: number, scale: number) => (line - 1) * Math.round(wideColumnUnits * scale) * pointsPerUnit,
      } : {}),
    }
    const target = { width, height, deviceGeometry: device, paginationReserve: { height: trailingClearance } }
    const inputs = book.sheets.map((sheet) => ({
      settings: { sheet: sheet.name, ...sheet.print },
      geometry: printSheetGeometry(sheet, { font: book.font, usedRange }),
      target,
      hidden: sheet.hidden,
    }))
    if (c.workbook) {
      const plan = layoutPrintWorkbook(inputs)
      const observed = pages.map((page) => page.band?.filter((word) => !/^R\d{3}$/.test(word.text)).map((word) => word.text).join(' '))
      const predicted = plan.pages.map((page) => formatPrintHeaderFooter(page.footer?.text ?? '', {
        sheetName: page.sheet, pageNumber: page.number, pageCount: plan.totalPages,
      }).right.map((run) => 'text' in run ? run.text : '').join(''))
      total++
      if (JSON.stringify(observed) !== JSON.stringify(predicted)) {
        console.log(c.name, 'Excel:', observed, 'xlcalc:', predicted)
        different++
      }
      continue
    }
    if (c.noteMode) {
      const notes: PrintNotes = {
        cells: book.sheets[0].cells,
        labels: noteLabels,
        metrics(scale) {
          const metric = noteFont.get(scale)
          if (!metric) throw new Error('unmeasured note scale ' + scale)
          return {
            lineHeight: metric.pitch * pointsPerUnit,
            measureText: (text) => [...text].reduce((sum, char) => sum + advance(char), 0) * metric.size / 2048,
          }
        },
      }
      const plan = layoutPrintPages(settings, inputs[0].geometry, target, notes)
      const error = checkNotes(pages, plan.pages)
      total++
      if (error) {
        console.log(c.name, error)
        different++
      }
      continue
    }
    const plan = layoutPrintPages(settings, inputs[0].geometry, target)
    const observed = pages.map((page) => page.last?.row)
    const predicted = plan.pages.map((page) => {
      let row = Number(page.regions.at(-1)?.range.match(/\d+$/)?.[0])
      while (hidden.has(row)) row--
      return row
    })
    total++
    if (observed.length !== predicted.length || observed.some((row, i) => row !== undefined && row !== predicted[i])) {
      console.log(c.name, 'Excel:', observed, 'xlcalc:', predicted)
      different++
    }
  }
}
console.log(`${total - different}/${total} match, ${unknown} without calibration`)
if (different) process.exitCode = 1
