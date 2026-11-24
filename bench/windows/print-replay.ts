/** Replays the print matrix on the measured Windows Excel 16.0.20430 Microsoft Print to PDF profile. */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { formatPrintHeaderFooter, layoutPrintPages, layoutPrintWorkbook, printPaperSize, printSheetGeometry, readXlsxWithPrint } from '../../src/xlsx-print.js'

interface Trace {
  case: { name: string; height: number; columnWidth?: number; paper?: number; hiddenRows?: number[]; workbook?: object; noteMode?: string; font?: { name: string } }
  setup: { printer: string }
  pages: { width: number; height: number; pitch?: number; last?: { row: number }; band?: { text: string }[] }[]
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

let total = 0
let unknown = 0
let different = 0
for (const directory of [controlDirectory, ...also]) {
  for (const run of traces(directory)) {
    const { case: c, setup, pages } = run
    const reference = c.font?.name === 'Times New Roman' ? 'times' : `height-${c.height}`
    const measuredPitch = controlByName.get(reference)?.pages[0].pitch
    if (c.noteMode || measuredPitch === undefined || !setup.printer.startsWith('Microsoft Print to PDF')) {
      console.log(c.name, 'missing calibration')
      unknown++
      continue
    }
    const { book, print, sheetDefaults } = await readXlsxWithPrint(readFileSync(join(directory, c.name + '.xlsx')))
    const settings = print[0]
    const usedRange = c.columnWidth ? 'A1:D125' : 'A1:A125'
    const defaultBySheet = new Map(sheetDefaults.map((defaults) => [defaults.sheet, defaults]))
    const printBySheet = new Map(print.map((sheet) => [sheet.sheet, sheet]))
    // The driver printed Letter through an A4 PDF. Excel laid the breaks out on the requested stock before that adaptation.
    const requested = c.paper === 1 ? printPaperSize(settings.setup) : undefined
    const { width, height } = requested ?? pages[0]
    const baseRowUnits = Math.round(measuredPitch / pointsPerUnit)
    const hidden = new Set(c.hiddenRows)
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
      settings: printBySheet.get(sheet.name) ?? { sheet: sheet.name },
      geometry: printSheetGeometry(sheet, { ...defaultBySheet.get(sheet.name), font: book.font, usedRange }),
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
