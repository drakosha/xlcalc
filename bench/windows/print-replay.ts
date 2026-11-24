/** Replays the print matrix on the measured Windows Excel 16.0.20430 Microsoft Print to PDF profile. */
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { layoutPrintPages, printPaperSize, printSheetGeometry, readXlsxWithPrint } from '../../src/xlsx-print.js'

interface Trace {
  case: { name: string; height: number; columnWidth?: number; paper?: number; hiddenRows?: number[]; font?: { name: string } }
  setup: { printer: string }
  pages: { width: number; height: number; pitch?: number; last?: { row: number } }[]
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
    if (measuredPitch === undefined || !setup.printer.startsWith('Microsoft Print to PDF')) {
      console.log(c.name, 'missing calibration')
      unknown++
      continue
    }
    const { book, print, sheetDefaults } = await readXlsxWithPrint(readFileSync(join(directory, c.name + '.xlsx')))
    const settings = print[0]
    const geometry = printSheetGeometry(book.sheets[0], {
      ...sheetDefaults[0], font: book.font, usedRange: c.columnWidth ? 'A1:D125' : 'A1:A125',
    })
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
    const plan = layoutPrintPages(settings, geometry, {
      width, height, deviceGeometry: device, paginationReserve: { height: trailingClearance },
    })
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
