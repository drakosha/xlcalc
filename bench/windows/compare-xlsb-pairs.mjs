/**
 * Compare the workbook, cached results and issues of paired Excel saves.
 * Run after npm run build: node bench/windows/compare-xlsb-pairs.mjs PAIR_DIRECTORY [--recalculate]
 */
import { readFile, readdir } from 'node:fs/promises'
import { join } from 'node:path'
import { inspect, isDeepStrictEqual } from 'node:util'
import { readXlsb } from '../../dist/xlsb.js'
import { records } from '../../dist/xlsb/records.js'
import { readXlsx } from '../../dist/xlsx.js'
import { relsOf } from '../../dist/xlsx/read/package.js'
import { openArchive } from '../../dist/xlsx/zip/read.js'

const [, , directory, option] = process.argv
if (!directory || (option !== undefined && option !== '--recalculate')) {
  throw new Error('usage: compare-xlsb-pairs.mjs PAIR_DIRECTORY [--recalculate]')
}

const calculation = { comparableCells: 0, discrepancies: 0, xlsxCannotCalculate: 0, xlsbCannotCalculate: 0 }
let calculate
if (option === '--recalculate') {
  const [{ compute, registerPivot }, { pivotEngine }] = await Promise.all([
    import('../../dist/index.js'), import('../../dist/pivot/index.js'),
    import('../../dist/locales/ru-RU.js'),
  ])
  registerPivot(pivotEngine)
  const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
  calculate = (read) => compute(read.book, { cached: read.cached, clock })
}

function firstDifference(expected, actual, path) {
  if (isDeepStrictEqual(expected, actual)) return undefined
  if (expected && actual && typeof expected === 'object' && typeof actual === 'object' && Array.isArray(expected) === Array.isArray(actual)) {
    const keys = new Set([...Object.keys(expected), ...Object.keys(actual)])
    for (const key of keys) {
      const childPath = `${path}.${key}`
      if (Object.hasOwn(expected, key) !== Object.hasOwn(actual, key)) {
        return { path: childPath, xlsx: preview(expected[key]), xlsb: preview(actual[key]) }
      }
      const difference = firstDifference(expected[key], actual[key], childPath)
      if (difference) return difference
    }
  }
  return { path, xlsx: preview(expected), xlsb: preview(actual) }
}

function preview(value) {
  return inspect(value, { depth: 0, maxArrayLength: 8, maxStringLength: 360 })
}

async function onlyMissingSparklines(expected, actual, binaryData) {
  const issues = expected.issues.filter(({ code }) => code !== 'sparkline')
  if (issues.length === expected.issues.length || !isDeepStrictEqual({ ...expected, issues }, actual)) return false

  // Verify the file itself so a reader regression cannot pass as an Excel save difference.
  const archive = openArchive(binaryData)
  const workbook = (await relsOf(archive, '')).find(({ kind, external }) => kind === 'officeDocument' && !external)
  if (!workbook) return false
  const worksheets = (await relsOf(archive, workbook.target)).filter(({ kind, external }) => kind === 'worksheet' && !external)
  for (const { target } of worksheets) {
    if (!target.toLowerCase().endsWith('.bin')) return false
    const data = await archive.bytes(target)
    if (!data) return false
    for (const { type } of records(data, target)) {
      if (type === 1058) return false // BrtBeginSparklineGroups
    }
  }
  return worksheets.length > 0
}

/** Compare saved values wherever the engine calculates the Excel result from the XLSX half. */
function checkCalculation(pair, expected, actual, same) {
  let expectedResult
  try {
    expectedResult = calculate(expected)
  } catch (error) {
    calculation.xlsxCannotCalculate++
    console.log(JSON.stringify({ pair, xlsxCalculationError: String(error) }))
    return
  }
  let actualResult = expectedResult
  if (!same) {
    try {
      actualResult = calculate(actual)
    } catch (error) {
      calculation.xlsbCannotCalculate++
      console.log(JSON.stringify({ pair, calculationError: String(error) }))
      return
    }
  }
  let example
  for (const [index, sheet] of expected.cached.sheets.entries()) {
    for (const [address, saved] of Object.entries(sheet.cells)) {
      if (!isDeepStrictEqual(expectedResult.sheets[index]?.cells[address]?.value, saved.value)) continue
      calculation.comparableCells++
      const value = actualResult.sheets[index]?.cells[address]?.value
      if (isDeepStrictEqual(value, saved.value)) continue
      calculation.discrepancies++
      example ??= { sheet: sheet.name, address, saved: saved.value, xlsb: value }
    }
  }
  if (example) console.log(JSON.stringify({ pair, firstCalculationDifference: example }))
}

const names = await readdir(directory)
const files = new Set(names)
const summary = { pairs: 0, exact: 0, expectedSparklineLoss: 0, different: 0, readErrors: 0, missingXlsx: 0 }
for (const binary of names.filter((name) => name.endsWith('.xlsb')).sort()) {
  const spreadsheet = binary.slice(0, -1) + 'x'
  if (!files.has(spreadsheet)) {
    summary.missingXlsx++
    console.log(JSON.stringify({ pair: binary, error: 'missing XLSX save' }))
    continue
  }
  summary.pairs++
  try {
    const expected = await readXlsx(await readFile(join(directory, spreadsheet)))
    const binaryData = await readFile(join(directory, binary))
    const actual = await readXlsb(binaryData)
    const same = isDeepStrictEqual(expected, actual)
    if (same) summary.exact++
    else if (await onlyMissingSparklines(expected, actual, binaryData)) {
      summary.expectedSparklineLoss++
      console.log(JSON.stringify({ pair: binary, expectedDifference: 'sparkline groups absent from the XLSB save' }))
    } else {
      summary.different++
      const differences = ['book', 'cached', 'issues'].flatMap((key) => firstDifference(expected[key], actual[key], key) ?? [])
      console.log(JSON.stringify({ pair: binary, differences }))
    }
    if (calculate) checkCalculation(binary, expected, actual, same)
  } catch (error) {
    summary.readErrors++
    console.log(JSON.stringify({ pair: binary, error: String(error) }))
  }
}
console.log(JSON.stringify(summary))
if (calculate) console.log(JSON.stringify({ calculation }))
if (summary.different || summary.readErrors || summary.missingXlsx || calculation.discrepancies || calculation.xlsbCannotCalculate) {
  process.exitCode = 1
}
