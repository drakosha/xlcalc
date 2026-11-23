/**
 * Compare the workbook, cached results and issues of paired Excel saves.
 * Run after npm run build: node bench/windows/compare-xlsb-pairs.mjs PAIR_DIRECTORY
 */
import { readFile, readdir } from 'node:fs/promises'
import { join } from 'node:path'
import { inspect, isDeepStrictEqual } from 'node:util'
import { readXlsb } from '../../dist/xlsb.js'
import { readXlsx } from '../../dist/xlsx.js'

const directory = process.argv[2]
if (!directory) throw new Error('usage: compare-xlsb-pairs.mjs PAIR_DIRECTORY')

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

const names = await readdir(directory)
const files = new Set(names)
const summary = { pairs: 0, exact: 0, different: 0, readErrors: 0, missingXlsx: 0 }
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
    const actual = await readXlsb(await readFile(join(directory, binary)))
    if (isDeepStrictEqual(expected, actual)) summary.exact++
    else {
      summary.different++
      const differences = ['book', 'cached', 'issues'].flatMap((key) => firstDifference(expected[key], actual[key], key) ?? [])
      console.log(JSON.stringify({ pair: binary, differences }))
    }
  } catch (error) {
    summary.readErrors++
    console.log(JSON.stringify({ pair: binary, error: String(error) }))
  }
}
console.log(JSON.stringify(summary))
if (summary.different || summary.readErrors || summary.missingXlsx) process.exitCode = 1
