/**
 * Saves the same open book as xlsx and xlsb. Pass tracked fixture paths after the output directory,
 * or omit them for the small probe: XLWIN_HOST=… XLWIN_CODEPAGE=any npx tsx bench/windows/xlsb-pairs.ts OUT [XLSX…]
 */
import { createHash } from 'node:crypto'
import { existsSync, mkdirSync, readFileSync, renameSync, writeFileSync } from 'node:fs'
import { basename, join } from 'node:path'
import { writeXlsx } from '../../src/xlsx.js'
import { metaOf, runJob, withWindowsExcel, type Book, type Step } from './remote.js'

const [, , out, ...sources] = process.argv
if (!out) throw new Error('usage: xlsb-pairs.ts OUT [XLSX…] [--formulas] [--files-from PATH]')
mkdirSync(out, { recursive: true })

const locale = process.env.XLWIN_LOCALE ?? 'ru-RU'
const probe = join(out, 'probe.source.xlsx')
const save: Step = { do: 'saveXlsbPair' }
const files: string[] = []
for (let i = 0; i < sources.length; i++) {
  const source = sources[i]
  if (source === '--formulas') continue
  if (source === '--files-from') {
    const manifest = sources[++i]
    if (!manifest) throw new Error('--files-from takes a text file with one workbook path per line')
    files.push(...readFileSync(manifest, 'utf8').split(/\r?\n/).filter(Boolean))
  } else files.push(source)
}
const cases: { id: string; book: Book }[] = files.map((file) => {
  const suffix = createHash('sha256').update(file).digest('hex').slice(0, 12)
  return { id: `${basename(file, '.xlsx')}-${suffix}`, book: { file, steps: [save] } }
})
if (!files.length || sources.includes('--formulas')) {
  const input = await writeXlsx({
    book: {
      sheets: [{ name: 'Data', cells: { A1: { value: 3 }, A2: { formula: 'SUM(A1,2)' } } }],
      names: [{ name: 'Amount', ref: 'Data!$A$1' }],
    },
  })
  writeFileSync(probe, input.data)
  if (!sources.length) cases.push({ id: 'probe', book: { file: probe, steps: [save] } })
  if (sources.includes('--formulas')) {
    const formulas = [
      ['C1', '=SEQUENCE(3)'], ['D1', '=ROWS(C1#)'], ['D2', '=@C1:C3'],
      ['D3', '=LET(aValue,5,aValue+1)'], ['D4', '=LAMBDA(param,param+1)(10)'],
      ['E1', '=FILTER(A1:A2,A1:A2>10)'], ['E2', '=SEQUENCE(2)'], ['E3', '=1'],
      ['F1', '=0.123456789012345+0.12345678901234567'],
      ['F2', '=1E+20+1E-7+0.000001+0.00000012'],
      ['F3', '=-0+-12.5+( (42) )'], ['F4', '=2^53+9007199254740992'],
      ['F5', '=1.0000000000000002+999999999999999'],
      ['G1', '=FIELDVALUE(A1,"Value")'],
    ]
    const steps = formulas.map(([cell, value]) => ({ do: 'set', sheet: 'Data', cell, kind: 'f2', value }))
    cases.push({ id: 'formulas', book: { file: probe, steps: [...steps, save] } })
  }
}

const available = ({ id }: { id: string }) => existsSync(join(out, `${id}.xlsx`)) && existsSync(join(out, `${id}.xlsb`))
const pending = cases.filter((item) => !available(item))

function publish(name: string, data: string | Uint8Array): void {
  const destination = join(out, name)
  const temporary = `${destination}.${process.pid}.tmp`
  writeFileSync(temporary, data)
  renameSync(temporary, destination)
}

const batchSize = 25
const elapsed = Date.now()
let failures = 0
while (pending.length) {
  const batch = pending.splice(0, batchSize)
  const start = Date.now()
  await withWindowsExcel(locale, async (env) => {
    // A second run might have saved these pairs while this one waited for Excel.
    const missing = batch.filter((item) => !available(item))
    if (!missing.length) return
    const answers = await runJob('xlsb-pairs', missing.map(({ book }) => book), 45_000 + 4_000 * missing.length)
    const meta = metaOf(env, locale)
    const firstUnfinished = answers.indexOf(undefined)
    answers.forEach((answer, i) => {
      const item = missing[i]
      if (firstUnfinished >= 0 && i > firstUnfinished) {
        pending.push(item)
        return
      }
      const refused = !answer ? 'Excel stopped while opening or saving this book' : 'rejected' in answer ? answer.rejected : undefined
      if (answer && 'answers' in answer && answer.saved && answer.savedXlsb) {
        publish(`${item.id}.xlsx`, answer.saved)
        publish(`${item.id}.xlsb`, answer.savedXlsb)
      } else {
        failures++
        console.error(`${item.id}: ${refused ?? 'Excel did not save both files'}`)
      }
      const details = answer && 'answers' in answer ? { steps: answer.answers } : { refused }
      publish(`${item.id}.json`, JSON.stringify({ meta, seconds: answer?.seconds, ...details }, null, 2) + '\n')
    })
  })
  console.log(`${batch.length} books in ${((Date.now() - start) / 1000).toFixed(1)} s, ${pending.length} left`)
}
console.log(`${cases.length - failures}/${cases.length} pairs available, ${((Date.now() - elapsed) / 1000).toFixed(1)} s`)
if (failures) process.exitCode = 1
