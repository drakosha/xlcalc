/**
 * The engine's answers to a conditional format set as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/cond/predict.ts <set> [run]` writes `cond/<set>.predicted.json` (`<set>.<run>.predicted.json`
 * for a later run, such as Excel for Windows's after the Mac's answers changed the engine), for each case the engine's refusal or
 * each sheet's rules after the operation and what the read cells show, at noon UTC of the day it runs (TODAY of time
 * periods). A case whose rules Excel makes (Windows) is predicted on its book with the rules it is expected to make; a
 * raw case on the book the reader reads from its file, with what the reader reports.
 */
import { setModule, writePrediction } from '../stored.js'
import { readXlsx } from '../../../src/xlsx.js'
import { condFile, engineCase, type CondSet } from './cases.js'

const [set, run] = process.argv.slice(2)
if (run !== undefined && !/^[a-z0-9-]+$/.test(run)) throw new Error(`run name: [a-z0-9-], not ${run}`)
const mod = await setModule<CondSet>('cond', set)
const built = mod.buildCondCases()
const day = new Date().toLocaleDateString('sv-SE')
const clock = { now: `${day}T12:00:00Z`, seed: 1 }
const out: Record<string, unknown> = { day }
for (const c of built.cases) {
  const read = c.raw ? await readXlsx(await condFile(built, c)) : undefined
  const issues = read?.issues.map((i) => `${i.code} ${i.sheet ?? ''} ${(i.cells ?? []).join(' ')}: ${i.message}`)
  out[c.id] = { ...engineCase(built, c, clock, read?.book ?? (c.make && mod.expectedBook ? mod.expectedBook(c) : undefined)), ...(issues ? { issues } : {}) }
}
writePrediction(import.meta.url, `${set}${run ? `.${run}` : ''}`, out)
console.log(`${set}: ${built.cases.length} cases predicted on ${day}`)
