/**
 * The engine's answers to a conditional format set as its predictions, written before Excel runs the set:
 * `npx tsx bench/coverage/cond/predict.ts <set> [run]` writes `cond/<set>.predicted.json` (`<set>.<run>.predicted.json`
 * for a later run, such as Excel for Windows's after the Mac's answers changed the engine), for each case the engine's refusal or
 * each sheet's rules after the operation and what the read cells show, at noon UTC of the day it runs (TODAY of time
 * periods). A case whose rules Excel makes (Windows) is predicted on its book with the rules it is expected to make.
 */
import { writeFileSync } from 'node:fs'
import { engineCase, type CondSet } from './cases.js'

const [set, run] = process.argv.slice(2)
if (!/^[a-z0-9-]+$/.test(set ?? '') || (run !== undefined && !/^[a-z0-9-]+$/.test(run))) throw new Error(`set and run names: [a-z0-9-], not ${set} ${run}`)
const mod = (await import(`./${set}.js`)) as CondSet
const built = mod.buildCondCases()
const day = new Date().toLocaleDateString('sv-SE')
const clock = { now: `${day}T12:00:00Z`, seed: 1 }
const out: Record<string, unknown> = { day }
for (const c of built.cases) out[c.id] = engineCase(built, c, clock, c.make && mod.expectedBook ? mod.expectedBook(c) : undefined)
writeFileSync(new URL(`./${set}${run ? `.${run}` : ''}.predicted.json`, import.meta.url), JSON.stringify(out, null, 1) + '\n')
console.log(`${set}: ${built.cases.length} cases predicted on ${day}`)
