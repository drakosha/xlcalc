import { Session } from 'node:inspector/promises'
import { mkdirSync, writeFileSync } from 'node:fs'
import { execFileSync } from 'node:child_process'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { workload } from './workloads.mjs'
import { open, close, prepareEdit, reference, verify } from './shared.mjs'
const [id, engine = 'xl-eval', countText = '100', directory = '/tmp/xlcalc-cpu-profiles'] = process.argv.slice(2)
const spec = workload(id), operation = spec.operations[0], count = Number(countText)
if (!['xl-closures', 'xl-eval'].includes(engine) || !Number.isInteger(count) || count < 1 || count > 1000) throw new Error('Expected an xlcalc engine and 1 to 1000 edits')
const root = fileURLToPath(new URL('../../', import.meta.url))
const revision = execFileSync('git', ['rev-parse', 'HEAD'], { cwd: root, encoding: 'utf8' }).trim()
const startedAt = new Date().toISOString()
const session = open(engine, spec.book)
for (let i = 0; i < 30; i++) prepareEdit(engine, session, operation.edits(i))()
global.gc()
const profiler = new Session()
profiler.connect()
await profiler.post('Profiler.enable')
await profiler.post('Profiler.setSamplingInterval', { interval: 100 })
await profiler.post('Profiler.start')
const started = performance.now()
for (let i = 0; i < count; i++) prepareEdit(engine, session, operation.edits(i))()
const elapsedMs = performance.now() - started
const { profile } = await profiler.post('Profiler.stop')
profiler.disconnect()
const checked = verify(engine, session, spec.book, reference(spec.book, operation.edits(count - 1)))
if (checked.differences) throw new Error(JSON.stringify(checked))
const nodes = new Map(profile.nodes.map(n => [n.id, n]))
const parent = new Map()
for (const n of profile.nodes) for (const child of n.children ?? []) parent.set(child, n.id)
const self = new Map(), inclusive = new Map()
for (let i = 0; i < profile.samples.length; i++) {
 let id = profile.samples[i]
 const time = profile.timeDeltas[i] / 1000
 self.set(id, (self.get(id) ?? 0) + time)
 while (id !== undefined) { inclusive.set(id, (inclusive.get(id) ?? 0) + time); id = parent.get(id) }
}
const rows = map => [...map].map(([id, ms]) => {
 const f = nodes.get(id).callFrame
 return { fn: f.functionName || '(anonymous)', file: f.url.replace(`file://${root}`, ''), line: f.lineNumber + 1, ms }
}).sort((a,b) => b.ms - a.ms).slice(0, 25)
const dir = resolve(directory)
mkdirSync(dir, {recursive:true})
writeFileSync(`${dir}/${id}-${engine}.cpuprofile`,JSON.stringify(profile))
const summary = { revision, startedAt, node: process.version, id, engine, count, elapsedMs, checked, self: rows(self), inclusive: rows(inclusive) }
writeFileSync(`${dir}/${id}-${engine}.json`, JSON.stringify(summary,null,2))
console.log(JSON.stringify(summary))
close(engine, session)
