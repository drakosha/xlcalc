// node bench/compare/run.mjs [--part engines,structure,xlsx,memory] [--out dir]: each case and library in its own Node
// process, one after another; the engines' cases as bench/hyperformula runs them (a phase limited to 30 s, a 1024 MiB heap),
// the others with 60 s and 4096 MiB. Raw results with the load average of each case and `uptime` at the start and end.
import { execFileSync, fork } from 'node:child_process'
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { cpus, loadavg, totalmem } from 'node:os'
import { resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { WORKLOADS } from '../hyperformula/workloads.mjs'

const args = process.argv.slice(2)
const parts = args.includes('--part') ? args[args.indexOf('--part') + 1].split(',') : ['engines', 'structure', 'xlsx', 'memory']
const out = resolve(args.includes('--out') ? args[args.indexOf('--out') + 1] : 'bench/compare/results')
const root = fileURLToPath(new URL('../../', import.meta.url))
const engineRoot = process.env.XLCALC_BENCH_ROOT ? resolve(process.env.XLCALC_BENCH_ROOT) : root
const here = (file) => fileURLToPath(new URL(file, import.meta.url))
const version = (name) => JSON.parse(readFileSync(here(`node_modules/${name}/package.json`), 'utf8')).version
const uptime = () => execFileSync('uptime', { encoding: 'utf8' }).trim()

const jobs = [
  ...(parts.includes('engines') ? WORKLOADS.map((w) => ({ part: 'engines', file: 'iron-child.mjs', args: [w.id, 'ironcalc'] })) : []),
  ...(parts.includes('structure')
    ? [
        ...['business-3400', 'prefix-sum'].flatMap((id) => ['xl-closures', 'xl-eval', 'hf-default', 'ironcalc'].map((e) => ({ part: 'structure', file: 'structure-child.mjs', args: [id, e] }))),
        ...['xl-closures', 'xl-eval'].map((e) => ({ part: 'structure', file: 'structure-child.mjs', args: ['business-34000', e] })),
      ]
    : []),
  ...(parts.includes('xlsx') ? ['3400', '34000'].flatMap((rows) => ['xlcalc', 'exceljs', 'sheetjs'].map((lib) => ({ part: 'xlsx', file: 'xlsx-child.mjs', args: [rows, lib] }))) : []),
  ...(parts.includes('memory') ? WORKLOADS.map((w) => ({ part: 'memory', file: 'iron-memory.mjs', args: [w.id] })) : []),
]

const report = {
  started: new Date().toISOString(), uptimeStart: uptime(), node: process.version, engineRoot,
  engineCommit: execFileSync('git', ['rev-parse', 'HEAD'], { cwd: engineRoot, encoding: 'utf8' }).trim(),
  versions: { hyperformula: version('hyperformula'), ironcalc: version('@ironcalc/wasm'), exceljs: version('exceljs'), xlsx: version('xlsx') },
  machine: { cpu: cpus()[0].model, memoryGiB: totalmem() / 2 ** 30 },
  results: [],
}
mkdirSync(out, { recursive: true })
const save = () => writeFileSync(resolve(out, 'raw.json'), JSON.stringify(report, null, 1) + '\n')

function execute(job) {
  return new Promise((done) => {
    const [phaseLimitMs, heapMb] = job.part === 'engines' ? [30_000, 1024] : [60_000, 4096]
    const child = fork(here(job.file), job.args, { cwd: root, silent: true, execArgv: ['--expose-gc', `--max-old-space-size=${heapMb}`], env: { ...process.env, XLCALC_BENCH_KEEP: '1' } })
    let timer, phase = 'startup', received, timedOut = false, stderr = ''
    const arm = () => {
      clearTimeout(timer)
      timer = setTimeout(() => { timedOut = true; child.kill('SIGKILL') }, phaseLimitMs)
    }
    arm()
    child.stderr.on('data', (chunk) => { stderr = (stderr + chunk).slice(-3000) })
    child.stdout.resume()
    child.on('message', (m) => {
      if (m.type === 'phase') { phase = m.name; arm() }
      if (m.type === 'result') received = m.result
    })
    child.on('exit', (code, signal) => {
      clearTimeout(timer)
      done(received ?? { status: timedOut ? 'timeout' : 'error', phase, code, signal, stderr })
    })
  })
}

for (const job of jobs) {
  const loadBefore = loadavg()
  const result = await execute(job)
  report.results.push({ ...job, loadBefore, loadAfter: loadavg(), ...result })
  save()
  console.log(JSON.stringify({ part: job.part, args: job.args, status: result.status, phase: result.phase, open: result.coldMedianMs, edits: result.edits && Object.fromEntries(Object.entries(result.edits).map(([k, e]) => [k, e.medianMs])), operations: result.operations && Object.fromEntries(Object.entries(result.operations).map(([k, e]) => [k, e.medianMs])), readMs: result.readMs, writeMs: result.writeMs, heapMb: result.retainedHeapMb ?? result.heapMb, wasmMb: result.wasmMb, differences: result.differences ?? result.details?.differences }))
}
report.finished = new Date().toISOString()
report.uptimeEnd = uptime()
save()
