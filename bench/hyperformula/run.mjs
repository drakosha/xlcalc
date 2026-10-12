import { fork, execFileSync } from 'node:child_process'
import { mkdirSync, writeFileSync, readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { resolve } from 'node:path'
import { cpus, loadavg, platform, arch, totalmem } from 'node:os'
import { WORKLOADS } from './workloads.mjs'

const args = process.argv.slice(2)
const selected = args.includes('--case') ? args[args.indexOf('--case') + 1].split(',') : WORKLOADS.map((w) => w.id)
if (selected.some((id) => !WORKLOADS.some((w) => w.id === id))) throw new Error('Unknown workload in --case')
const requestedEngines = args.includes('--engine') ? args[args.indexOf('--engine') + 1].split(',') : undefined
if (requestedEngines?.some((e) => !['xl-closures', 'xl-eval', 'hf-default', 'hf-index', 'hf-sparse', 'hf-auto'].includes(e))) throw new Error('Unknown --engine')
const out = resolve(args.includes('--out') ? args[args.indexOf('--out') + 1] : 'bench/hyperformula/results')
const root = fileURLToPath(new URL('../../', import.meta.url))
const engineRoot = process.env.XLCALC_BENCH_ROOT ? resolve(process.env.XLCALC_BENCH_ROOT) : root
const revision = dir => { try { return execFileSync('git', ['rev-parse', 'HEAD'], { cwd: dir, encoding: 'utf8', stdio: ['ignore','pipe','pipe'] }).trim() } catch { return readFileSync(resolve(dir,'revision.txt'),'utf8').trim() } }
const phaseLimitMs = 30_000, heapLimitMb = 1024
const report = {
  engineCommit: revision(engineRoot), harnessCommit: revision(root),
  started: new Date().toISOString(), node: process.version, hyperformula: '3.5.0',
  machine: { platform: platform(), arch: arch(), cpu: cpus()[0].model, memoryGiB: totalmem() / 2 ** 30 },
  protocol: { warmups: 1, coldSamples: 5, warmupEdits: 2, editSamples: 20, phaseLimitMs, heapLimitMb,
    numericTolerance: '1e-9 * max(1, abs(a), abs(b)); other types and error codes exact',
    timed: 'Session construction and synchronous edits including their returned changes. Input generation/conversion, formatting, result extraction, verification and hashes excluded.',
    isolation: 'One sequential child process per workload and engine; explicit GC before each fresh session. HF 3.5.0 defaults to AlwaysDense. Column index is tested separately on lookup; AlwaysSparse and automatic mapping (threshold 0.5) separately on the sparse book. The 34000-row HF business case is a single bounded cold probe.' },
  workloads: [],
}
mkdirSync(out, { recursive: true })

function execute(id, engine) {
  return new Promise((done) => {
    const started = performance.now()
    const child = fork(new URL('./child.mjs', import.meta.url), [id, engine], { cwd: root, silent: true, execArgv: ['--expose-gc', `--max-old-space-size=${heapLimitMb}`] })
    let timer, phase = 'startup', phaseStarted = started, received, timedOut = false, stderr = ''
    const arm = () => {
      clearTimeout(timer)
      timer = setTimeout(() => { timedOut = true; child.kill('SIGKILL') }, phaseLimitMs)
    }
    arm()
    child.stderr.on('data', (chunk) => { stderr = (stderr + chunk).slice(-3000) })
    child.stdout.resume()
    child.on('message', (m) => {
      if (m.type === 'phase') { phase = m.name; phaseStarted = performance.now(); arm() }
      if (m.type === 'result') { received = m.result; clearTimeout(timer) }
    })
    child.on('error', (error) => { clearTimeout(timer); done({ status: 'error', message: error.message }) })
    child.on('exit', (code, signal) => {
      clearTimeout(timer)
      done(received ?? { status: timedOut ? 'timeout' : /heap out of memory/i.test(stderr) ? 'out_of_memory' : 'error', phase, phaseElapsedMs: performance.now() - phaseStarted, elapsedMs: performance.now() - started, code, signal, stderr })
    })
  })
}

function save() {
  writeFileSync(resolve(out, 'raw.json'), JSON.stringify(report, null, 2) + '\n')
  const rows = [['workload', 'engine', 'status', 'operation', 'median_wall_ms', 'p95_wall_ms', 'median_cpu_ms', 'first_edit_ms', 'retained_heap_mb']]
  for (const w of report.workloads) for (const r of w.engines) {
    rows.push([w.id, r.engine, r.status, 'open', r.coldMedianMs ?? '', '', r.coldMedianCpuMs ?? '', '', r.retainedHeapMb ?? ''])
    for (const [name, e] of Object.entries(r.edits ?? {})) rows.push([w.id, r.engine, r.status, name, e.medianMs, e.p95Ms, e.medianCpuMs, e.firstEditMs, r.retainedHeapMb ?? ''])
  }
  writeFileSync(resolve(out, 'summary.csv'), rows.map((r) => r.join(',')).join('\n') + '\n')
}

for (const [index, id] of selected.entries()) {
  const spec = WORKLOADS.find((w) => w.id === id)
  const item = { id, title: spec.title, engines: [] }
  report.workloads.push(item)
  const engines = requestedEngines ?? ['xl-closures', 'hf-default', 'xl-eval', ...(id === 'lookup' ? ['hf-index'] : []), ...(id === 'sparse' ? ['hf-sparse', 'hf-auto'] : [])]
  if (index % 2) engines.reverse()
  for (const engine of engines) {
    const loadBefore = loadavg()
    const result = await execute(id, engine)
    item.engines.push({ engine, loadBefore, loadAfter: loadavg(), ...result })
    save()
    console.log(JSON.stringify({ workload: id, engine, status: result.status, openMs: result.coldMedianMs, edits: Object.fromEntries(Object.entries(result.edits ?? {}).map(([k, e]) => [k, { medianMs: e.medianMs, p95Ms: e.p95Ms }])), phase: result.phase, differences: result.details?.differences }))
  }
}
report.finished = new Date().toISOString()
save()
if (report.workloads.some((w) => w.engines.some((e) => ['mismatch', 'error'].includes(e.status)))) process.exitCode = 1
