import { workload } from './workloads.mjs'
import { HyperFormula, close, isHf, measure, median, open, p95, prepareEdit, prepareInput, reference, verify } from './shared.mjs'

const [id, engine] = process.argv.slice(2)
const spec = workload(id)
if (HyperFormula.version !== '3.5.0') throw new Error(`Expected HyperFormula 3.5.0, got ${HyperFormula.version}`)
const send = (message) => process.send?.(message)
const phase = (name) => send({ type: 'phase', name })
const prepared = prepareInput(engine, spec.book)
phase('reference')
const initial = reference(spec.book)
const metadata = {
  cells: spec.book.sheets.reduce((n, s) => n + Object.keys(s.cells).length, 0),
  formulas: spec.book.sheets.reduce((n, s) => n + Object.values(s.cells).filter((c) => c.formula !== undefined).length, 0),
  hash: initial.hash, inputHash: initial.inputHash,
}
const samples = [], edits = {}, verification = []
let held, retainedHeapMb
// XLCALC_BENCH_KEEP=1 (bench/compare): differing values are recorded and timing goes on, for a time given for reference
const keep = process.env.XLCALC_BENCH_KEEP === '1'
let differing

function check(session, expected) {
  phase('verify')
  const result = verify(engine, session, spec.book, expected)
  verification.push({ checked: result.checked, maxRelativeError: result.maxRelativeError })
  if (result.differences && keep) {
    if (!differing || result.differences > differing.differences) differing = result
  } else if (result.differences) {
    const error = new Error('Computed values differ from a full calculation')
    error.details = result
    throw error
  }
}

try {
  const probe = spec.hfProbe && isHf(engine)
  for (let i = 0; i < (probe ? 1 : 6); i++) {
    global.gc()
    const heapBefore = process.memoryUsage().heapUsed
    phase(i ? 'build' : 'warmup-build')
    const measured = measure(() => open(engine, spec.book, prepared))
    held = measured.value
    if (i || probe) samples.push({ wallMs: measured.wallMs, cpuMs: measured.cpuMs })
    if (i === (probe ? 0 : 5)) {
      global.gc()
      retainedHeapMb = (process.memoryUsage().heapUsed - heapBefore) / 2 ** 20
    }
    check(held, initial)
    if (!probe && i < 5) { close(engine, held); held = undefined }
  }

  if (!(spec.hfProbe && isHf(engine))) for (const operation of spec.operations) {
    phase('reference')
    const expected = [reference(spec.book, operation.edits(0)), initial]
    const measured = []
    let firstEditMs
    for (let i = 0; i < 22; i++) {
      const apply = prepareEdit(engine, held, operation.edits(i))
      phase('edit')
      const timing = measure(apply)
      if (i === 0) firstEditMs = timing.wallMs
      if (i >= 2) measured.push({ wallMs: timing.wallMs, cpuMs: timing.cpuMs })
      check(held, expected[i % 2])
    }
    edits[operation.id] = {
      firstEditMs, medianMs: median(measured.map((m) => m.wallMs)), p95Ms: p95(measured.map((m) => m.wallMs)),
      medianCpuMs: median(measured.map((m) => m.cpuMs)), samples: measured,
    }
  }
  send({ type: 'result', result: {
    status: differing ? 'mismatch' : 'ok', ...metadata, ...(differing ? { details: differing } : {}), retainedHeapMb, addressMappingPolicy: isHf(engine) ? held.getConfig().chooseAddressMappingPolicy.constructor.name : undefined, coldMedianMs: median(samples.map((s) => s.wallMs)),
    coldMedianCpuMs: median(samples.map((s) => s.cpuMs)), samples, edits,
    verification: { comparisons: verification.length, checkedCells: verification.reduce((n, v) => n + v.checked, 0), maxRelativeError: Math.max(...verification.map((v) => v.maxRelativeError)) },
    probeOnly: spec.hfProbe && isHf(engine),
  } })
} catch (error) {
  send({ type: 'result', result: { status: error.details ? 'mismatch' : 'error', ...metadata, samples, edits, message: error.message, details: error.details } })
} finally {
  if (held) close(engine, held)
  process.disconnect?.()
}
