// node --expose-gc --max-old-space-size=8192 workload-pair.mjs <base> <branch> [workload ...]: bench/hyperformula's
// workloads in two builds of xlcalc in one process, taking turns, the order changed each round: opening (median of 7)
// and each edit operation (median and p95 of 30, the two states alternating), closures and eval, display false. After
// the rounds the branch's values are checked against the base's full calculation.
import { WORKLOADS, workload } from '../hyperformula/workloads.mjs'

const [base, branch, ...ids] = process.argv.slice(2)
const libs = await Promise.all([base, branch].map((r) => import(r + '/dist/index.js')))
const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const names = ['base', 'branch']
const f = (v) => (v < 1 ? v.toFixed(3) : v < 10 ? v.toFixed(2) : v.toFixed(1))
const at = (xs, q) => [...xs].sort((a, b) => a - b)[Math.min(xs.length - 1, Math.floor(xs.length * q))]
const pair = (round) => (round % 2 ? [1, 0] : [0, 1])
const line = (what, xs, q = 0.5) => `${what} ${names.map((n, i) => `${n} ${f(at(xs[i], q))}`).join(' / ')} (${(at(xs[0], q) / at(xs[1], q)).toFixed(2)}x)`

for (const id of ids.length ? ids : WORKLOADS.map((w) => w.id)) {
  const spec = workload(id)
  for (const execution of ['closures', 'eval']) {
    const options = { clock, execution, display: false }
    const open = [[], []]
    for (let round = 0; round < 8; round++)
      for (const i of pair(round)) {
        global.gc()
        const t = performance.now()
        const s = libs[i].createSession(spec.book, options)
        if (round) open[i].push(performance.now() - t)
        s.close()
      }
    const sessions = libs.map((l) => l.createSession(spec.book, options))
    const out = [line('open', open)]
    for (const op of spec.operations) {
      const ms = [[], []]
      for (let round = 0; round < 32; round++)
        for (const i of pair(round)) {
          const edits = op.edits(round)
          const t = performance.now()
          sessions[i].setCells(edits)
          if (round >= 2) ms[i].push(performance.now() - t)
        }
      out.push(`${op.id}: ${line('p50', ms)}, ${line('p95', ms, 0.95)}`)
    }
    const [a, b] = sessions.map((s) => s.getHash())
    if (a !== b) throw new Error(`${id} ${execution}: the builds' values differ`)
    sessions.forEach((s) => s.close())
    console.log(`${id}, ${execution}: ${out.join('; ')} ms`)
  }
}
