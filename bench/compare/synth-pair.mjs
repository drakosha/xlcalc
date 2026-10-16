// node --max-old-space-size=8192 synth-pair.mjs <base> <branch> [rows ...]: two builds of xlcalc in one process, taking
// turns, the order changed each round, on the synthetic book: a cold computation (median of 5 each, both executions),
// 200 cell edits in a session (p50 and p95), and rows inserted and deleted at row 10 with undo and redo (median of 7,
// display false and true). Noise of a busy host falls on both builds alike.
const [base, branch, ...sizes] = process.argv.slice(2)
const libs = await Promise.all([base, branch].map(async (r) => ({ ...(await import(r + '/dist/index.js')), ...(await import(r + '/dist/synth.js')) })))
const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
const names = ['base', 'branch']
const f = (v) => (v < 10 ? v.toFixed(3) : v.toFixed(1))
const at = (xs, q) => [...xs].sort((a, b) => a - b)[Math.min(xs.length - 1, Math.floor(xs.length * q))]
const pair = (round) => (round % 2 ? [1, 0] : [0, 1])

for (const rows of (sizes.length ? sizes : ['3400', '34000']).map(Number)) {
  const synth = libs[0].synthBook({ seed: 1, rows })
  for (const execution of ['closures', 'eval']) {
    const cold = [[], []]
    for (let round = 0; round < 10; round++)
      for (const i of pair(round)) {
        const t = performance.now()
        libs[i].compute(synth.book, { clock, execution })
        cold[i].push(performance.now() - t)
      }
    const sessions = libs.map((l) => l.createSession(synth.book, { clock, display: false, execution }))
    const edits = [[], []]
    for (const [k, e] of libs[0].classAEdits(synth, 3, 200).entries())
      for (const i of pair(k)) {
        const t = performance.now()
        sessions[i].setCells(Array.isArray(e) ? e : [e])
        edits[i].push(performance.now() - t)
      }
    sessions.forEach((s) => s.close())
    console.log(`${rows} rows, ${execution}: cold ${names.map((n, i) => `${n} ${f(at(cold[i], 0.5))}`).join(' / ')} ms; edit p50 ${names.map((n, i) => `${n} ${f(at(edits[i], 0.5))}`).join(' / ')}, p95 ${names.map((n, i) => `${n} ${f(at(edits[i], 0.95))}`).join(' / ')} ms`)
  }
  for (const display of [false, true]) {
    const sheet = synth.book.sheets[0].name
    // the second session opened in a process is a little slower whatever its code: each build opens first once
    const sessions = [libs[0].createSession(synth.book, { clock, display }), libs[1].createSession(synth.book, { clock, display })]
    const steps = {
      insertRows: (s) => s.editStructure({ op: 'insertRows', sheet, row: 10 }),
      deleteRows: (s) => s.editStructure({ op: 'deleteRows', sheet, row: 10 }),
    }
    const ms = {}
    for (let round = 0; round < 8; round++)
      for (const [op, call] of Object.entries(steps))
        for (const i of pair(round + (op === 'deleteRows' ? 1 : 0)))
          for (const [what, g] of [[op, call], [`undo ${op}`, (s) => s.undo()], [`redo ${op}`, (s) => s.redo()], [`undo ${op} again`, (s) => s.undo()]]) {
            const t = performance.now()
            g(sessions[i])
            // the first round warms up
            if (round) ((ms[what] ??= [[], []])[i]).push(performance.now() - t)
          }
    sessions.forEach((s) => s.close())
    console.log(`${rows} rows, structure, display ${display}: ${Object.entries(ms).map(([what, xs]) => `${what} ${names.map((n, i) => `${n} ${f(at(xs[i], 0.5))}`).join(' / ')}`).join('; ')} ms`)
  }
}
