// node structure-child.mjs <workload> <engine>: rows inserted and deleted in the middle of a book's first sheet, each
// undone, in a fresh process by bench/hyperformula's protocol. Each operation and each undo is timed alone; after each,
// every populated cell is checked against a full calculation of the book as xlcalc's session leaves it.
import './iron.mjs'
import { workload } from '../hyperformula/workloads.mjs'
import { close, isHf, measure, median, open, p95, prepareInput, reference, verify } from '../hyperformula/shared.mjs'
import { parseAddress } from '../../dist/model.js'

const [id, engine] = process.argv.slice(2)
const { book } = workload(id)
const send = (message) => process.send?.(message)
const phase = (name) => send({ type: 'phase', name })
const keep = process.env.XLCALC_BENCH_KEEP === '1'
const rounds = id === 'business-34000' ? 3 : 5
const sheet = book.sheets[0].name
const middle = Math.floor(Math.max(...Object.keys(book.sheets[0].cells).map((a) => parseAddress(a).row + 1)) / 2)
const OPS = [
  { name: 'insert 1 row', op: 'insertRows', count: 1 },
  { name: 'delete 1 row', op: 'deleteRows', count: 1 },
  { name: 'insert 100 rows', op: 'insertRows', count: 100 },
]

/** An operation and its undo as the engine makes them. */
function calls(session, { op, count }) {
  if (isHf(engine)) {
    const s = session.getSheetId(sheet)
    return [() => (op === 'insertRows' ? session.addRows(s, [middle - 1, count]) : session.removeRows(s, [middle - 1, count])), () => session.undo()]
  }
  if (engine === 'ironcalc') return [() => session[op](0, middle, count), () => session.undo()]
  return [() => session.editStructure({ op, sheet, row: middle, count }), () => session.undo()]
}

let differing
function check(session, state) {
  phase('verify')
  const result = verify(engine, session, state.book, state.expected)
  if (!result.differences) return
  if (!keep) throw Object.assign(new Error('Values differ from a full calculation'), { details: result })
  if (!differing || result.differences > differing.differences) differing = result
}

const out = {}
let session
try {
  // the books each operation leaves, by xlcalc's session, and their full calculations
  phase('reference')
  const model = open('xl-closures', book)
  const states = new Map([['', { book, expected: reference(book) }]])
  for (const o of OPS) {
    model.editStructure({ op: o.op, sheet, row: middle, count: o.count })
    const after = model.getBook()
    states.set(o.name, { book: after, expected: reference(after) })
    model.undo()
  }
  close('xl-closures', model)
  phase('build')
  session = open(engine, book, prepareInput(engine, book))
  check(session, states.get(''))
  for (const o of OPS) {
    const [apply, undo] = calls(session, o)
    const done = [], undone = []
    for (let i = 0; i <= rounds; i++) {
      phase('operation')
      const a = measure(apply)
      check(session, states.get(o.name))
      phase('undo')
      const u = measure(undo)
      check(session, states.get(''))
      // the first round warms up
      if (i) done.push(a.wallMs), undone.push(u.wallMs)
    }
    out[o.name] = { medianMs: median(done), p95Ms: p95(done), samples: done }
    out[`undo ${o.name}`] = { medianMs: median(undone), p95Ms: p95(undone), samples: undone }
  }
  send({ type: 'result', result: { status: differing ? 'mismatch' : 'ok', middle, rounds, operations: out, ...(differing ? { details: differing } : {}) } })
} catch (error) {
  send({ type: 'result', result: { status: error.details ? 'mismatch' : 'error', operations: out, message: error.message, details: error.details } })
} finally {
  if (session) close(engine, session)
  process.disconnect?.()
}
