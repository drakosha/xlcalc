/**
 * Child process for the tests of the groups of functions (`test/function-groups.test.ts`): each case in a process of its
 * own, as a host loads the entries, without the test setup that loads every group. Prints the case's findings as JSON.
 */
import type { BookInput, Calc, CellValue, ComputeResult } from '../../src/contract.js'
import { libUrl } from './target.js'
import { threadScript } from './thread-script.js'

type Root = typeof import('../../src/index.js')
type Core = typeof import('../../src/core.js')
type Groups = typeof import('../../src/groups/index.js')

const lib = <T>(file: string) => import(libUrl(file)) as Promise<T>
const CLOCK = { now: '2026-10-04T12:00:00Z', seed: 7 }

/** A book over one function of each group, a cell that reads one, and a name over a function of a group. */
export const BOOK: BookInput = {
  sheets: [
    {
      name: 'S',
      cells: {
        A1: { value: 2 },
        B1: { formula: 'BESSELJ(A1,1)' },
        C1: { formula: 'B1*2' },
        D1: { formula: 'NORM.S.DIST(A1,TRUE)' },
        E1: { formula: 'SLN(1000,100,A1)' },
        F1: { formula: 'FILTERXML("<a><b>7</b></a>","//b")' },
        G1: { formula: 'REDUCE(0,A1:A1,LAMBDA(x,y,x+y))' },
        H1: { formula: 'Rate*2' },
        I1: { formula: 'SUM(A1,1)' },
      },
    },
  ],
  names: [{ name: 'Rate', ref: 'ERF(1)' }],
}

const values = (r: ComputeResult): Record<string, CellValue> => Object.fromEntries(Object.entries(r.sheets[0].cells).map(([a, c]) => [a, c.value]))
const issues = (r: ComputeResult) => r.diagnostics.flatMap((d) => (d.kind === 'cell' ? [{ cell: d.cell, check: d.check, fn: d.fn, reason: d.reason, fromFile: d.fromFile }] : []))
const both = (f: (execution: 'closures' | 'eval') => unknown) => ({ closures: f('closures'), eval: f('eval') })

/** What the checks of formula text say of a function of a group, the engine loaded or not. */
async function checks(entry: Root | Core) {
  const text = await lib<typeof import('../../src/function-text/index.js')>('function-text/index.js')
  return {
    validate: entry.validateFormula('BESSELJ(1,2)'),
    count: entry.validateFormula('BESSELJ(1)'),
    entry: entry.parseEntry('=besselj(A1,1)'),
    description: text.functionDescription('BESSELJ')?.slice(0, 20),
    names: entry.FUNCTION_NAMES,
    external: (await lib<typeof import('../../src/arity.js')>('arity.js')).externalNames(['besselj', 'MyFn', 'GETPIVOTDATA', 'SUM']),
  }
}

/** A pool of the thread under test, as `createCalcPool` makes it in Node. */
async function pool(): Promise<Calc> {
  const { checkPoolOptions } = await lib<typeof import('../../src/async/options.js')>('async/options.js')
  const { nodeSpawn, workerThreads } = await lib<typeof import('../../src/worker/platform.js')>('worker/platform.js')
  const { createPool } = await lib<typeof import('../../src/worker/pool.js')>('worker/pool.js')
  return createPool(checkPoolOptions({ threads: 1 }), nodeSpawn(workerThreads()!, await threadScript(), undefined))
}

const cases: Record<string, () => Promise<unknown>> = {
  /** The root before any calculation knows every function; its first calculation loads them all. */
  async root() {
    const root = await lib<Root>('index.js')
    const before = await checks(root)
    return { before, values: both((execution) => values(root.compute(BOOK, { clock: CLOCK, execution }))), hash: both((execution) => root.compute(BOOK, { clock: CLOCK, execution }).hash.slice(0, 16)) }
  },
  /** The light entry: refusals, kept values, the checks of text, a name's formula, then the groups loaded. */
  async core() {
    const core = await lib<Core>('core.js')
    const groups = await lib<Groups>('groups/index.js')
    const before = both((execution) => {
      const r = core.compute(BOOK, { clock: CLOCK, execution })
      return { values: values(r), issues: issues(r), names: r.diagnostics.filter((d) => d.kind === 'name') }
    })
    const held = issues(core.compute(BOOK, { clock: CLOCK, unknownFunctions: 'held', cached: { sheets: [{ name: 'S', cells: { B1: { value: 0.5 } } }] } })).find((i) => i.cell === 'B1')
    const heldValue = core.compute(BOOK, { clock: CLOCK, unknownFunctions: 'held', cached: { sheets: [{ name: 'S', cells: { B1: { value: 0.5 } } }] } }).sheets[0].cells.C1.value
    const refused = [{}, 'engineering', { kind: 'xlcalc-functions', name: 'engineering' }].map((g) => {
      try {
        core.registerFunctions(g as never)
        return 'taken'
      } catch (e) {
        return (e as Error).message
      }
    })
    const text = await checks(core)
    core.registerFunctions(groups.engineering, groups.statistics)
    const some = values(core.compute(BOOK, { clock: CLOCK }))
    core.registerFunctions(...groups.allFunctions)
    return { before, held, heldValue, refused, text, some, all: both((execution) => values(core.compute(BOOK, { clock: CLOCK, execution }))), hash: both((execution) => core.compute(BOOK, { clock: CLOCK, execution }).hash.slice(0, 16)), groups: groups.allFunctions.map((g) => [g.name, g.functions.length]) }
  },
  /** A host with both entries has one table: the light entry calculates every function once the root has calculated. */
  async both() {
    const core = await lib<Core>('core.js')
    const root = await lib<Root>('index.js')
    const first = values(core.compute(BOOK, { clock: CLOCK }))
    root.compute({ sheets: [{ name: 'S', cells: {} }] }, { clock: CLOCK })
    return { first, after: values(core.compute(BOOK, { clock: CLOCK })), names: core.FUNCTION_NAMES.length === root.FUNCTION_NAMES.length }
  },
  /** Sessions opened before a group is loaded, with a history of edits, in both kinds of execution and an inline calculator. */
  async sessions() {
    const core = await lib<Core>('core.js')
    const groups = await lib<Groups>('groups/index.js')
    const made = (['closures', 'eval'] as const).map((execution) => core.createSession(BOOK, { clock: CLOCK, execution }))
    const inline = core.createInlineCalc()
    const s3 = await inline.createSession(BOOK, { clock: CLOCK })
    for (const s of made) s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
    await s3.setCells([{ sheet: 'S', cell: 'A1', input: { value: 3 } }])
    const revision = made[0].revision
    core.registerFunctions(groups.engineering)
    const read = (s: { getValues(sheet: string, range: string): CellValue[][] }) => s.getValues('S', 'A1:D1')[0]
    const out = made.map((s) => {
      const after = read(s)
      const undo = s.undo()?.changes
      const undone = read(s)
      const redo = s.redo()?.changes
      return { after, undo, undone, redo, redone: read(s), issue: s.getCellIssue('S', 'D1')?.check, revision: s.revision }
    })
    const inlineAfter = (await s3.getValues('S', 'A1:D1'))[0]
    await inline.close()
    return { revision, sessions: out, inline: inlineAfter }
  },
  /** A pool's thread loads the groups the host registered once it calculates with the light entry, those registered later too; the root's pool has them all. */
  async pool() {
    const core = await lib<Core>('core.js')
    const groups = await lib<Groups>('groups/index.js')
    core.compute({ sheets: [{ name: 'S', cells: {} }] })
    const calc = await pool()
    const s = await calc.createSession(BOOK, { clock: CLOCK })
    const before = (await s.getValues('S', 'A1:F1'))[0]
    core.registerFunctions(groups.engineering, groups.finance)
    const after = (await s.getValues('S', 'A1:F1'))[0]
    const fresh = (await calc.compute(BOOK, { clock: CLOCK })).sheets[0].cells.E1.value
    await calc.close()
    return { before, after, fresh }
  },
  async rootPool() {
    await lib<Root>('index.js')
    const calc = await pool()
    const r = await calc.compute(BOOK, { clock: CLOCK })
    await calc.close()
    return values(r)
  },
}

console.log(JSON.stringify(await cases[process.argv[2]]()))
