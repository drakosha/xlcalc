/**
 * The groups of rarely used functions (`docs/function-groups.md`): the package root calculates all of them as before,
 * the light entry `xlcalc/core` those a host registers, a formula over a group not loaded is `#NAME?` naming its group,
 * and the checks of formula text know every function. Each case runs in a process of its own (`fixtures/groups-child.ts`),
 * without the test setup that loads every group.
 */
import { execFileSync } from 'node:child_process'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { externalNames, SYNTAX } from '../src/arity.js'
import { FUNCTION_NAMES } from '../src/declared-arity.js'
import { FUNCTION_WORDS } from '../src/function-words.js'
import { FUNCTIONS } from '../src/functions.js'
import { allFunctions } from '../src/groups/index.js'
import { GROUP_ARITY } from '../src/groups/manifest.js'

const child = fileURLToPath(new URL('./fixtures/groups-child.ts', import.meta.url))
const run = (name: string): any => JSON.parse(execFileSync(process.execPath, ['--import', 'tsx', child, name], { encoding: 'utf8' }))

/** The values of the child's book with every group (`fixtures/groups-child.ts`, `BOOK`). */
const ALL = { A1: 2, B1: 0.576724807896294, C1: 1.15344961579259, D1: 0.977249868051821, E1: 450, F1: 7, G1: 2, H1: 1.68540158589943, I1: 3 }
const NAME = { error: 'NAME' }
const notLoaded = (cell: string, fn: string, group: string) => ({ cell, check: 'function_not_loaded', fn, reason: `${fn} is in the function group ${group}, which is not loaded (registerFunctions)` })

/** What the checks of text answer of BESSELJ whichever groups are loaded. */
const KNOWN = {
  validate: { ok: true, canonical: 'BESSELJ(1,2)', functions: ['BESSELJ'], sheets: [], names: [], externals: [], tables: [] },
  count: { ok: false, code: 'syntax', message: 'BESSELJ takes 2 arguments, got 1', at: 0 },
  entry: { ok: true, input: { formula: 'BESSELJ(A1,1)' } },
  description: 'The Bessel function ',
  names: FUNCTION_NAMES,
  external: ['MYFN', 'GETPIVOTDATA'],
}

describe('the groups of functions', () => {
  it("the whitelist is 0.68.0's whatever is loaded: the functions the engine calculates and LAMBDA, LET", () => {
    expect(FUNCTION_NAMES).toEqual([...FUNCTION_WORDS, ...SYNTAX].sort())
    expect(FUNCTION_NAMES).toEqual([...Object.keys(FUNCTIONS), ...SYNTAX].sort())
    expect(FUNCTION_NAMES.length).toBe(517)
    expect(externalNames(['besselj', 'MyFn', 'GETPIVOTDATA', 'SUM'])).toEqual(KNOWN.external)
  })

  it("each group's functions are its manifest's, none in two groups", () => {
    expect(allFunctions.map((g) => [g.name, g.functions])).toEqual(Object.entries(GROUP_ARITY).map(([g, defs]) => [g, Object.keys(defs).sort()]))
    const all = allFunctions.flatMap((g) => g.functions)
    expect(new Set(all).size).toBe(all.length)
    expect(allFunctions.map((g) => [g.name, g.functions.length])).toEqual([['engineering', 54], ['finance', 33], ['lambda', 8], ['statistics', 77], ['xml', 1]])
  })

  it('the root knows every function before it calculates, and calculates all of them in both kinds of execution', () => {
    const r = run('root')
    expect(r.before).toEqual(KNOWN)
    expect(r.values).toEqual({ closures: ALL, eval: ALL })
    // the hash 0.68.0 gives the book, before the groups
    expect(r.hash).toEqual({ closures: 'eac2febc78ff563e', eval: 'eac2febc78ff563e' })
  }, 60_000)

  it('the light entry: #NAME? naming the group, a name over one, the value its file stores when held; the checks of text know them; registered, they calculate', () => {
    const r = run('core')
    const issues = [notLoaded('B1', 'BESSELJ', 'engineering'), { cell: 'C1', reason: "unknown function or name, from 'S'!B1" }, notLoaded('D1', 'NORM.S.DIST', 'statistics'), notLoaded('E1', 'SLN', 'finance'), notLoaded('F1', 'FILTERXML', 'xml'), notLoaded('G1', 'REDUCE', 'lambda'), { cell: 'H1', reason: 'unknown function or name' }]
    const before = { values: { A1: 2, B1: NAME, C1: NAME, D1: NAME, E1: NAME, F1: NAME, G1: NAME, H1: NAME, I1: 3 }, issues, names: [{ kind: 'name', name: 'Rate', reason: 'ERF is in the function group engineering, which is not loaded (registerFunctions)', dropped: false }] }
    expect(r.before).toEqual({ closures: before, eval: before })
    expect(r.held).toEqual({ ...notLoaded('B1', 'BESSELJ', 'engineering'), reason: notLoaded('B1', 'BESSELJ', 'engineering').reason + ': not calculated, the value the file stores is held', fromFile: true })
    expect(r.heldValue).toBe(1)
    expect(r.refused).toEqual(Array(3).fill("groups[0]: expected a group of 'xlcalc/functions' (engineering, finance, lambda, statistics, xml)"))
    expect(r.text).toEqual(KNOWN)
    expect(r.some).toEqual({ ...ALL, E1: NAME, F1: NAME, G1: NAME })
    expect(r.all).toEqual({ closures: ALL, eval: ALL })
    expect(r.hash).toEqual({ closures: 'eac2febc78ff563e', eval: 'eac2febc78ff563e' })
  }, 60_000)

  it('a host with both entries has one table: once the root calculates, the light entry has every function', () => {
    expect(run('both')).toEqual({ first: { A1: 2, B1: NAME, C1: NAME, D1: NAME, E1: NAME, F1: NAME, G1: NAME, H1: NAME, I1: 3 }, after: ALL, names: true })
  }, 60_000)

  it('a session opened before a group is loaded calculates its formulas at its next call; its undo and redo work as before', () => {
    const r = run('sessions')
    const at = (a: number, b: number, c: number) => [{ sheet: 'S', cell: 'A1', value: a, text: String(a) }, { sheet: 'S', cell: 'I1', value: a + 1, text: String(a + 1) }, ...[b, c].map((v, i) => ({ sheet: 'S', cell: i ? 'C1' : 'B1', value: v, text: String(Number(v.toPrecision(6))) }))]
    const three = [3, 0.339058958272536, 0.678117916545072, NAME]
    const session = { after: three, undo: at(2, ALL.B1, ALL.C1), undone: [2, ALL.B1, ALL.C1, NAME], redo: at(3, three[1] as number, three[2] as number), redone: three, issue: 'function_not_loaded', revision: 3 }
    // the state built again tells no change and takes no revision: the host reads what it shows again
    expect(r).toEqual({ revision: 1, sessions: [session, session], inline: three })
  }, 60_000)

  it("a pool's threads load the groups registered, before or after the pool; the root's pool has every group", () => {
    expect(run('pool')).toEqual({ before: [2, NAME, NAME, NAME, NAME, NAME], after: [2, ALL.B1, ALL.C1, NAME, 450, NAME], fresh: 450 })
    expect(run('rootPool')).toEqual(ALL)
  }, 60_000)
})
