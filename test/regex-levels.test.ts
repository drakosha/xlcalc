/**
 * The three levels of match mode 3 (`docs/regex.md`): the built-in one (JavaScript's syntax with PCRE2's commonest
 * constructs rewritten), `pcreSyntax` and `pcreEngine` of `xlcalc/regex`, each against Excel's answers on the recorded
 * sets with its own known differences, registered in this order; a pool loads the registered level in its threads,
 * whether registered before the pool or after it; the engine's match limit is `#VALUE!` and `timeoutMs` stops a heavy
 * pattern in every kind of session, the calculator going on.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { compute, createInlineCalc, createSession, InputError, registerRegex, type BookInput, type CalcError } from '../src/index.js'
import { pcreEngine, pcreSyntax } from '../src/regex/index.js'
import { CLOCK, show, toVal, verdict } from './excel-cases.js'
import { caseBook, LEVELED_SETS, loadSet, STORED_SETS } from './excel-sets.js'
import { pool } from './fixtures/pool.js'

const rows = (n: number, text: string) => Object.fromEntries(Array.from({ length: n }, (_, i) => [`A${i + 1}`, { value: text }]))
/** XMATCH of `pattern` over `n` cells of `text`, in B1. */
const book = (pattern: string, n = 1, text = 'aab'): BookInput => ({ sheets: [{ name: 'S', cells: { ...rows(n, text), B1: { formula: `XMATCH("${pattern}",A1:A${n},3)` } } }] })
const b1 = (r: { sheets: { cells: Record<string, { value: unknown }> }[] }) => r.sheets[0].cells.B1.value
const VALUE = { error: 'VALUE' }
const NA = { error: 'NA' }
/** 40 letters a and b: Excel's match limit at the first start position. */
const forty = book('(a+)+$', 1, 'a'.repeat(40) + 'b')
/** 2,000 cells of 18 letters a and b, each under the limit: minutes of matching. */
const heavy = book('(a+)+$', 2000, 'a'.repeat(18) + 'b')
const fails = (p: Promise<unknown>) => p.then(() => null, (e: unknown) => e as CalcError)

const takes = (re: string) => {
  try {
    return new RegExp(re, 'v') && true
  } catch {
    return false
  }
}
/** What JavaScript takes from Node 23: duplicate group names (`pcreSyntax` needs them for a back reference to a name or number of several groups) and modifiers (`(?i:…)`). */
const DUPLICATE_NAMES = takes('(?<x>a)|(?<x>b)')
const MODIFIERS = takes('(?i:a)')

const ids = (bySet: Record<string, string>) => Object.entries(bySet).flatMap(([set, cases]) => cases.split(' ').map((c) => `${set}:${c}`)).sort()

/** The cases of each level that differ from Excel, by set. */
const KNOWN = {
  // JavaScript's syntax: PCRE2's constructs it lacks or reads otherwise are #VALUE! or another answer, no match limit;
  // duplicate group names and modifiers as the Node running the tests takes them
  builtin: ids({
    'xmatch-regex-syntax': `${DUPLICATE_NAMES ? 'group#10 ' : ''}${MODIFIERS ? '' : 'option#2 '}class#3 class#4 class#8 class#13 class#15 class#17 class#19 class#20 class#21 class#22 anchor#4 anchor#6 anchor#8 anchor#9 anchor#10 anchor#11 dot#2 dot#5 dot#6 dot#9 dot#11 quant#5 quant#6 quant#8 quant#10 quant#13 quant#14 group#3 group#4 group#5 group#6 group#7 group#8 group#11 group#14 look#6 option#1 option#3 option#7 option#8 option#9 option#10 option#11 option#12 option#13 option#14 option#15 option#16 option#17 recurse#1 recurse#2 recurse#3 recurse#4 recurse#5 cond#1 cond#2 cond#3 cond#4 cond#5 cond#6 verb#1 verb#2 verb#3 verb#4 verb#5 verb#6 verb#7 verb#8 bad#4 backtrack#2 backtrack#3 backtrack#4`,
    'xmatch-regex-syntax-2': 'line#2 line#3 line#4 line#5 line#8 line#9 line#10 line#11 line#12 class#6 class#7 class#9 class#10 class#11 class#12 class#13 option#1 option#2 option#4 option#5 option#6 recurse#1 recurse#2 recurse#3 bad#2 bad#4 backtrack#3 backtrack#5 backtrack#6 backtrack#8 backtrack#11 limit#3 limit#4 limit#5 limit#6',
    'xmatch-regex-own-held': 'class#1 look#2 verb#1 verb#2 verb#3 verb#4 group#1 group#2 group#3 group#4 group#6 group#7',
    'xmatch-regex-own-held-2': 'line#1 line#2 line#3',
  }),
  // recursion, conditions on groups, (*ACCEPT) and the verbs that cut backtracking, \X, Bidi_Class and (*NOTEMPTY)
  // have no translation (#VALUE!); no match limit (#N/A where Excel's is #VALUE!); duplicate group names before Node 23
  pcreSyntax: ids({
    'xmatch-regex-syntax': `${DUPLICATE_NAMES ? '' : 'group#8 '}dot#11 recurse#1 recurse#2 recurse#3 recurse#4 recurse#5 cond#1 cond#2 cond#3 cond#5 cond#6 verb#2 verb#3 verb#4 verb#5 verb#6 backtrack#2 backtrack#3 backtrack#4`,
    'xmatch-regex-syntax-2': 'class#13 option#3 recurse#1 recurse#2 recurse#3 bad#2 backtrack#3 backtrack#5 backtrack#6 backtrack#8 backtrack#11 limit#1 limit#2',
    // and a back reference to a group not set matches empty text in JavaScript
    'xmatch-regex-own-held': `${DUPLICATE_NAMES ? '' : 'group#3 '}class#1 verb#1 verb#2 verb#3 verb#4 group#1 group#2 group#6 group#7`,
  }),
  // JavaScript has no Bidi_Class property; the engine counts PCRE2's work approximately (a group without capture)
  pcreEngine: ['xmatch-regex-own-held-2:limit#6', 'xmatch-regex-syntax-2:class#13'],
}

/** The cases of `xmatch-regex-own-held` that JavaScript's backtracking takes hours over: the engine's alone. */
const HEAVY = /^xmatch-regex-own-held(-2)?:(limit|required|range)#/

/** The cases of the sets of match mode 3 that differ from Excel at the level registered now; `heavy`, those of HEAVY too. */
async function offExcel(heavy = false): Promise<string[]> {
  const off: string[] = []
  for (const set of STORED_SETS.filter((s) => s.name.startsWith('xmatch-regex'))) {
    const mod = await loadSet(set.name)
    const probes = mod.buildProbes().filter((p) => heavy || !HEAVY.test(`${set.name}:${p.id}`))
    const result = compute(caseBook(probes, mod.SHEETS), { clock: CLOCK }).sheets.at(-1)!.cells
    probes.forEach((p, i) => {
      const got = toVal(result[`A${i + 1}`].value)
      if (!verdict(`${set.name}:${p.id}`, p.fn, set.values[p.id], got).ok) off.push(`${set.name}:${p.id}`)
    })
  }
  return off.sort()
}

const early = await pool({ threads: 1 })
afterAll(() => early.close())

describe('the built-in level', () => {
  it("matches Excel but for JavaScript's syntax, Unicode \\w and \\b, \\A \\z \\Z, {,m} and a leading (?i) rewritten", async () => {
    expect(await offExcel()).toEqual(KNOWN.builtin)
    expect(b1(compute(book('^\\w+$', 1, 'привет')))).toBe(1)
    expect(b1(compute(book('а\\b', 1, 'мама')))).toBe(1)
    expect(b1(compute(book('\\Aa{,2}b\\z')))).toBe(1)
    expect(b1(compute(book('a++b')))).toEqual(VALUE)
    expect(b1(await early.compute(book('^\\w+$', 1, 'привет')))).toBe(1)
  })

  it('the sets of the levels are checked here alone', () => expect([...LEVELED_SETS].every((s) => STORED_SETS.some((x) => x.name === s))).toBe(true))

  it('registerRegex takes pcreSyntax or pcreEngine alone', () => {
    expect(() => registerRegex({} as never)).toThrow(new InputError('level', "expected pcreSyntax or pcreEngine of 'xlcalc/regex'"))
    expect(() => registerRegex({ kind: 'xlcalc-pcre-engine' } as never)).toThrow(InputError)
  })
})

describe('pcreSyntax', () => {
  it("matches Excel but for what has no translation, and has no match limit; a pool's threads load it", async () => {
    registerRegex(pcreSyntax)
    expect(await offExcel()).toEqual(KNOWN.pcreSyntax)
    expect(b1(compute(book('a(?i)AB')))).toBe(1)
    // past Excel's match limit (19 letters a): JavaScript backtracks to the end, #N/A
    expect(b1(compute(book('(a+)+$', 1, 'a'.repeat(19) + 'b')))).toEqual(NA)
    expect(b1(await early.compute(book('a(?i)AB')))).toBe(1)
  }, 30_000)
})

describe('pcreEngine', () => {
  it("matches Excel, in a pool's threads too, registered before the pool and after it", async () => {
    registerRegex(pcreEngine)
    registerRegex(pcreEngine)
    expect(await offExcel(true)).toEqual(KNOWN.pcreEngine)
    const late = await pool({ threads: 1 })
    try {
      for (const calc of [early, late]) expect(b1(await calc.compute(book('^(\\((?:[^()]|(?1))*\\))$', 1, '(a(b)c)')))).toBe(1)
    } finally {
      await late.close()
    }
  }, 30_000)

  it("Excel's match limit is #VALUE! at once, in every kind of session", async () => {
    const t0 = performance.now()
    expect(b1(compute(forty))).toEqual(VALUE)
    const s = createSession(forty)
    expect(b1(s.getResult())).toEqual(VALUE)
    s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 'aaa' } }])
    expect(b1(s.getResult())).toBe(1)
    expect(b1(await createInlineCalc().compute(forty))).toEqual(VALUE)
    expect(b1(await early.compute(forty))).toEqual(VALUE)
    expect(performance.now() - t0).toBeLessThan(5000)
  }, 20_000)

  it('timeoutMs stops a heavy pattern on the same thread, and the calculator and its session go on', async () => {
    const calc = createInlineCalc()
    const t0 = performance.now()
    expect((await fails(calc.compute(heavy, { timeoutMs: 200 })))?.code).toBe('soft_timeout')
    expect(performance.now() - t0).toBeLessThan(1500)
    // a stopped edit changes nothing: the texts stay short, and the next edit is calculated
    const s = await calc.createSession(book('(a+)+$', 2000, 'ab'))
    const long = Object.entries(rows(2000, 'a'.repeat(18) + 'b')).map(([cell, input]) => ({ sheet: 'S', cell, input }))
    expect((await fails(s.setCells(long, { timeoutMs: 200 })))?.code).toBe('soft_timeout')
    expect((await s.setCells([{ sheet: 'S', cell: 'A1', input: { value: 'aa' } }])).changes.find((c) => c.cell === 'B1')?.value).toBe(1)
    // a signal is seen between formulas (docs/async.md): one aborted before the call stops it at once
    expect((await fails(calc.compute(heavy, { signal: AbortSignal.abort() })))?.code).toBe('aborted')
    await s.close()
  }, 30_000)

  it("timeoutMs stops it in a pool's thread without hardTimeoutMs, and the thread goes on", async () => {
    const t0 = performance.now()
    expect((await fails(early.compute(heavy, { timeoutMs: 200 })))?.code).toBe('soft_timeout')
    expect(performance.now() - t0).toBeLessThan(2000)
    expect((await fails(early.compute(heavy, { signal: AbortSignal.abort() })))?.code).toBe('aborted')
    expect(b1(await early.compute(book('a*+b')))).toBe(1)
  }, 30_000)
})
