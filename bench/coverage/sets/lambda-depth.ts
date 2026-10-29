/**
 * How deep Excel lets LAMBDAs call one another, which the set `lambda` found past 2,000 and short of 5,000 calls through
 * LET: the same count-down at more depths, with a heavier calculation at each call (more functions around the call),
 * and with two LET names in it. Answers: `test/excel/cases/lambda-depth.json`; `PREDICTED`, the engine's answers
 * (`sets/predict.ts lambda-depth`), was committed before Excel ran the set.
 */
import { readFileSync } from 'node:fs'
import type { Val } from '../formulajs.js'
import type { Probe } from '../probes.js'
import { xl } from './arrays/common.js'
import { xlpm } from './let/xlpm.js'

/** A count-down from n through LET, each call of `step` around the next one. */
const deep = (n: number, step = (next: string) => `1+${next}`) => `LET(f,LAMBDA(g,n,IF(n<=0,0,${step('g(g,n-1)')})),f(f,${n}))`
const heavy = (next: string) => `ROUND(ABS(SUM(1,MAX(0,${next}))),0)`
const named = (next: string) => `LET(a,1,b,a+${next},b)`

const CASES: [string, string][] = [
  ...[2048, 2500, 3000, 3500, 4000, 4096, 4500].map((n): [string, string] => ['simple', deep(n)]),
  ...[1000, 2000, 3000, 4000].map((n): [string, string] => ['heavy', deep(n, heavy)]),
  ...[1000, 2000, 3000, 4000].map((n): [string, string] => ['named', deep(n, named)]),
]

export function buildProbes(): Probe[] {
  const count = new Map<string, number>()
  return CASES.map(([group, f]) => {
    const k = (count.get(group) ?? 0) + 1
    count.set(group, k)
    return { id: `${group}#${k}`, group, fn: 'LAMBDA', formula: xl(xlpm(f)) }
  })
}

/** The engine's answers by case id, committed before Excel ran the set. */
export const PREDICTED: Readonly<Record<string, Val>> = JSON.parse(readFileSync(new URL('./lambda-depth.predicted.json', import.meta.url), 'utf8'))
