/**
 * Held-out array formulas of the old kind around lookups (`run.ts spill cse-lookup-held`): other aggregations over a
 * positional XLOOKUP or INDEX, and other functions that select by an array of positions. The answers were predicted
 * from the rules of the set `cse-lookup` and committed before Excel computed them (`PREDICTED`).
 */
import { one, type CellSpec, type SpillCase } from './cases.js'
import { arr, DATA } from './cse-lookup.js'
import { PREDICTED } from './cse-lookup-held.predicted.js'

const held = (id: string, formula: string, area: string): SpillCase => {
  const at = area.split(':')[0]
  return one(id, `${formula} over ${area}`, { ...DATA, [at]: arr(formula, area) } as Record<string, CellSpec>, [area], { predicted: PREDICTED[id] })
}

/** The cases whose prediction Excel answered otherwise. */
export const MISSED: readonly string[] = ['sumproduct-two']

const KEYS = 'XLOOKUP({1;2},A1:A3,B1:B3)'

export function buildSpillCases(): SpillCase[] {
  return [
    held('average-two', `AVERAGE(${KEYS})`, 'G1:G2'),
    held('average-single', `AVERAGE(${KEYS})`, 'G1'),
    held('max-two', `MAX(${KEYS})`, 'G1:G2'),
    held('min-two', `MIN(${KEYS})`, 'G1:G2'),
    held('count-two', `COUNT(${KEYS})`, 'G1:G2'),
    held('count-single', `COUNT(${KEYS})`, 'G1'),
    held('product-two', `PRODUCT(${KEYS})`, 'G1:G2'),
    held('sumproduct-two', `SUMPRODUCT(${KEYS})`, 'G1:G2'),
    held('average-index-two', 'AVERAGE(INDEX(B1:B3,{2;3}))', 'G1:G2'),
    held('index-xmatch-two', 'SUM(INDEX(B1:B3,XMATCH({1;2},A1:A3)))', 'G1:G2'),
    held('max-plus-one-two', `MAX(${KEYS}+1)`, 'G1:G2'),
    held('offset-two', 'SUM(OFFSET(B1,{1;2},0))', 'G1:G2'),
    held('offset-single', 'SUM(OFFSET(B1,{1;2},0))', 'G1'),
    held('choose-two', 'SUM(CHOOSE({1;2},B2,B3))', 'G1:G2'),
    held('lookup-two', 'SUM(LOOKUP({1;2},A2:A3,B2:B3))', 'G1:G2'),
    held('union-key-two', 'SUM(XLOOKUP((D1,D2),A1:A3,B1:B3))', 'G1:G2'),
  ]
}
