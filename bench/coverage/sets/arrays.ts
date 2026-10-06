/**
 * Excel cases for arrays inside a cell (`docs/arrays-incell.md`), computed as dynamic array formulas
 * (Excel 365): operators over ranges and arrays with Excel's size rules, array constants, the result of
 * a formula that is an array (`#SPILL!`, every neighbour of a case is taken), aggregates and lookups over
 * arrays, lifting of every whitelisted function by argument, and the functions with an array answer.
 * Answers: `docs/excel-cases/arrays.json`.
 */
import type { Probe } from '../probes.js'
import { collector } from './arrays/common.js'
import { aggregates, keyed, lazy, lifted, reference } from './arrays/consumers.js'
import { arrayFunctions, SHEETS } from './arrays/functions.js'
import { lift } from './arrays/lift.js'
import { constants, operators, results } from './arrays/ops.js'

export { SHEETS }

export const DYNAMIC = true

export function buildProbes(): Probe[] {
  const { out, add } = collector()
  for (const part of [operators, constants, results, aggregates, lazy, keyed, lifted, reference, arrayFunctions, lift]) part(add)
  return out
}
