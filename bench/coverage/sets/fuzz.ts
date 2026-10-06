/**
 * The random check: formulas generated from the whitelist (`./fuzz/`), deterministic by seed, computed as
 * dynamic array formulas (Excel 365). A case's id is `<batch>#<n>`, its `fn` the function of the outer call.
 * Answers: `docs/excel-cases/fuzz.json`.
 */
import type { Probe } from '../probes.js'
import { xl } from './arrays/common.js'
import { fuzz, type Batch } from './fuzz/gen.js'

export const DYNAMIC = true

/** Batches: a seed, how many formulas, the most levels of nested calls. A new batch goes last, so stored answers stay. */
export const BATCHES: readonly Batch[] = [
  { seed: 20261006, count: 2000, depth: 3 },
  { seed: 20261007, count: 2000, depth: 4 },
]

export function buildProbes(): Probe[] {
  return BATCHES.flatMap((b, j) => fuzz(b).map((f, k): Probe => ({ id: `${j + 1}#${k + 1}`, group: f.fn, fn: f.fn, formula: xl(f.formula) })))
}
