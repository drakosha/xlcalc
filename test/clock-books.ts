/** Helpers of the tests of cells that depend on the clock (`dates-noyear`, `locale-clock`). */
import type { ComputeResult } from '../src/contract.js'

/** A clock in the year `year`, on 5 October at noon. */
export const at = (year: number, seed = 1) => ({ now: `${year}-10-05T12:00:00Z`, seed })

/** The values of the first sheet of a result, by address. */
export const values = (r: ComputeResult) => Object.fromEntries(Object.entries(r.sheets[0].cells).map(([a, c]) => [a, c.value]))

/** The addresses of a result's volatile cells, sorted. */
export const volatile = (r: ComputeResult) => r.volatile.map((c) => c.cell).sort()
