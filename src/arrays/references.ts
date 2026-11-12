/**
 * Arrays of references. OFFSET and INDIRECT lifted over an array (`OFFSET(B1,{0;1},0)`) answer, as any lifted call,
 * the top-left values of the references they find; the references themselves are kept here, beside that array. A
 * function that takes a reference (SUBTOTAL, COUNTBLANK, SUMIF and the other criteria functions:
 * `FunctionSpec.referencesAt`) given such an array is called once for each reference and answers an array of the same
 * shape, as Excel does (the spill sets `ref-arrays-held`, `cse-lookup-held-5`, the workbook set `ref-arrays-hidden-held`).
 */
import { liftN } from './lift.js'
import { Arr } from './value.js'
import type { V } from '../values.js'

/** The call's answers by position of the array's occupied part, row by row, then the one for its fill. */
const KEPT = new WeakMap<Arr, readonly unknown[]>()

/** `a`, the lifted call's array of values, with the call's answers `found` (references, errors) kept beside it. */
export function keepReferences(a: Arr, found: readonly unknown[]): Arr {
  KEPT.set(a, found)
  return a
}

/** Whether `x` is an array a lifted OFFSET or INDIRECT answered, with its references. */
export const holdsReferences = (x: unknown): x is Arr => x instanceof Arr && KEPT.has(x)

/**
 * `call` once for each position of the arrays of references at `at` among `args`, each given its reference there:
 * the arrays lift together as any arrays do (one row or column repeated, #N/A past a shorter one).
 */
export function overReferences(args: readonly unknown[], at: readonly number[], call: (args: unknown[]) => V): Arr {
  const found = at.map((i) => KEPT.get(args[i] as Arr)!)
  // each array as the numbers of its places, its fill the number after them
  const places = at.map((i) => {
    const a = args[i] as Arr
    const n = a.h * a.w
    return new Arr(a.rows, a.cols, a.h, a.w, Array.from({ length: n }, (_, k) => k), n)
  })
  const b = args.slice()
  return liftN(places, (els) => {
    for (let k = 0; k < at.length; k++) {
      const e = els[k]
      b[at[k]] = typeof e === 'number' ? found[k][e] : e
    }
    return call(b)
  }, true)
}
