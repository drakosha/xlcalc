/**
 * Order statistics of a list of numbers as a stable ascending sort gives them, without sorting where one or two
 * positions are wanted: a quickselect over a Float64Array copy. Typed arrays order -0 before 0 and a quickselect
 * does not keep input order, so a zero answered is the zero the stable sort puts there, by its rank among the zeros.
 */

/** Quickselect: the value of sorted position `k` of `a` moved there, smaller ones before it, larger ones after. */
function select(a: Float64Array, k: number): void {
  let lo = 0
  let hi = a.length - 1
  // a full sort of what is left past the depth limit keeps the worst case at n log n
  let depth = 2 * Math.ceil(Math.log2(a.length + 1)) + 8
  while (hi > lo) {
    if (hi - lo < 16 || --depth < 0) {
      a.subarray(lo, hi + 1).sort()
      return
    }
    const x = a[lo], y = a[(lo + hi) >>> 1], z = a[hi]
    const p = x < y ? (y < z ? y : x < z ? z : x) : x < z ? x : y < z ? z : y
    let i = lo
    let j = hi
    while (i <= j) {
      while (a[i] < p) i++
      while (a[j] > p) j--
      if (i <= j) {
        const t = a[i]
        a[i++] = a[j]
        a[j--] = t
      }
    }
    if (k <= j) hi = j
    else if (k >= i) lo = i
    else return
  }
}

/** The zero a stable sort of `xs` puts at sorted position `k`, given that a zero is there: the zeros keep their order. */
function zeroAt(xs: readonly number[], k: number): number {
  let below = 0
  for (const x of xs) if (x < 0) below++
  for (const x of xs) if (x === 0 && below++ === k) return x
  return 0
}

/**
 * Whether the stable sort `(a, b) => a - b` is the way: a short list, which it sorts faster than a typed array is
 * made, or one with an infinity, where its order and a typed array's differ more than in zeros.
 */
const plain = (xs: readonly number[]) => xs.length < 16 || !xs.every(Number.isFinite)
const stable = (xs: readonly number[]) => [...xs].sort((a, b) => a - b)

/** The value at sorted position `k` (0-based), and the one after it when `next` (the same past the last). */
export function orderAt(xs: readonly number[], k: number, next = false): [number, number] {
  if (plain(xs)) {
    const s = stable(xs)
    return [s[k], s[next ? Math.min(k + 1, s.length - 1) : k]]
  }
  const a = Float64Array.from(xs)
  select(a, k)
  const v = a[k] === 0 ? zeroAt(xs, k) : a[k]
  if (!next || k + 1 >= a.length) return [v, v]
  // the next one is the smallest after k
  let w = a[k + 1]
  for (let i = k + 2; i < a.length; i++) if (a[i] < w) w = a[i]
  return [v, w === 0 ? zeroAt(xs, k + 1) : w]
}

/** The whole list sorted as `[...xs].sort((a, b) => a - b)` sorts it, zeros in their input order. */
export function sortedNumbers(xs: readonly number[]): number[] {
  if (plain(xs)) return stable(xs)
  const a = Float64Array.from(xs).sort()
  let below = 0
  while (below < a.length && a[below] < 0) below++
  for (const x of xs) if (x === 0) a[below++] = x
  return Array.from(a)
}
