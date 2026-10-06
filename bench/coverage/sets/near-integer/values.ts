/**
 * The values of the near-integer set on its sheet NI, column A, each computed by a formula (no literal of more
 * than 15 digits): `N-N*2^-j` and `N+N*2^-j` on a ladder of relative distances, then absolute offsets around
 * the boundaries the ladder found. A new group goes last, so the rows of the stored cases stay.
 */

export interface Value {
  /** The cell holding it, `NI!A<row>`. */
  at: string
  /** The whole number (or 0.3) it is near, with its sign. */
  n: number
  /** The relative distance 2^-j, when on the ladder. */
  j?: number
}

/** Relative distances 2^-j from the integer: from one unit in the last place to about 10^-6. */
const LADDER = [53, 52, 51, 50, 49, 48, 47, 46, 45, 44, 43, 42, 40, 37, 33, 27, 20]
const ABOVE = [52, 51, 50, 49, 48, 47, 46, 45, 44, 42, 40, 33, 20]
const SIGNED = [52, 50, 48, 47, 46, 45, 44, 42, 40, 33, 20]
const NEAR = [52, 51, 50, 49, 48, 47, 46, 45, 44, 42]

/** Absolute offsets below N around a tolerance of about 10^-7 relative (whole-number arguments). */
const FINE: [number, string[]][] = [
  [1, ['2E-08', '2.5E-08', '2.9E-08', '3E-08', '3.1E-08', '4E-08', '4.9E-08', '5.1E-08', '6E-08', '1E-07', '2E-07', '5E-07']],
  [2, ['5E-08', '5.9E-08', '6E-08', '1E-07', '1.2E-07', '3E-07', '4.9E-07', '5.1E-07', '1E-06']],
  [3, ['1E-07', '1.15E-07', '1.25E-07', '2E-07', '4.9E-07', '5.1E-07', '1E-06']],
  [10, ['4E-07', '4.7E-07', '4.85E-07', '4.95E-07', '5.05E-07', '6E-07', '1E-06']],
  [100, ['3.7E-06', '3.9E-06', '4.9E-06', '5.1E-06', '1E-05']],
  [1000000, ['0.03', '0.0312', '0.0313', '0.04', '0.049', '0.051', '0.06']],
]

/** Absolute offsets below N for ISEVEN and ISODD (a tolerance between 10^-11 and 10^-10 on the ladder). */
const PARITY: [number, string[]][] = [
  [1, ['2E-11', '5E-11', '8E-11', '1E-10', '2E-10']],
  [10, ['5E-11', '8E-11', '1E-10', '2E-10', '5E-10', '1E-09']],
  [1000, ['1E-10', '1E-09', '1E-08', '2E-08', '1E-07']],
]

/** Seconds below midnight of a date serial (the date functions read a moment to the second). */
const SECONDS = ['0.4', '0.49', '0.499', '0.501', '0.51', '0.6', '1']

export function values() {
  const sheet: Record<string, { formula: string }> = {}
  const put = (n: number, formula: string, j?: number): Value => {
    const at = `NI!A${Object.keys(sheet).length + 1}`
    sheet[at.slice(3)] = { formula }
    return { at, n, j }
  }
  const lit = (n: number) => (n === 1e15 ? '1E+15' : String(n))
  const below = (n: number, js = LADDER) => js.map((j) => put(n, `${lit(n)}-${lit(n)}*2^-${j}`, j))
  const offsets = (table: [number, string[]][]) => table.flatMap(([n, ds]) => ds.map((d) => put(n, `${n}-${d}`)))
  const g = {
    small: [1, 2, 3, 10].flatMap((n) => below(n)),
    big: [1e6, 1e15].flatMap((n) => below(n)),
    date: below(45688),
    above: [1, 2, 3, 10].flatMap((n) => ABOVE.map((j) => put(n, `${n}+${n}*2^-${j}`, j))),
    // away from zero (INT) and toward it (TRUNC)
    away: [1, 2, 1e6].flatMap((n) => SIGNED.map((j) => put(-n, `-${n}-${n}*2^-${j}`, j))),
    toward: [1, 2, 1e6].flatMap((n) => SIGNED.map((j) => put(-n, `-${n}+${n}*2^-${j}`, j))),
    point3: [...NEAR.map((j) => put(0.3, `0.3-0.3*2^-${j}`, j)), ...NEAR.map((j) => put(0.3, `0.3+0.3*2^-${j}`, j))],
    fine: offsets(FINE),
    // 2^-40 and 2^-30 below 1, 2, 3, 4 and 9: past the 15 digits, inside a tolerance of 10^-7
    sample: [1, 2, 3, 4, 9].flatMap((n) => below(n, [40, 30])),
    parity: offsets(PARITY),
    seconds: [45688, 45658].flatMap((n) => SECONDS.map((s) => put(n, `${n}-${s}/86400`))),
  }
  return { sheet, g }
}
