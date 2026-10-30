/** Shared parts of the FORECAST.ETS sets: the sheet of series, a column pair each, and the series' makings. */

type Cell = number | string | boolean | { formula: string }

const letters = (i: number): string => (i < 26 ? String.fromCharCode(65 + i) : letters(Math.floor(i / 26) - 1) + letters(i % 26))

/**
 * A sheet E of series: `series(t, y)` puts a timeline and its values in two new columns from row 1 (null or undefined
 * leaves a cell blank) and gives their two ranges.
 */
export function seriesSheet(): { cells: Record<string, Cell>; series: (t: readonly (Cell | null)[], y: readonly (Cell | null)[]) => [string, string] } {
  const cells: Record<string, Cell> = {}
  let col = 0
  const series = (t: readonly (Cell | null)[], y: readonly (Cell | null)[]): [string, string] => {
    const [a, b] = [letters(col), letters(col + 1)]
    col += 2
    for (let i = 0; i < Math.max(t.length, y.length); i++) {
      if (t[i] !== null && t[i] !== undefined) cells[`${a}${i + 1}`] = t[i]!
      if (y[i] !== null && y[i] !== undefined) cells[`${b}${i + 1}`] = y[i]!
    }
    return [`E!${a}1:${a}${t.length}`, `E!${b}1:${b}${y.length}`]
  }
  return { cells, series }
}

export const steps = (n: number, from = 1, step = 1) => Array.from({ length: n }, (_, i) => from + i * step)

export const round2 = (v: number) => Math.round(v * 100) / 100

/** Noise from a seed: a linear congruential stream, centred, to two decimals. */
export function noise(seed: number, n: number): number[] {
  let x = seed
  return Array.from({ length: n }, () => {
    x = (x * 1103515245 + 12345) % 2147483648
    return round2((x / 2147483648 - 0.5) * 2)
  })
}

/** A date's serial number. */
export const day = (y: number, m: number, d: number) => Date.UTC(y, m, d) / 86400000 + 25569

/** n dates `every` months apart from a month of a year (1 to 12), on a day or a shorter month's last. */
export const dates = (year: number, month: number, d: number, every: number, n: number) =>
  steps(n, 0).map((i) => {
    const m = month - 1 + i * every
    return day(year, m, Math.min(d, new Date(Date.UTC(year, m + 1, 0)).getUTCDate()))
  })

export const ETS = 'FORECAST.ETS'
export const CI = 'FORECAST.ETS.CONFINT'
export const SEA = 'FORECAST.ETS.SEASONALITY'
export const STAT = 'FORECAST.ETS.STAT'
