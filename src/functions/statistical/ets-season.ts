/**
 * The season FORECAST.ETS and its kin detect where seasonality is 1 (their default) and FORECAST.ETS.SEASONALITY
 * reports, as Excel detects it, measured: the series less its line and normalized, a quick answer from its periodogram,
 * candidates from the periodogram's peaks checked for a hill in the autocorrelation, ranked by power and correlation,
 * and a choice between the two. `round` is floor(x + 0.5) throughout; comparisons are strict and the first of equals wins.
 */
import { line, places } from './ets-model.js'

/** The longest season FORECAST.ETS takes. */
export const MAX_SEASON = 8784
/** The points detection reads at most, from the end. */
const MAX_POINTS = 175680

const round = (x: number) => Math.floor(x + 0.5)

/** The residuals of the series about its line on 1…n, and them less their mean over a deviation (them, for none). */
function preprocessed(y: readonly number[]): [number[], number[]] {
  const n = y.length
  const [a, b] = line(places(1, n), y)
  const r = y.map((v, i) => v - (a * (i + 1) + b))
  let s = 0
  let q = 0
  for (const v of r) {
    s += v
    q += v * v
  }
  const mean = s / n
  const sd = Math.sqrt(Math.max(0, n > 1 ? q / (n - 1) - mean * mean : 0))
  return [r, sd ? r.map((v) => (v - mean) / sd) : [...r]]
}

/** The discrete Fourier transform of real values at bins 0…⌊n/2⌋: a radix-2 FFT for a power of two, else by a table of roots. */
function dft(v: readonly number[], bins: number): [Float64Array, Float64Array] {
  const n = v.length
  if ((n & (n - 1)) === 0) {
    const [re, im] = fft(v, n)
    return [re.subarray(0, bins), im.subarray(0, bins)]
  }
  const cos = new Float64Array(n)
  const sin = new Float64Array(n)
  for (let m = 0; m < n; m++) {
    cos[m] = Math.cos((2 * Math.PI * m) / n)
    sin[m] = -Math.sin((2 * Math.PI * m) / n)
  }
  const re = new Float64Array(bins)
  const im = new Float64Array(bins)
  for (let k = 0; k < bins; k++) {
    let a = 0
    let b = 0
    for (let j = 0, m = 0; j < n; j++, m = (m + k) % n) {
      a += v[j] * cos[m]
      b += v[j] * sin[m]
    }
    re[k] = a
    im[k] = b
  }
  return [re, im]
}

/** An iterative radix-2 FFT of real values padded with zeros to `size`, a power of two. */
function fft(v: readonly number[], size: number): [Float64Array, Float64Array] {
  const re = new Float64Array(size)
  const im = new Float64Array(size)
  for (let i = 0, j = 0; i < size; i++) {
    if (i < v.length) re[j] = v[i]
    for (let bit = size >> 1; bit && (j ^= bit) < bit; bit >>= 1);
  }
  for (let len = 2; len <= size; len <<= 1) {
    const ang = (-2 * Math.PI) / len
    for (let i = 0; i < size; i += len)
      for (let k = 0; k < len >> 1; k++) {
        const wr = Math.cos(ang * k)
        const wi = Math.sin(ang * k)
        const a = i + k
        const b = a + (len >> 1)
        const xr = re[b] * wr - im[b] * wi
        const xi = re[b] * wi + im[b] * wr
        re[b] = re[a] - xr
        im[b] = im[a] - xi
        re[a] += xr
        im[a] += xi
      }
  }
  return [re, im]
}

/** The periodogram of values less their line, tapered at each tenth end: powers of bins 1…⌊n/2⌋, their mean and deviation. */
interface Spectrum {
  power: number[]
  mean: number
  sd: number
}

function spectrum(y: readonly number[]): Spectrum {
  const n = y.length
  const half = Math.ceil(n / 2)
  let s = 0
  let sxy = 0
  for (let i = 0; i < n; i++) {
    s += y[i]
    sxy += (i + 1 - half) * y[i]
  }
  const mean = s / n
  const den = (2 * n * (n * n - 1)) / 24
  const slope = den ? sxy / den : 0
  const v = y.map((x, i) => x - (slope * (i + 1 - half) + mean))
  const edge = Math.trunc(0.1 * n)
  for (let j = 0; j < edge; j++) {
    const w = 0.5 * (1 - Math.cos((Math.PI * (2 * j + 1)) / (2 * edge)))
    v[j] *= w
    v[n - 1 - j] *= w
  }
  const bins = (n >> 1) + 1
  const [re, im] = dft(v, bins)
  const power: number[] = []
  for (let k = 1; k < bins; k++) power.push((re[k] * re[k] + im[k] * im[k]) / (0.875 * n))
  const j = power.length
  let ps = 0
  let pq = 0
  for (const p of power) {
    ps += p
    pq += p * p
  }
  const mu = j ? ps / j : 0
  return { power, mean: mu, sd: j > 1 ? Math.sqrt(Math.max(0, pq / (j - 1) - mu * mu)) : 0 }
}

/** The quick answer: the period of the periodogram's highest peak if it stands out enough, else 1; and that peak's power. */
function quick(n: number, { power, mean, sd }: Spectrum): [number, number] {
  if (power.length < 2) return [1, 0]
  let best = power[0]
  let second = power[1]
  let period = n
  if (second > best) {
    ;[best, second] = [second, best]
    period = round(n / 2 - 3e-7)
  }
  for (let k = 2; k < power.length; k++) {
    const p = power[k]
    if (p > best) {
      ;[second, best] = [best, p]
      period = round(n / (k + 1) - 3e-7)
    } else if (p > second) second = p
  }
  const valid =
    period >= 2 && period <= Math.min(5, Math.trunc((2 * n) / 5)) ? (second ? best / second : 0) > 2 - 3 / n && best > mean : best > mean + 2.5 * sd && n >= 2.5 * period
  return [valid ? period : 1, best]
}

/** The periods of the three strongest bins above the mean power, and of the bins that stand alone far above their neighbours. */
function candidates(n: number, { power, mean }: Spectrum): number[] {
  const ordinary: [number, number][] = []
  const strong: number[] = []
  power.forEach((value, k) => {
    const p = round(n / (k + 1))
    if (!(value > mean && p >= 4 && 2 * p < n)) return
    ordinary.push([p, value])
    if (k > 0 && k < power.length - 1) {
      const [left, right] = [power[k - 1], power[k + 1]]
      if ((left && right && value / left > 4 && value / right > 4) || (left === 0 && right === 0)) strong.push(p)
    }
  })
  const list = ordinary.sort((a, b) => b[1] - a[1]).slice(0, 3).map(([p]) => p)
  for (const p of strong) if (!list.includes(p)) list.push(p)
  return list
}

/**
 * The autocorrelation of values with their mean taken out (above 3E-7), each lag's sum over its pairs and normalized by
 * lag 0's, through FFTs of the values and of ones padded to twice their length.
 */
function autocorrelation(y: readonly number[]): number[] {
  const n = y.length
  let s = 0
  for (const v of y) s += v
  const mu = s / n
  const z = mu > 3e-7 ? y.map((v) => v - mu) : [...y]
  const flat = [1, ...new Array<number>(n - 1).fill(0)]
  if (!z.some((v) => v)) return flat
  let size = 1
  while (size < 2 * n) size <<= 1
  const lagged = (v: readonly number[]) => {
    const [re, im] = fft(v, size)
    for (let k = 0; k < size; k++) {
      re[k] = re[k] * re[k] + im[k] * im[k]
      im[k] = 0
    }
    // the inverse transform of a real spectrum is the conjugate transform over its size
    const [back] = fft(Array.from(re), size)
    return Array.from({ length: n }, (_, k) => back[k] / size)
  }
  const sums = lagged(z)
  const pairs = lagged(new Array<number>(n).fill(1))
  const scale = sums[0] / pairs[0] / n
  if (scale === 0) return flat
  return [1, ...places(1, n - 1).map((k) => sums[k] / n / pairs[k] / scale)]
}

/** Whether the correlation rises into lag t and falls after it: lines through d lags on each side. */
function smallHill(acf: readonly number[], t: number): boolean {
  const d = t === 5 ? 4 : 2
  const left = places(t - d, t)
  const right = places(t + 1, t + d)
  return line(left, left.map((x) => acf[x]))[0] > 0 && line(right, right.map((x) => acf[x]))[0] < 0
}

/** Whether the correlation over [lo, hi], less the line through its edges, is best split into a rise and a fall. */
function largeHill(acf: readonly number[], lo: number, hi: number): boolean {
  const n = acf.length
  const bx = [...places(lo - 2, lo).filter((x) => x >= 3), ...places(hi, hi + 2).filter((x) => x <= n)]
  const [ba, bb] = line(bx, bx.map((x) => acf[x - 1]))
  const xs = places(lo, hi)
  const ys = xs.map((x) => acf[x] - (ba * x + bb))
  const len = xs.length
  const step = Math.ceil((len - 4) / 100)
  const loss = (x: number[], y: number[], [a, b]: [number, number]) => y.reduce((s, v, i) => s + (v - (a * x[i] + b)) ** 2, 0) / x.length
  let best = Infinity
  let valid = false
  for (let cut = 2; cut < len - 2; cut += step) {
    const [lx, ly, rx, ry] = [xs.slice(0, cut + 1), ys.slice(0, cut + 1), xs.slice(cut), ys.slice(cut)]
    const l = line(lx, ly)
    const r = line(rx, ry)
    const c = loss(lx, ly, l) + loss(rx, ry, r)
    if (c < best) {
      best = c
      valid = l[0] > 0 && r[0] < 0
    }
  }
  return valid
}

/** The period of the correlation's peak over [lo, hi] if it is a hill, and the correlation there; half of it when doubled. */
function peak(acf: readonly number[], lo: number, hi: number, doubled: boolean): [number, number] | undefined {
  const n = acf.length
  if (lo > hi || hi >= n) return undefined
  let t = lo
  for (let x = lo + 1; x <= hi; x++) if (acf[x] > acf[t]) t = x
  const q = doubled ? t >> 1 : t
  if (2.5 * q > n || t === lo || t === hi) return undefined
  let hill: boolean
  if ((t >= 3 && t <= 5) || (doubled && t >= 6 && t <= 11)) hill = smallHill(acf, t)
  else if (hi - lo + 1 < 5) return undefined
  else hill = largeHill(acf, lo, hi)
  return hill ? [q, acf[q]] : undefined
}

/** The interval of lags around a candidate period. */
function interval(n: number, p: number): [number, number] {
  const f = n / p
  let lo = round(n / (f + 1) - 1)
  let hi = round(n / (f - 1) + 1)
  if (lo <= 3) {
    lo = 3
    if (hi < 3) hi = 5
  }
  if (hi > n) hi = lo
  if (p >> 1 < hi) lo = Math.max(lo, p >> 1)
  return [lo, hi]
}

/** The answer of the autocorrelation: the candidates that make a hill, ranked by power and then correlation; 1 for none. */
function correlated(n: number, zs: Spectrum, acf: readonly number[], fast: number, fastPower: number): number {
  const accepted: [number, number, number][] = []
  for (const p of candidates(n, zs)) {
    const [lo, hi] = interval(n, p)
    const found = peak(acf, lo, hi, false) ?? (2 * hi < n ? peak(acf, 2 * lo, 2 * hi, true) : undefined)
    if (!found || found[1] <= 0.45) continue
    const [q, corr] = found
    const same = accepted.find((r) => r[0] === q)
    if (same) same[2] = Math.max(same[2], corr)
    else {
      let bin = 0
      for (let k = 1; k < zs.power.length; k++) if (Math.abs((k + 1) / n - 1 / q) < Math.abs((bin + 1) / n - 1 / q)) bin = k
      accepted.push([q, zs.power[bin], corr])
    }
  }
  if (accepted.length < 2) return accepted.length ? accepted[0][0] : 1
  const ranked = accepted.map((r): [number, number, number] => [...r])
  if (fast >= 2 && fast <= 5) {
    const same = ranked.find((r) => r[0] === fast)
    if (same) [same[1], same[2]] = [Math.max(same[1], fastPower), Math.max(same[2], acf[fast])]
    else ranked.push([fast, fastPower, acf[fast]])
  }
  ranked.sort((a, b) => b[1] - a[1])
  let keep = ranked.length
  for (let i = 1; i < ranked.length; i++)
    if (ranked[i - 1][1] > 2 * ranked[i][1]) {
      keep = i
      break
    }
  let best = ranked[keep - 1]
  for (let i = keep - 2; i >= 0; i--) if (ranked[i][2] > best[2]) best = ranked[i]
  return best[0]
}

/** The strength of a period in the last whole periods of the values: the modulus of their Fourier sum over their count. */
function strength(y: readonly number[], p: number): number {
  const len = Math.trunc(y.length / p) * p
  const tail = y.slice(y.length - len)
  let re = 0
  let im = 0
  tail.forEach((v, i) => {
    re += v * Math.cos((-2 * Math.PI * i) / p)
    im += v * Math.sin((-2 * Math.PI * i) / p)
  })
  return Math.sqrt(re * re + im * im) / len
}

/** The length of the season of a series of two points or more, 0 for none. */
export function detectSeason(series: readonly number[]): number {
  let y = series.slice(-MAX_POINTS)
  if (y.length > 2048) y = y.slice(y.length - 2 ** Math.floor(Math.log2(y.length)))
  const n = y.length
  const [r, z] = preprocessed(y)
  const [fast, fastPower] = quick(n, spectrum(r))
  const zs = spectrum(z)
  const corr = correlated(n, zs, autocorrelation(z), fast, fastPower)
  const chosen = corr === 1 ? fast : fast === 1 || corr >= 6 || 2 * fast > n ? corr : strength(y, fast) > strength(y, corr) ? fast : corr
  return chosen > 1 && chosen <= MAX_SEASON ? chosen : 0
}
