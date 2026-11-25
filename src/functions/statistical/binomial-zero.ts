import { dyadic } from './exact-moments.js'

const LOG_DIRECT_THRESHOLD = 0.375

/** Excel's ln(1 - p), with the intermediate operations rounded to doubles. */
function logOneMinus(p: number): number {
  if (p >= LOG_DIRECT_THRESHOLD) return Math.log(1 - p)
  if (p <= 2 * Number.EPSILON) return -p * (1 + p / 2)

  // ln(1 - p) = -2 atanh(p / (2 - p)). Add the leading term after summing the smaller terms.
  const head = p / (2 - p)
  const square = head * head
  let power = head
  let tail = 0
  for (let divisor = 3; ; divisor += 2) {
    power *= square
    tail += power / divisor
    if (power < tail * Number.EPSILON) return -2 * (head + tail)
  }
}

const PRECISION = 160n
const ONE = 1n << PRECISION
/** ln(2) scaled by 2^160 and rounded down. */
const LN2 = 0xb17217f7d1cf79abc9e3b39803f2f6af40f34326n
/** x87 loads log2(e) with this 64-bit significand and an exponent of 0. */
const LOG2_E = 0xb8aa3b295c17f0bcn

/** The natural logarithm of a positive dyadic number, scaled by ONE. */
function fixedLog(mantissa: bigint, scale: number): bigint {
  const leadingBit = mantissa.toString(2).length - 1
  const leading = 1n << BigInt(leadingBit)
  // Normalize to [1, 2) for the atanh series. Its argument is at most 1/3.
  const z = (mantissa - leading) * ONE / (mantissa + leading)
  const square = z * z / ONE
  let sum = z
  let power = z
  for (let divisor = 3n; power !== 0n; divisor += 2n) {
    power = power * square / ONE
    sum += power / divisor
  }
  return 2n * sum + BigInt(scale + leadingBit) * LN2
}

/** The fixed-point value of x * log2(e), after x87's 64-bit significand rounding. */
function exp2Argument(x: number): bigint {
  const [mantissa, scale] = dyadic(-x)
  const product = mantissa * LOG2_E
  // A normal double's 53-bit significand times LOG2_E has either 116 or 117 bits.
  const discarded = product >> 116n ? 53n : 52n
  const kept = product >> discarded
  // Add half an ULP, with a bias of one bit less for an even tie.
  const bias = (1n << (discarded - 1n)) - 1n + (kept & 1n)
  const rounded = (product + bias) >> discarded
  return -(rounded << (BigInt(scale) - 63n + discarded + PRECISION))
}

/** Excel's zero-quantile branch uses a strict comparison after rounding the exponential. */
export function binomialZeroExceeds(trials: number, probability: number, alpha: number): boolean {
  const logProbability = trials * logOneMinus(probability)
  // This close to zero, exp rounds to 1.
  if (logProbability >= -(2 ** -54)) return true
  const approximate = Math.exp(logProbability)
  // An error in ln(q) is multiplied by trials before exp; allow for both roundings.
  const margin = 1e-12 * alpha + Number.MIN_VALUE
  if (approximate === 0 || Math.abs(approximate - alpha) > margin) return approximate > alpha

  let preciseLogProbability = logProbability
  if (probability >= LOG_DIRECT_THRESHOLD) {
    const [mantissa, scale] = dyadic(1 - probability)
    const log = Number(fixedLog(mantissa, scale)) / Number(ONE)
    preciseLogProbability = trials * log
  }
  // Rounding can exceed alpha only past the midpoint to its next double. Take the
  // logarithm of that exact midpoint to compare without rounding the exponential.
  const [lower, lowerScale] = dyadic(alpha)
  return exp2Argument(preciseLogProbability) * LN2 / ONE > fixedLog(2n * lower + 1n, lowerScale - 1)
}
