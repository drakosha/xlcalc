/** Excel's x87 arithmetic rounds to a 64-bit significand before storing a binary64 value. */
const bits = new DataView(new ArrayBuffer(8))
const FRACTION = (1n << 52n) - 1n
const MIN_NORMAL = 2 ** -1022
const SPLIT = 2 ** 27 + 1
const MIDPOINT = 0.5 - 2 ** -12

/** An odd binary64 significand is the only result that double rounding can change. */
function odd(x: number): boolean {
  if (!Number.isFinite(x) || x === 0) return false
  bits.setFloat64(0, x)
  return (bits.getUint32(4) & 1) !== 0
}

/** The positive significand scaled into [1, 2), also for a subnormal input. */
function normalized(x: number): number {
  x = Math.abs(x)
  bits.setFloat64(0, x < MIN_NORMAL ? x * 2 ** 52 : x)
  bits.setUint32(0, (bits.getUint32(0) & 0xfffff) | 0x3ff00000)
  return bits.getFloat64(0)
}

const mantissa = (x: number): bigint => {
  bits.setFloat64(0, x)
  return (bits.getBigUint64(0) & FRACTION) | (1n << 52n)
}

/** The adjacent even significand, toward a larger or smaller magnitude. */
function adjacent(x: number, larger: boolean): number {
  bits.setFloat64(0, x)
  bits.setBigUint64(0, bits.getBigUint64(0) + (larger ? 1n : -1n))
  return bits.getFloat64(0)
}

/** Exact residual of a product of the normalized operands used here (Dekker splitting). */
function productError(a: number, b: number, product: number): number {
  const ac = SPLIT * a, ah = ac - (ac - a), al = a - ah
  const bc = SPLIT * b, bh = bc - (bc - b), bl = b - bh
  return ((ah * bh - product) + ah * bl + al * bh) + al * bl
}

export function pivotAdd(a: number, b: number): number {
  const out = a + b
  if (!odd(out)) return out
  const z = out - a, error = (a - (out - z)) + (b - z)
  if (error === 0) return out
  bits.setFloat64(0, out)
  const exponent = (bits.getUint32(0) >>> 20) & 2047
  const unit = 2 ** (exponent ? exponent - 1075 : -1074)
  return Math.abs(error) >= unit * MIDPOINT ? adjacent(out, (error > 0) === (out > 0)) : out
}

export function pivotMultiply(a: number, b: number): number {
  const out = a * b
  if (!odd(out) || Math.abs(out) < MIN_NORMAL) return out
  const x = normalized(a), y = normalized(b), product = x * y
  const error = productError(x, y, product), unit = product < 2 ? 2 ** -52 : 2 ** -51
  return Math.abs(error) >= unit * MIDPOINT ? adjacent(out, error > 0) : out
}

export function pivotDivide(a: number, b: number): number {
  const out = a / b
  if (!odd(out) || Math.abs(out) < MIN_NORMAL) return out
  const x = normalized(a), y = normalized(b), q = x / y, product = q * y
  const error = ((x - product) - productError(q, y, product)) / y, unit = q < 1 ? 2 ** -53 : 2 ** -52
  if (Math.abs(error) < unit * 0.499) return out
  const denominator = mantissa(y)
  const remainder = (mantissa(x) << BigInt(q < 1 ? 53 : 52)) - mantissa(q) * denominator
  const magnitude = remainder < 0n ? -remainder : remainder
  return magnitude * 4096n >= denominator * 2047n ? adjacent(out, remainder > 0n) : out
}

export function pivotRoot(x: number): number {
  const out = Math.sqrt(x)
  if (!odd(out)) return out
  let scaled = x
  let adjustment = 0
  if (x < MIN_NORMAL) { scaled *= 2 ** 52; adjustment = 52 }
  bits.setFloat64(0, scaled)
  const exponent = ((bits.getUint32(0) >>> 20) & 2047) - 1023 - adjustment
  const m = normalized(x), q = normalized(out), source = exponent & 1 ? m * 2 : m, product = q * q
  const error = ((source - product) - productError(q, q, product)) / (2 * q)
  if (Math.abs(error) < 2 ** -52 * 0.499) return out
  const threshold = (mantissa(q) << 12n) + (error > 0 ? 2047n : -2047n)
  const input = mantissa(m) << BigInt(exponent & 1 ? 77 : 76)
  return (error > 0 ? input >= threshold * threshold : input <= threshold * threshold) ? adjacent(out, error > 0) : out
}
