/**
 * The General format in narrow columns: numbers of every shape (long fractions, large and small magnitudes,
 * negatives, values that round up) in columns from half a character to 14 characters wide.
 */
import { collector, sixths } from './cases.js'

const VALUES = [
  1234567.891, 0.000123456789, 123456789, 3.14159265358979, -3.14159265358979, 99999.99, 1e-10, 1.5e20, 12345678901234, 0.5, -0.5, 100000, 1234.5, 0.30000000000000004,
  999999999999, 9.99999e-5, -1234567, 0.1, 1e100, 2.5e-300, 7, 70, 700, -7, -70, 0.04, -0.04, 1e-5, 5e-5, 1.5, 9.6, 0.96, 99.6, 999999.5, 0.00051,
  12345.678, 1e11, 99999999999, 1e10, -1e-10, -0.000123456789, 0,
  // where positional and scientific forms compete, and the last resort by real width
  1.5e-5, 5.5e-5, 0.00015, 123456.7, 12345.6, 99999.4, -99999.99, 2.5e-100, 1.5e-11, 1 / 3, 2 / 3, 1e-4, 9.999e-5, 1.23e-4, 999999999.5, 0.99999, 9.9999,
  12345678901, 123456789012, -12345.6, -123456.7,
]

export function buildDisplayCases() {
  const { out, add } = collector()
  for (const width of sixths(0.5, 14)) for (const v of VALUES) add('general', v, undefined, { width })
  return out
}
