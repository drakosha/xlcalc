/**
 * `to15` gives what rounding to 15 significant digits through text gives, for every kind of number: its shortcut for a
 * number of hundredths is the same rounding.
 */
import { describe, expect, it } from 'vitest'
import { mulberry32 } from '../src/synth.js'
import { to15 } from '../src/values.js'

const byText = (x: number) => (Number.isInteger(x) && Math.abs(x) < 1e15 ? x : Number(x.toPrecision(15)))

describe('to15', () => {
  it('rounds as text does: hundredths of every size, numbers of many digits, arithmetic on hundredths, edges', () => {
    const rnd = mulberry32(15)
    const xs: number[] = [0.1 + 0.2, 1.005, 2.675, 9999999999999.99, 1e13 - 0.01, 1e13 + 0.01, -1e13 + 0.01, 999999999999.995, 4503599627370.49, 5e-324, -0, NaN, Infinity, -Infinity]
    for (let i = 0; i < 100_000; i++) {
      const cents = Math.round(rnd() * 10 ** Math.floor(rnd() * 16) * 100) / 100
      const a = Math.round(rnd() * 1e6) / 100
      xs.push(rnd() < 0.5 ? cents : -cents, rnd() * 10 ** Math.floor(rnd() * 40 - 20), (Math.round(rnd() * 1e15) / 1e15) * 10 ** Math.floor(rnd() * 30 - 15), a + cents, a - cents, a * 3, a / 7)
    }
    for (const x of [...xs]) for (let k = -3; k <= 3; k++) xs.push(x + k * Number.EPSILON * Math.max(1, Math.abs(x)))
    const differ = xs.filter((x) => !Object.is(to15(x), byText(x)))
    expect(differ).toEqual([])
  })
})
