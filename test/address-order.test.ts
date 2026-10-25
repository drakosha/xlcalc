import { expect, it } from 'vitest'
import { addressOrder, formatAddress, MAX_COLS, MAX_ROWS, parseAddress } from '../src/model.js'
import { mulberry32 } from '../src/synth.js'

/** The order as parsing the address gives it, or the error it throws. */
const parsed = (a: string) => {
  try {
    const { row, col } = parseAddress(a)
    return row * MAX_COLS + col
  } catch (e) {
    return (e as Error).message
  }
}
const read = (a: string) => {
  try {
    return addressOrder(a)
  } catch (e) {
    return (e as Error).message
  }
}

it("reads an address's place as parsing it does, and refuses what parsing refuses", () => {
  const rnd = mulberry32(5)
  const addrs = ['A1', 'XFD1048576', 'ZZZ1', 'AAA999999999999999', 'A1234567890123456', 'A0', 'A01', 'a1', 'A', '1', '', 'AAAA1', 'A1B', 'A-1', 'Ä1', 'A1 ', ' A1', 'A1.5']
  for (let i = 0; i < 2000; i++) addrs.push(formatAddress(Math.floor(rnd() * MAX_ROWS), Math.floor(rnd() * MAX_COLS)))
  for (const a of addrs) expect(read(a), a).toBe(parsed(a))
})
