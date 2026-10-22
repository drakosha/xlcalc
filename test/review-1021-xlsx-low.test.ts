/**
 * Reproductions from the review of 21.10 on reading xlsx files that are damaged or written by hand: a zip archive whose
 * entries are cut short, named twice or do not match their CRC-32 is refused with the reader's error. Reading a file
 * does not calculate, so the execution modes do not apply here.
 */
import { deflateRawSync } from 'node:zlib'
import { describe, expect, it } from 'vitest'
import { openArchive } from '../src/xlsx/zip/read.js'
import { rawZip, type RawEntry } from './fixtures/raw-zip.js'

const enc = new TextEncoder()
const entry = (name: string, text: string): RawEntry => ({ name, data: enc.encode(text) })

/** The archive with its first central directory record's field at `at` (4 bytes) set to `value`. */
function patched(zip: Uint8Array, at: number, value: (zip: Uint8Array) => number): Uint8Array {
  const out = zip.slice()
  const v = new DataView(out.buffer)
  const dir = v.getUint32(out.length - 22 + 16, true)
  v.setUint32(dir + at, value(out), true)
  return out
}

const refusal = (zip: Uint8Array, name = 'a.xml') => openArchive(zip).bytes(name).then(() => 'read', (e: Error) => `${e.constructor.name}: ${e.message}`)

describe('xlsx: a damaged zip archive (x-xlsx-3, x-xlsx-12, x-xlsx-13)', () => {
  it("an entry whose local header or data lies past the end of the file is refused with the reader's error", async () => {
    const zip = rawZip([entry('a.xml', '<a/>'.repeat(10))])
    const offset = 42
    for (const at of [(z: Uint8Array) => z.length + 100, () => 0x7fffffff, (z: Uint8Array) => z.length - 2, (z: Uint8Array) => z.length - 29]) {
      expect(await refusal(patched(zip, offset, at))).toBe('Error: xlsx: not a readable zip archive: no local header for a.xml')
    }
    // the data cut short: the entry says it has more bytes than the file
    expect(await refusal(patched(zip, 20, () => 1000))).toBe('Error: xlsx: not a readable zip archive: entry a.xml is cut short')
  })

  it('two entries of one name, in any case, are refused', () => {
    const a = entry('xl/worksheets/sheet1.xml', '<a/>')
    for (const b of [{ ...a, data: enc.encode('<b/>') }, { ...a, name: 'XL/WORKSHEETS/SHEET1.XML' }]) {
      expect(() => openArchive(rawZip([a, b]))).toThrow(`xlsx: not a readable zip archive: entries xl/worksheets/sheet1.xml and ${b.name} have the same name`)
    }
  })

  it('an entry whose bytes do not match its CRC-32, or a stored one whose sizes differ, is refused; a sound one reads', async () => {
    const text = '<worksheet/>'
    const deflated: RawEntry = { name: 'a.xml', data: deflateRawSync(enc.encode(text)), method: 8, usize: text.length }
    const damaged = 'Error: xlsx: not a readable zip archive: entry a.xml is damaged (its CRC-32 does not match)'
    expect(await refusal(rawZip([{ ...entry('a.xml', text), crc: 0xdeadbeef }]))).toBe(damaged)
    expect(await refusal(rawZip([{ ...deflated, crc: 0xdeadbeef }]))).toBe(damaged)
    expect(await refusal(rawZip([{ ...entry('a.xml', text), usize: 5 }]))).toBe('Error: xlsx: not a readable zip archive: entry a.xml is stored in 12 bytes but says it holds 5')
    expect(await openArchive(rawZip([deflated])).text('a.xml')).toBe(text)
    expect(await openArchive(rawZip([entry('a.xml', text)])).text('a.xml')).toBe(text)
  })
})
