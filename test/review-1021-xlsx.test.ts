/**
 * Reproductions from the review of 21.10 on reading xlsx: a damaged deflate stream fails the read alone, without a
 * rejection left unhandled. Reading a file does not calculate, so the execution modes do not apply here.
 */
import { deflateRawSync } from 'node:zlib'
import { describe, expect, it } from 'vitest'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { rawZip, type RawEntry } from './fixtures/raw-zip.js'

/** The rejections nobody handled while `f` ran, and a moment after, when Node reports them. */
async function unhandledDuring(f: () => Promise<void>): Promise<unknown[]> {
  const seen: unknown[] = []
  const on = (e: unknown) => seen.push(e)
  process.on('unhandledRejection', on)
  try {
    await f()
    await new Promise((resolve) => setTimeout(resolve, 50))
  } finally {
    process.off('unhandledRejection', on)
  }
  return seen
}

describe('xlsx: a damaged deflate stream', () => {
  it('fails the read with the entry named, and leaves no rejection unhandled', async () => {
    const text = new TextEncoder().encode('<a>'.repeat(5000))
    const garbage: RawEntry = { name: 'a.xml', data: new Uint8Array(64).fill(0xff), method: 8, usize: 100 }
    // a valid stream cut in half: it ends before its last block
    const cut: RawEntry = { name: 'a.xml', data: deflateRawSync(text).subarray(0, 20), method: 8, usize: text.length }
    const { data } = await writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: 'hello' } } }] } })
    const archive = openArchive(data)
    const entries = await Promise.all(archive.names.map(async (name): Promise<RawEntry> => ({ name, data: (await archive.bytes(name))! })))
    const damaged = rawZip(entries.map((e) => (e.name === 'xl/sharedStrings.xml' ? { ...e, data: new Uint8Array(32).fill(0xff), method: 8, usize: e.data.length } : e)))
    const unhandled = await unhandledDuring(async () => {
      for (const e of [garbage, cut]) await expect(openArchive(rawZip([e])).bytes('a.xml')).rejects.toThrow(/^xlsx: not a readable zip archive: entry a\.xml is not valid deflate data/)
      await expect(readXlsx(damaged)).rejects.toThrow(/^xlsx: not a readable zip archive: entry xl\/sharedStrings\.xml is not valid deflate data/)
      // the same file whole reads as before
      expect((await readXlsx(rawZip(entries))).book.sheets[0].cells.A1).toEqual({ value: 'hello' })
    })
    expect(unhandled).toEqual([])
  })
})
