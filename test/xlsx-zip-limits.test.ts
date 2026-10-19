/**
 * How large a file's parts may be once inflated (`ReadXlsxOptions.limits`). A part that inflates past the size it declares
 * is refused as it inflates: a zip bomb that declares 1 MiB and inflates to 400 MiB leaves the heap within 32 MiB of where
 * it was, in a fresh process (`fixtures/zip-bomb-child.ts`; it reached about 800 MiB when the part was inflated whole before
 * its size was checked). A part past `entryBytes`, parts past `totalBytes` together and records of the central directory
 * whose bytes overlap are refused as a damaged file is; a file within the bounds reads as before.
 */
import { execFileSync } from 'node:child_process'
import { deflateRawSync } from 'node:zlib'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import type { XlsxLimits } from '../src/xlsx/types.js'
import { rawZip } from './fixtures/raw-zip.js'

const child = fileURLToPath(new URL('./fixtures/zip-bomb-child.ts', import.meta.url))

describe('xlsx: the bounds of a file once inflated', () => {
  it('a part that inflates past the size it declares is refused before it is inflated whole', () => {
    const r = JSON.parse(execFileSync(process.execPath, ['--expose-gc', '--import', 'tsx', child], { encoding: 'utf8' })) as { file: number; peak: number; refusal: string }
    expect(r.file).toBeLessThan(500_000)
    expect(r.peak, JSON.stringify(r)).toBeLessThan(32)
    expect(r.refusal).toBe('xlsx: not a readable zip archive: entry xl/workbook.xml inflates to more than the 1048576 bytes it declares')
  }, 120_000)

  it('a part past entryBytes, and parts past totalBytes together; a file within them reads as before', async () => {
    const cells = Object.fromEntries(Array.from({ length: 200 }, (_, i) => [`A${i + 1}`, { formula: `ROW()*${i}` }]))
    const { data } = await writeXlsx({ book: { sheets: [{ name: 'S', cells }] } })
    const archive = openArchive(data)
    const sizes = await Promise.all(archive.names.map(async (n) => [n, (await archive.bytes(n))!.length] as const))
    const [largest, most] = sizes.reduce((a, b) => (b[1] > a[1] ? b : a))
    const all = sizes.reduce((n, [, s]) => n + s, 0)
    expect(largest).toBe('xl/worksheets/sheet1.xml')
    await expect(readXlsx(data, { limits: { entryBytes: most - 1 } })).rejects.toThrow(`xlsx: not a readable zip archive: entry xl/worksheets/sheet1.xml is larger than ${most - 1} bytes (limits.entryBytes)`)
    await expect(readXlsx(data, { limits: { totalBytes: most } })).rejects.toThrow(`xlsx: not a readable zip archive: the parts read add up to more than ${most} bytes (limits.totalBytes)`)
    const read = await readXlsx(data, { limits: { entryBytes: most, totalBytes: all } })
    expect(read.book.sheets[0].cells.A200).toEqual({ formula: 'ROW()*199' })
    expect(read).toEqual(await readXlsx(data))
  })

  it('a part over 16 MiB that inflates more than limits.ratio times its size is refused before it is inflated; a part past 512 MiB by default', async () => {
    const spaces = (mib: number) => ({ name: 'xl/workbook.xml', data: new Uint8Array(deflateRawSync(Buffer.alloc(mib * 2 ** 20, 0x20))), method: 8, usize: mib * 2 ** 20 })
    // the length a part reads to, or its refusal: a short answer either way, which a failed expectation prints whole
    const read = (file: Uint8Array, limits?: XlsxLimits) => openArchive(file, limits).bytes('xl/workbook.xml').then((b) => b?.length, (e: Error) => e.message)
    const big = spaces(17)
    const refusal = `xlsx: not a readable zip archive: entry xl/workbook.xml inflates to ${big.usize} bytes from ${big.data.length}, more than 200 times its size (limits.ratio)`
    expect(await read(rawZip([big]))).toBe(refusal)
    expect(await readXlsx(rawZip([big])).then(() => 'read', (e: Error) => e.message)).toBe(refusal)
    expect(await read(rawZip([big]), { ratio: 1e6 })).toBe(big.usize)
    // a part of 16 MiB inflates whatever its ratio
    const small = spaces(16)
    expect(await read(rawZip([small]))).toBe(small.usize)
    // a part is bounded at 512 MiB by default, which a JavaScript string could hardly hold anyway
    expect(await read(rawZip([{ ...small, usize: 2 ** 29 + 1 }]), { ratio: 1e9 })).toBe(`xlsx: not a readable zip archive: entry xl/workbook.xml is larger than ${2 ** 29} bytes (limits.entryBytes)`)
    expect(await read(rawZip([{ ...small, usize: 2 ** 29 }]), { ratio: 1e9 })).toBe(`xlsx: not a readable zip archive: entry xl/workbook.xml inflates to ${small.usize} bytes, not ${2 ** 29}`)
  })

  it('records of the central directory whose bytes overlap', async () => {
    const xml = new TextEncoder().encode('<workbook/>')
    const one = { name: 'xl/workbook.xml', data: deflateRawSync(xml), method: 8, usize: xml.length }
    const overlapping = rawZip([one], [{ name: 'xl/workbook.xml', at: 0 }, { name: 'xl/other.xml', at: 0 }])
    expect(() => openArchive(overlapping)).toThrow('xlsx: not a readable zip archive: entries xl/workbook.xml and xl/other.xml overlap')
    await expect(readXlsx(overlapping)).rejects.toThrow('entries xl/workbook.xml and xl/other.xml overlap')
    expect(await openArchive(rawZip([one, { ...one, name: 'xl/other.xml' }])).text('xl/other.xml')).toBe('<workbook/>')
  })
})
