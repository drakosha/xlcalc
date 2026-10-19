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
    await expect(readXlsx(data, { limits: { totalBytes: most } })).rejects.toThrow(`xlsx: not a readable zip archive: the parts read are larger than ${most} bytes together (limits.totalBytes)`)
    const read = await readXlsx(data, { limits: { entryBytes: most, totalBytes: all } })
    expect(read.book.sheets[0].cells.A200).toEqual({ formula: 'ROW()*199' })
    expect(read).toEqual(await readXlsx(data))
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
