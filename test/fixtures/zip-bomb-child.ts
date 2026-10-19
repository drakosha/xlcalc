/**
 * A zip bomb read in a fresh process (for `xlsx-zip-limits.test.ts`; `--expose-gc`): `xl/workbook.xml` declares 1 MiB and
 * inflates to 400 MiB of spaces (a file of about 400 KB). The heap (`heapUsed` and `arrayBuffers`) is read after the
 * library and the file are loaded, then every millisecond while `readXlsx` runs, since inflating is asynchronous and the
 * event loop turns between the stream's chunks. Prints the peak above the start in MiB and the refusal, as JSON.
 */
import { once } from 'node:events'
import { createDeflateRaw } from 'node:zlib'
import { rawZip } from './raw-zip.js'

const { libUrl } = await import('./target.js')
const { readXlsx } = (await import(libUrl('xlsx.js'))) as typeof import('../../src/xlsx.js')

/** 400 MiB of spaces deflated a MiB at a time, so the process never holds them inflated. */
async function spaces(mib: number): Promise<Uint8Array> {
  const z = createDeflateRaw({ level: 9 })
  const out: Buffer[] = []
  z.on('data', (b: Buffer) => out.push(b))
  const chunk = Buffer.alloc(2 ** 20, 0x20)
  for (let i = 0; i < mib; i++) if (!z.write(chunk)) await once(z, 'drain')
  z.end()
  await once(z, 'end')
  return new Uint8Array(Buffer.concat(out))
}

const file = rawZip([{ name: 'xl/workbook.xml', data: await spaces(400), method: 8, usize: 2 ** 20 }])
const used = () => {
  const m = process.memoryUsage()
  return m.heapUsed + m.arrayBuffers
}
const turn = () => new Promise((done) => setImmediate(done))
await turn()
await turn()
const gc = (globalThis as { gc: () => void }).gc
gc()
gc()
const start = used()
let peak = start
const timer = setInterval(() => (peak = Math.max(peak, used())), 1)
let refusal = ''
try {
  await readXlsx(file)
} catch (e) {
  refusal = (e as Error).message
}
clearInterval(timer)
peak = Math.max(peak, used())
console.log(JSON.stringify({ file: file.length, peak: +((peak - start) / 2 ** 20).toFixed(1), refusal }))
