/**
 * The web platform's codecs the xlsx subpath needs, looked up at call time: raw deflate streams
 * (`CompressionStream`, `DecompressionStream` with `deflate-raw`, Node 20.19+ and current browsers) and UTF-8
 * (`TextEncoder`, `TextDecoder`). The engine itself needs none of them.
 */

interface Pair {
  readable: { getReader(): { read(): Promise<{ done: boolean; value?: Uint8Array }> } }
  writable: { getWriter(): { write(chunk: Uint8Array): Promise<void>; close(): Promise<void> } }
}

interface Globals {
  CompressionStream?: new (format: 'deflate-raw') => Pair
  DecompressionStream?: new (format: 'deflate-raw') => Pair
  TextEncoder?: new () => { encode(s: string): Uint8Array }
  TextDecoder?: new (label: 'utf-8', opts: { fatal: boolean }) => { decode(b: Uint8Array): string }
}

const g = globalThis as unknown as Globals

function need<K extends keyof Globals>(name: K): NonNullable<Globals[K]> {
  const v = g[name]
  if (!v) throw new Error(`xlsx: this platform has no ${name} (Node 20.19 or later, or a current browser, is needed)`)
  return v as NonNullable<Globals[K]>
}

/** All bytes a stream pair gives for `data`; the write runs alongside the read so neither side blocks. */
async function through(pair: Pair, data: Uint8Array): Promise<Uint8Array> {
  const writer = pair.writable.getWriter()
  const written = writer.write(data).then(() => writer.close())
  const reader = pair.readable.getReader()
  const parts: Uint8Array[] = []
  let size = 0
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    parts.push(value!)
    size += value!.length
  }
  await written
  return concat(parts, size)
}

export function concat(parts: Uint8Array[], size = parts.reduce((n, p) => n + p.length, 0)): Uint8Array {
  if (parts.length === 1) return parts[0]
  const out = new Uint8Array(size)
  let at = 0
  for (const p of parts) {
    out.set(p, at)
    at += p.length
  }
  return out
}

export const inflate = (data: Uint8Array): Promise<Uint8Array> => through(new (need('DecompressionStream'))('deflate-raw'), data)

export const deflate = (data: Uint8Array): Promise<Uint8Array> => through(new (need('CompressionStream'))('deflate-raw'), data)

let encoder: { encode(s: string): Uint8Array } | undefined
let decoder: { decode(b: Uint8Array): string } | undefined

export const utf8 = (s: string): Uint8Array => (encoder ??= new (need('TextEncoder'))()).encode(s)

/** UTF-8 text; malformed bytes are replaced, as XML readers of office files do. */
export const text = (b: Uint8Array): string => (decoder ??= new (need('TextDecoder'))('utf-8', { fatal: false })).decode(b)
