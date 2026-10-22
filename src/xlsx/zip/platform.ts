/**
 * The web platform's codecs the xlsx subpath needs, looked up at call time: raw deflate streams
 * (`CompressionStream`, `DecompressionStream` with `deflate-raw`, Node 20.19+ and current browsers) and UTF-8
 * (`TextEncoder`, `TextDecoder`). The engine itself needs none of them.
 */

interface Pair {
  readable: { getReader(): { read(): Promise<{ done: boolean; value?: Uint8Array }>; cancel(): Promise<void> } }
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
  if (!v) throw new Error(`xlsx: this platform has no ${name} (it needs Node 20.19 or later, or a current browser)`)
  return v as NonNullable<Globals[K]>
}

/**
 * All bytes a stream pair gives for `data`, or undefined as soon as they pass `max`: the stream is cancelled then, so no
 * more is made. The write runs alongside the read so neither side blocks; `each` sees each chunk as it comes.
 */
async function through(pair: Pair, data: Uint8Array, max = Infinity, each?: (chunk: Uint8Array) => void): Promise<Uint8Array | undefined> {
  const writer = pair.writable.getWriter()
  const written = writer.write(data).then(() => writer.close())
  // damaged data fails the read and the write alike: the read's error is the one thrown, the write's is not left unhandled
  written.catch(() => {})
  const reader = pair.readable.getReader()
  const parts: Uint8Array[] = []
  let size = 0
  for (;;) {
    const { done, value } = await reader.read()
    if (done) break
    size += value!.length
    if (size > max) {
      // the cancelled stream fails the write, or leaves it waiting: it is not awaited
      await reader.cancel()
      return undefined
    }
    each?.(value!)
    parts.push(value!)
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

/** The inflated bytes of `data`, undefined when they are more than `max`; `each` sees each chunk as it is inflated. */
export const inflate = (data: Uint8Array, max: number, each?: (chunk: Uint8Array) => void): Promise<Uint8Array | undefined> =>
  through(new (need('DecompressionStream'))('deflate-raw'), data, max, each)

export const deflate = (data: Uint8Array): Promise<Uint8Array> => through(new (need('CompressionStream'))('deflate-raw'), data) as Promise<Uint8Array>

let encoder: { encode(s: string): Uint8Array } | undefined
let decoder: { decode(b: Uint8Array): string } | undefined

export const utf8 = (s: string): Uint8Array => (encoder ??= new (need('TextEncoder'))()).encode(s)

/** UTF-8 text; malformed bytes are replaced, as XML readers of office files do. */
export const text = (b: Uint8Array): string => (decoder ??= new (need('TextDecoder'))('utf-8', { fatal: false })).decode(b)
