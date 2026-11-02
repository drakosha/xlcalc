/**
 * A CSV file's bytes as text, and text as a file's bytes. Reading takes a byte order mark (UTF-8, UTF-16LE, UTF-16BE),
 * then UTF-16 without one by its zero bytes in every other place, then UTF-8 if the bytes are valid UTF-8, else the
 * locale's single-byte Windows code page (Windows-1252 for en-US, Windows-1251 for ru-RU), as Excel for Windows reads a
 * file without a mark. The platform's `TextDecoder` decodes, by its WHATWG label, but the locale's code page both ways
 * by the locale's table of it (`Rules.codePage`, the one CHAR and CODE read): Node 20 decodes windows-1252 as Latin-1.
 */
import { InputError } from '../input-error.js'
import type { CodePage } from '../locale/rules/data.js'
import { utf8 } from '../xlsx/zip/platform.js'

type Decoder = { readonly encoding: string; decode(b: Uint8Array): string }
const Decoder = (globalThis as unknown as { TextDecoder?: new (label: string, opts: { fatal: boolean; ignoreBOM: boolean }) => Decoder }).TextDecoder

/** A decoder of `label`, refusing malformed bytes when `fatal`; undefined for a label the platform does not know. */
function decoder(label: string, fatal: boolean): Decoder | undefined {
  if (!Decoder) throw new Error('csv: this platform has no TextDecoder (it needs Node 20.19 or later, or a current browser)')
  try {
    return new Decoder(label, { fatal, ignoreBOM: true })
  } catch (e) {
    if (e instanceof RangeError) return undefined
    throw e
  }
}

/** The WHATWG name of an encoding's label (`UTF8` is `utf-8`, `cp1251` is `windows-1251`), checked: an `InputError` for a label the platform does not know. */
export function encodingName(label: string, path: string): string {
  const d = decoder(label, false)
  if (!d) throw new InputError(path, `expected an encoding's label such as "utf-8", "utf-16le" or "windows-1251", got ${JSON.stringify(label)}`)
  return d.encoding
}

/** The byte order marks, longest first, and the encoding each gives. */
const MARKS: [number[], string][] = [
  [[0xef, 0xbb, 0xbf], 'utf-8'],
  [[0xff, 0xfe], 'utf-16le'],
  [[0xfe, 0xff], 'utf-16be'],
]

const marked = (b: Uint8Array, mark: number[]) => mark.every((x, i) => b[i] === x)

/**
 * UTF-16 without a mark: in the first 1,024 bytes, at least a third of the pairs with a zero byte on one side and none
 * on the other, as ASCII and most Latin and Cyrillic text in UTF-16 has; `le` when the zeros are the odd bytes.
 */
function unmarked16(b: Uint8Array): string | undefined {
  const n = Math.min(b.length, 1024) & ~1
  if (n < 2) return undefined
  let odd = 0
  let even = 0
  for (let i = 0; i < n; i += 2) {
    if (b[i] === 0) even++
    if (b[i + 1] === 0) odd++
  }
  const third = n / 2 / 3
  if (odd >= third && even === 0) return 'utf-16le'
  if (even >= third && odd === 0) return 'utf-16be'
  return undefined
}

/** Bytes of the locale's code page by its table, as `encodeCsv` writes them. */
function pageText(b: Uint8Array, page: CodePage): string {
  const out: string[] = []
  for (let i = 0; i < b.length; i += 8192) out.push(String.fromCharCode(...Array.from(b.subarray(i, i + 8192), (x) => (x < 128 ? x : page.chars.charCodeAt(x - 128)))))
  return out.join('')
}

/**
 * A file's text, its encoding and whether a byte order mark gave it: `encoding` as given (a byte order mark of it skipped), else as the bytes show it,
 * a file without a mark that is no valid UTF-8 in `page`. Malformed bytes in the encoding taken become U+FFFD.
 */
export function decodeCsv(b: Uint8Array, page: CodePage, encoding?: string): { text: string; encoding: string; marked: boolean } {
  const mark = MARKS.find(([m, name]) => (encoding === undefined || encoding === name) && marked(b, m))
  const body = mark ? b.subarray(mark[0].length) : b
  const name = encoding ?? mark?.[1] ?? unmarked16(b)
  if (name !== undefined) return { text: name === page.name ? pageText(body, page) : decoder(name, false)!.decode(body), encoding: name, marked: !!mark }
  try {
    return { text: decoder('utf-8', true)!.decode(b), encoding: 'utf-8', marked: false }
  } catch {
    return { text: pageText(b, page), encoding: page.name, marked: false }
  }
}

/** The encodings `writeCsv` writes: UTF-8 and UTF-16LE, each with its byte order mark unless `bom` is false, and the locale's code page. */
export type WriteEncoding = 'utf-8' | 'utf-16le' | 'ansi'

/** Text as bytes of an encoding; in a code page a character it lacks is `?`, as Windows converts it. Counts those in `lost`. */
export function encodeCsv(text: string, encoding: WriteEncoding, bom: boolean, page: CodePage): { data: Uint8Array; lost: number } {
  if (encoding === 'utf-8') {
    const body = utf8(text)
    if (!bom) return { data: body, lost: 0 }
    const out = new Uint8Array(body.length + 3)
    out.set(MARKS[0][0])
    out.set(body, 3)
    return { data: out, lost: 0 }
  }
  if (encoding === 'utf-16le') {
    const at = bom ? 2 : 0
    const out = new Uint8Array(at + text.length * 2)
    if (bom) out.set(MARKS[1][0])
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i)
      out[at + 2 * i] = c & 0xff
      out[at + 2 * i + 1] = c >> 8
    }
    return { data: out, lost: 0 }
  }
  const byChar = new Map<number, number>()
  for (let i = 0; i < page.chars.length; i++) byChar.set(page.chars.charCodeAt(i), 128 + i)
  const out = new Uint8Array(text.length)
  let n = 0
  let lost = 0
  for (const ch of text) {
    const c = ch.codePointAt(0)!
    const byte = c < 128 ? c : byChar.get(c)
    if (byte === undefined) lost++
    out[n++] = byte ?? 0x3f
  }
  return { data: out.subarray(0, n), lost }
}
