/**
 * A book sent to and from a thread as JSON text (`docs/async.md`): its cells in chunks of text, written and read a
 * chunk at a time between turns of the event loop, so the calling side never stops for the whole book. The text is
 * what the pool keeps as a session's snapshot: a quarter of the memory of the book's objects. A book that JSON cannot
 * hold as it is (a function, a date object or a non-finite number in it) goes as before, by structured clone.
 */
import { pause } from '../async/drive.js'
import type { BookInput, CachedValues, CellInput } from '../contract.js'

/** A book as it crosses to or from a thread. */
export type Sent = BookInput | Packed

/** A book as text: everything but the cells (`head`), and each sheet's cells as chunks of `"A1":{…}` entries. */
export interface Packed {
  packed: 1
  head: string
  sheets: string[][]
}

/** The book without its sheets' cells. */
type Head = Omit<BookInput, 'sheets'> & { sheets: Omit<BookInput['sheets'][number], 'cells'>[] }

/** What a book holds when it is handed over: its fields and its sheets' cells taken at once, so later changes of the book do not reach it. */
interface Capture {
  /** The fields but the cells, the book's and its sheets' copied when handed over. */
  rest: Head
  /** `rest` as text; null when JSON cannot hold it. */
  head: string | null
  sheets: { keys: string[]; cells: CellInput[] }[]
}

/** Cells in one chunk: a few milliseconds of writing or reading. */
const CHUNK = 16_384

const plainObject = (v: object) => {
  const p = Object.getPrototypeOf(v)
  return p === Object.prototype || p === null
}

/**
 * A value JSON writes and reads back as it is. `loose`: a field set to undefined is as good as none, as the engine reads
 * its input; a cell's `meta` comes back to the host as given, so there it is strict.
 */
function isJson(v: unknown, loose: boolean): boolean {
  if (v === null || typeof v === 'string' || typeof v === 'boolean') return true
  if (typeof v === 'number') return Number.isFinite(v) && !Object.is(v, -0)
  if (typeof v !== 'object') return false
  if (Array.isArray(v)) {
    for (let i = 0; i < v.length; i++) if (!(i in v) || !isJson(v[i], loose)) return false
    return true
  }
  if (!plainObject(v)) return false
  for (const k in v) {
    const x = (v as Record<string, unknown>)[k]
    if (x === undefined ? !loose : !isJson(x, loose && k !== 'meta')) return false
  }
  return true
}

/** A sheet's cells as two arrays in one pass (`Object.values` of a large cells object takes twice as long). */
function entries(cells: Record<string, CellInput>): { keys: string[]; cells: CellInput[] } {
  const keys: string[] = []
  const values: CellInput[] = []
  for (const k in cells) {
    keys.push(k)
    values.push(cells[k])
  }
  return { keys, cells: values }
}

/** The book as it is now: its head as text, each sheet's cells as arrays; null for input of another shape. */
function capture(book: BookInput): Capture | null {
  const sheets = book && typeof book === 'object' && Array.isArray(book.sheets) ? book.sheets : null
  if (!sheets || !sheets.every((s) => s && typeof s === 'object' && s.cells && typeof s.cells === 'object')) return null
  const rest: Head = { ...book, sheets: sheets.map(({ cells: _, ...s }) => s) }
  return { rest, head: isJson(rest, true) ? JSON.stringify(rest) : null, sheets: sheets.map((s) => entries(s.cells)) }
}

/**
 * The captured book as objects, for structured clone: its fields as they were when it was handed over (read back from
 * the text when there is one, since a book written a chunk at a time comes here after the caller could change it).
 */
function objects(c: Capture): BookInput {
  const head = c.head === null ? c.rest : (JSON.parse(c.head) as Head)
  return { ...head, sheets: head.sheets.map((s, i) => ({ ...s, cells: Object.fromEntries(c.sheets[i].keys.map((k, j) => [k, c.sheets[i].cells[j]])) })) }
}

/** One chunk of a sheet's cells as text, or null when JSON cannot hold one of them. */
function chunk(keys: string[], cells: CellInput[], i: number): string | null {
  let text = ''
  for (let j = i; j < Math.min(i + CHUNK, keys.length); j++) {
    if (!isJson(cells[j], true)) return null
    text += `${j > i ? ',' : ''}${JSON.stringify(keys[j])}:${JSON.stringify(cells[j])}`
  }
  return text
}

/**
 * The book as text, as objects when JSON cannot hold it. The book is taken at once: what the caller changes in it
 * afterwards does not reach the text. A book of one chunk is written at once; a larger one a chunk at a time between
 * turns of the event loop.
 */
export function pack(book: BookInput): Sent | Promise<Sent> {
  const c = capture(book)
  // input of another shape goes as it is: the thread refuses it with InputError
  if (!c) return book
  if (c.head === null) return objects(c)
  const head = c.head
  if (c.sheets.reduce((n, s) => n + s.keys.length, 0) <= CHUNK) {
    const sheets = c.sheets.map(({ keys, cells }) => (keys.length ? [chunk(keys, cells, 0)] : []))
    return sheets.some((x) => x[0] === null) ? objects(c) : { packed: 1, head, sheets: sheets as string[][] }
  }
  return (async () => {
    const sheets: string[][] = []
    for (const { keys, cells } of c.sheets) {
      const chunks: string[] = []
      for (let i = 0; i < keys.length; i += CHUNK) {
        await pause()
        const text = chunk(keys, cells, i)
        if (text === null) return objects(c)
        chunks.push(text)
      }
      sheets.push(chunks)
    }
    return { packed: 1 as const, head, sheets }
  })()
}

/**
 * A file's stored values (`ComputeOptions.cached`) as text, as a book is: they have a book's shape, sheets of cells by
 * address, and go with their sheets' names alone (a result's other fields are not needed).
 */
export function packCached(v: CachedValues): Sent | Promise<Sent> {
  const sheets = (v as { sheets?: unknown } | null)?.sheets
  // input of another shape goes as it is: the thread refuses it with InputError
  if (!Array.isArray(sheets) || !sheets.every((s) => s !== null && typeof s === 'object')) return v as unknown as Sent
  return pack({ sheets: v.sheets.map(({ name, cells }) => ({ name, cells })) } as unknown as BookInput)
}

/** Stored values `packCached` gave, as objects again (in a thread). */
export const unpackCached = (v: Sent): CachedValues => unpack(v) as unknown as CachedValues

/** Data as JSON text when JSON holds it as it is (as a book's cells are checked), else as it is: text the calling side keeps for less than the objects. */
export const packData = <T>(v: T): T | { json: string } => (isJson(v, true) ? { json: JSON.stringify(v) } : v)

/** Data `packData` gave, as objects again. */
export const unpackData = <T>(v: T | { json: string }): T => (v && typeof v === 'object' && 'json' in v && Object.keys(v).length === 1 ? (JSON.parse(v.json) as T) : (v as T))

const isPacked = (b: Sent): b is Packed => (b as Packed).packed === 1

/** The book from text at once (in a thread, where nothing waits on it): each sheet's chunks read as one text. */
export function unpack(b: Sent): BookInput {
  if (!isPacked(b)) return b
  const book = JSON.parse(b.head) as BookInput
  book.sheets.forEach((s, i) => (s.cells = JSON.parse(`{${b.sheets[i].filter((x) => x).join(',')}}`) as Record<string, CellInput>))
  return book
}

/** The book from text a chunk at a time between turns of the event loop (on the calling side). */
export async function unpackSlowly(b: Sent): Promise<BookInput> {
  if (!isPacked(b)) return b
  const book = JSON.parse(b.head) as BookInput
  for (let i = 0; i < book.sheets.length; i++) {
    const cells: Record<string, CellInput> = {}
    for (const chunk of b.sheets[i]) {
      Object.assign(cells, JSON.parse(`{${chunk}}`))
      await pause()
    }
    book.sheets[i].cells = cells
  }
  return book
}
