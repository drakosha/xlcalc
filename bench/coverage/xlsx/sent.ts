/**
 * What identifies an xlsx file sent to Excel: sha256 of its parts' names and bytes, not of the archive. The platform's
 * deflate (`CompressionStream`) writes other bytes with another zlib (Linux, another Node), while the parts stay the same.
 * And a file changed by hand after `writeXlsx` wrote it (`editParts`, `rawFile`).
 */
import { createHash } from 'node:crypto'
import { openArchive } from '../../../src/xlsx/zip/read.js'
import { zip } from '../../../src/xlsx/zip/write.js'

export async function partsHash(data: Uint8Array): Promise<string> {
  const archive = openArchive(data)
  const h = createHash('sha256')
  for (const name of archive.names) {
    const bytes = (await archive.bytes(name))!
    h.update(`${name}\0${bytes.length}\0`).update(bytes)
  }
  return h.digest('hex')
}

/** The file with the text of each part `edit` names changed by it (`edit` answers undefined for a part it leaves). */
export async function editParts(data: Uint8Array, edit: (name: string, text: string) => string | undefined): Promise<Uint8Array> {
  const archive = openArchive(data)
  const enc = new TextEncoder()
  const files: { name: string; data: Uint8Array }[] = []
  for (const name of archive.names) {
    const changed = name.endsWith('.xml') ? edit(name, (await archive.text(name))!) : undefined
    files.push({ name, data: changed === undefined ? (await archive.bytes(name))! : enc.encode(changed) })
  }
  return zip(files)
}

/** A change by hand to a part of a written file, as another writer or a hand edit leaves it: every `from` in the part's text becomes `to`. */
export interface RawEdit {
  part: string
  from: string
  to: string
}

/** The file with a case's changes by hand made; a change whose text its part does not hold throws, so no case goes out unchanged. */
export async function rawFile(data: Uint8Array, edits: readonly RawEdit[] = []): Promise<Uint8Array> {
  if (!edits.length) return data
  const done = new Set<RawEdit>()
  const out = await editParts(data, (name, text) => {
    let x = text
    for (const e of edits.filter((e) => e.part === name && x.includes(e.from))) {
      x = x.replaceAll(e.from, e.to)
      done.add(e)
    }
    return x === text ? undefined : x
  })
  const missed = edits.filter((e) => !done.has(e))
  if (missed.length) throw new Error(`no ${missed.map((e) => `${JSON.stringify(e.from)} in ${e.part}`).join(', ')}`)
  return out
}

/** The file of each case, as `file` writes it, and what identifies each (`partsHash`). */
export async function caseFiles<C extends { id: string }>(cases: C[], file: (c: C) => Promise<Uint8Array>): Promise<{ files: Map<string, Uint8Array>; sent: Record<string, string> }> {
  const files = new Map(await Promise.all(cases.map(async (c) => [c.id, await file(c)] as const)))
  const sent = Object.fromEntries(await Promise.all(cases.map(async (c) => [c.id, await partsHash(files.get(c.id)!)] as const)))
  return { files, sent }
}
