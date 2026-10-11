/**
 * What identifies an xlsx file sent to Excel: sha256 of its parts' names and bytes, not of the archive. The platform's
 * deflate (`CompressionStream`) writes other bytes with another zlib (Linux, another Node), while the parts stay the same.
 * And a file changed by hand after `writeXlsx` wrote it (`editParts`).
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

/** The file of each case, as `file` writes it, and what identifies each (`partsHash`). */
export async function caseFiles<C extends { id: string }>(cases: C[], file: (c: C) => Promise<Uint8Array>): Promise<{ files: Map<string, Uint8Array>; sent: Record<string, string> }> {
  const files = new Map(await Promise.all(cases.map(async (c) => [c.id, await file(c)] as const)))
  const sent = Object.fromEntries(await Promise.all(cases.map(async (c) => [c.id, await partsHash(files.get(c.id)!)] as const)))
  return { files, sent }
}
