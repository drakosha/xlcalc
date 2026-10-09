/**
 * What identifies an xlsx file sent to Excel: sha256 of its parts' names and bytes, not of the archive. The platform's
 * deflate (`CompressionStream`) writes other bytes with another zlib (Linux, another Node), while the parts stay the same.
 */
import { createHash } from 'node:crypto'
import { openArchive } from '../../../src/xlsx/zip/read.js'

export async function partsHash(data: Uint8Array): Promise<string> {
  const archive = openArchive(data)
  const h = createHash('sha256')
  for (const name of archive.names) {
    const bytes = (await archive.bytes(name))!
    h.update(`${name}\0${bytes.length}\0`).update(bytes)
  }
  return h.digest('hex')
}
