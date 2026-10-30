/**
 * A saved book of a pivot set written again by the engine before Excel opens it (`PivotCase.rewrite`): `xlsx` read
 * (`readXlsx`) and written anew (`writeXlsx`) with its values, `kept` opened (`openKept`) and written over the file
 * (`writeKept`) as it is. Excel's reading of the book right after opening, and after a refresh, is then the engine's
 * file's: the PivotTables as the engine wrote them, and as Excel lays them out again from the cache.
 */
import { openKept, writeKept } from '../../../src/xlsx-keep.js'
import { readXlsx, writeXlsx } from '../../../src/xlsx.js'

export type Rewrite = 'xlsx' | 'kept'

export async function rewritten(data: Uint8Array, how: Rewrite | undefined): Promise<Uint8Array> {
  if (!how) return data
  const read = await (how === 'xlsx' ? readXlsx : openKept)(data)
  return how === 'xlsx' ? (await writeXlsx({ book: read.book }, { values: 'omit' })).data : (await writeKept(data, { book: read.book })).data
}
