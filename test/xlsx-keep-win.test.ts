/**
 * `xlcalc/xlsx/keep` in Excel for Windows: the files `writeKept` wrote for the Mac's keep sets (`open`, the edits over
 * the made files; `struct-open`, the operations of the structure; `formats-open`, the operations over the file of
 * formats), which Excel for Windows opened through COM (a file it would repair is refused), computed and saved
 * (`test/excel/win/keep/<set>.json`, `<set>/`, by `bench/windows/run.ts keep <set>`). Each file is the one Excel saw, it
 * opened without writing a repair log, and the file it saved holds what Excel for Mac saved, as predicted from the Mac's
 * answers (`bench/coverage/keep/win.predicted.json`): the parts counted and their references, or the cell formats.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { partsOf, refsOf } from '../bench/coverage/keep/books.js'
import { formatsOf } from '../bench/coverage/keep/formats.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import { KEEP_OPEN, type OpenAnswer } from '../bench/windows/keep.js'

const dir = new URL('./excel/win/keep/', import.meta.url)
const predicted = JSON.parse(readFileSync(new URL('../bench/coverage/keep/win.predicted.json', import.meta.url), 'utf8')) as {
  sets: Record<string, Record<string, { opens: true; parts?: Record<string, number>; refs?: Record<string, unknown>; formats?: unknown }>>
}

describe.each(Object.keys(KEEP_OPEN))('Excel for Windows over the files writeKept wrote for %s', (set) => {
  it('opened each without repairing it and saved what Excel for Mac saved', async () => {
    const run = JSON.parse(readFileSync(new URL(`${set}.json`, dir), 'utf8')) as { sent: Record<string, string>; answers: Record<string, OpenAnswer>; rejected?: Record<string, string> }
    const files = await KEEP_OPEN[set]()
    expect(files.map((f) => f.id).sort()).toEqual(Object.keys(predicted.sets[set]).sort())
    for (const f of files) {
      expect(run.sent[f.id], `${f.id}: writeKept changed since Excel saw the file: run bench/windows/run.ts keep ${set}`).toBe(await partsHash(f.data))
      expect(run.rejected?.[f.id], f.id).toBeUndefined()
      expect(run.answers[f.id].opened[2], f.id).toEqual([])
      const saved = new Uint8Array(readFileSync(new URL(`${set}/${f.id}.xlsx`, dir)))
      const p = predicted.sets[set][f.id]
      if (p.formats) expect(await formatsOf(saved), f.id).toEqual(p.formats)
      else {
        expect(await partsOf(saved), f.id).toEqual(p.parts)
        if (p.refs) expect(await refsOf(saved), f.id).toEqual(p.refs)
      }
    }
  }, 120_000)
})
