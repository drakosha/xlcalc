/**
 * Paste Special against the answers of Excel for Mac (`bench/coverage/paste.ts`): each case's book after
 * `editStructure`'s `copyRange` with `paste` and `transpose`, written and read back, shows what the file Excel saved
 * shows (`bench/coverage/paste/shown.ts`). A paste Excel refused (it leaves the book as it was, with or without an
 * error) is one `editStructure` refuses. Kinds the engine does not have are left out; a known difference must still
 * differ.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { writeXlsx } from '../src/xlsx.js'
import type { PasteAnswer, PasteCase, PasteCases } from '../bench/coverage/paste/cases.js'
import { ENGINE_KINDS, engineOutcome, shownFile } from '../bench/coverage/paste/shown.js'

const SETS = ['probe', 'held', 'merges', 'fit']

/** Where the engine pastes otherwise than Excel, by set and case. */
const KNOWN: Record<string, string> = {
  'held/table-formats': "Excel pastes the look a table's style gives its cells as their formats; the engine does not draw table styles",
}

describe.each(SETS)('paste set %s', (set) => {
  const stored = JSON.parse(readFileSync(new URL(`./excel/paste/${set}.json`, import.meta.url), 'utf8')) as PasteCases & { answers: Record<string, PasteAnswer> }
  const cases = stored.cases.filter((c) => ENGINE_KINDS.includes(c.what))
  it.each(cases.map((c) => [c.id, c] as [string, PasteCase]))('%s', async (_, c) => {
    const book = stored.books[c.book]
    const saved = await shownFile(new Uint8Array(readFileSync(new URL(`./excel/paste/${set}/${c.id}.xlsx`, import.meta.url))))
    const r = await engineOutcome(book, c)
    if (KNOWN[`${set}/${c.id}`]) expect('refused' in r ? null : r.shown, KNOWN[`${set}/${c.id}`]).not.toEqual(saved)
    else if ('refused' in r) expect(saved, `${c.id} refused: ${r.refused}`).toEqual(await shownFile((await writeXlsx({ book })).data))
    else expect(r.shown).toEqual(saved)
  })
})
