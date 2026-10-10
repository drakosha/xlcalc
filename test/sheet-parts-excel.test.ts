/**
 * Notes, hyperlinks and frozen panes against Excel 16 for Mac (`test/excel/parts/`, see `test/sheet-parts-excel.ts`):
 * Excel read in each book what the engine holds after the case (the book as written, with what Excel made, after the
 * structure operation), and the reader reads each file Excel saved as Excel read it. Known differences are pinned in
 * `PARTS_KNOWN` with their lines and reasons.
 */
import { describe, expect, it } from 'vitest'
import { checkSets, partsSets } from './sheet-parts-excel.js'

const SETS = partsSets('parts/')

/** Excel takes the part of an address after `#` as its place (Address and SubAddress); the book keeps the address as written. */
const anchor = (cell: string) => `S link ${cell}: Excel {"url":"https://example.com/page","location":"top"}, book {"url":"https://example.com/page#top"}`

/** `<set>:<case id>` → the difference lines pinned between Excel and the engine, each with its reason in a comment. */
const PARTS_KNOWN: Readonly<Record<string, string[]>> = {
  'files:written': [anchor('D13')],
  'files:after-insert': [anchor('D14')],
  'files:after-move': [anchor('D13')],
  'files:after-copy': [anchor('D13')],
  // the writer cuts a tooltip to Excel's 255 characters: Excel does not load a file with a longer one
  'files:tip': [`S link A1: Excel ${JSON.stringify({ url: 'https://example.com/', tooltip: 't'.repeat(255) })}, book ${JSON.stringify({ url: 'https://example.com/', tooltip: 't'.repeat(300) })}`],
}

/** Excel saves a space of an address as `%20` and shows it as a space; the reader gives the address as the file holds it. */
const space = (cell: string) => `S link ${cell}: Excel {"url":"https://example.com/a b/ж"}, book {"url":"https://example.com/a%20b/ж"}`

/** `<set>:<case id>` → the difference lines pinned between Excel and the reader of the file Excel saved. */
const PARTS_READ_KNOWN: Readonly<Record<string, string[]>> = {
  'files:written': [space('D6')],
  'files:after-insert': [space('D7')],
  'files:after-move': [space('D6')],
  'files:after-copy': [space('D6')],
}

describe('notes, hyperlinks and frozen panes against Excel for Mac', () => {
  it('every set is answered', () => expect(SETS.map((s) => s.name)).toEqual(['files', 'frozen-held', 'struct-held']))
  it('every pinned difference names a case', () => {
    const keys = new Set(SETS.flatMap((s) => s.cases.map((c) => `${s.name}:${c.id}`)))
    expect([...Object.keys(PARTS_KNOWN), ...Object.keys(PARTS_READ_KNOWN)].filter((k) => !keys.has(k))).toEqual([])
  })
})

checkSets(SETS, PARTS_KNOWN, PARTS_READ_KNOWN, (name) => import(`../bench/coverage/parts/${name}.ts`))
