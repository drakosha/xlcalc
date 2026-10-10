/**
 * Notes, hyperlinks and frozen panes as files hold them, both ways. Books written by `writeXlsx` with every kind of
 * them (and after structure operations in the engine): Excel must load each book and read in it what the book holds.
 * Books Excel gave its own: notes, hyperlinks and frozen panes made through its object model on a plain book. Every
 * saved file must read back as Excel read it. Predicted before Excel ran it (`files.predicted.json`).
 */
import type { BookInput, CellInput } from '../../../src/contract.js'
import { restructure } from '../../../src/index.js'
import type { PartsCase, PartsCases } from './cases.js'

/** Notes with an author, without one, of several lines and Cyrillic; links outside the book, into it, both, with tooltips; frozen panes of each kind. */
function writtenBook(): BookInput {
  const cells: Record<string, CellInput> = {
    A1: { value: 'Header', note: { author: 'Ann', text: 'A note' } },
    B2: { value: 5, note: { text: 'without an author' } },
    C3: { value: 'multi', note: { author: 'Борис', text: 'Line 1\nLine 2\n\nЛиния 4 & <tags> "quotes"' } },
    D4: { value: 'site', link: { url: 'https://example.com/' } },
    D5: { value: 'query', link: { url: 'https://example.com/path?q=1&r=2' } },
    D6: { value: 'space', link: { url: 'https://example.com/a b/ж' } },
    D7: { value: 'mail', link: { url: 'mailto:ann@example.com?subject=Hi', tooltip: 'Write to Ann' } },
    D8: { value: 'to T T', link: { location: "'T T'!B2" } },
    D9: { value: 'to Total', link: { location: 'Total' } },
    D10: { value: 'doc part', link: { url: 'https://example.com/doc', location: 'part2', tooltip: 'Part two' } },
    D11: { value: 'both', link: { location: 'S!A1' }, note: { author: 'Ann', text: 'a link and a note' } },
    D12: { value: 'file', link: { url: 'other.xlsx' } },
    E1: { link: { url: 'https://example.com/empty' } },
    E2: { note: { author: 'Ann', text: 'a note on an empty cell' } },
  }
  for (let r = 1; r <= 12; r++) cells[`A${r + 1}`] = { value: r }
  return {
    sheets: [
      { name: 'S', cells, frozen: { rows: 1, columns: 1 } },
      { name: 'T T', cells: { B2: { value: 'target', note: { author: 'Ann', text: 'on the second sheet' } } }, frozen: { rows: 2 } },
      { name: 'T2', cells: { A1: { value: 1 } }, frozen: { columns: 3 } },
    ],
    names: [{ name: 'Total', ref: "'S'!$B$2" }],
  }
}

/** A sheet of plain values and two more, for what Excel makes itself. */
function plainBook(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let r = 1; r <= 6; r++) Object.assign(cells, { [`A${r}`]: { value: r }, [`B${r}`]: { value: `b${r}` } })
  return { sheets: [{ name: 'S', cells }, { name: 'T', cells: { A1: { value: 1 } } }, { name: 'T2', cells: { A1: { value: 2 } } }] }
}

/** 1,100 notes on one sheet and three on another: the VML shape ids take two blocks of 1,024 on the first. */
function manyNotes(): BookInput {
  const cells: Record<string, CellInput> = {}
  for (let i = 0; i < 1100; i++) cells[`${String.fromCharCode(65 + (i % 10))}${Math.floor(i / 10) + 1}`] = { value: i, note: { author: i % 2 ? 'Ann' : 'Борис', text: `note ${i}` } }
  return { sheets: [{ name: 'S', cells }, { name: 'T', cells: { A1: { note: { text: 'a' } }, B2: { note: { text: 'b' } }, C3: { note: { text: 'c' } } } }] }
}

const after = (book: BookInput, op: Parameters<typeof restructure>[1]): BookInput => {
  const r = restructure(book, op)
  if (!r.ok) throw new Error(`files: ${r.message}`)
  return r.book
}

export function buildPartsCases(): PartsCases {
  const written = writtenBook()
  const books: Record<string, BookInput> = {
    written,
    plain: plainBook(),
    many: manyNotes(),
    // Excel's limits: a ScreenTip of 255 characters, an address of 2,079
    tip: { sheets: [{ name: 'S', cells: { A1: { value: 'x', link: { url: 'https://example.com/', tooltip: 't'.repeat(300) } } } }] },
    url: { sheets: [{ name: 'S', cells: { A1: { value: 'x', link: { url: `https://example.com/${'u'.repeat(2100)}` } } } }] },
    'after-insert': after(written, { op: 'insertRows', sheet: 'S', row: 1 }),
    'after-move': after(written, { op: 'moveRange', from: { sheet: 'S', range: 'C3:E5' }, to: { sheet: 'S', cell: 'G3' } }),
    'after-copy': after(written, { op: 'copyRange', from: { sheet: 'S', range: 'A1:D4' }, to: { sheet: 'T2', range: 'B2' } }),
  }
  const cases: PartsCase[] = [
    ...['written', 'many', 'tip', 'url', 'after-insert', 'after-move', 'after-copy'].map((book) => ({ id: book, book })),
    {
      id: 'made-notes',
      book: 'plain',
      make: [
        { note: 'B2', sheet: 'S', text: 'Made\nby Excel' },
        { note: 'C3', sheet: 'S', text: 'Второе примечание' },
        { note: 'A1', sheet: 'T', text: 'on T' },
      ],
    },
    {
      id: 'made-links',
      book: 'plain',
      make: [
        { link: 'B1', sheet: 'S', url: 'https://example.org/x' },
        { link: 'B2', sheet: 'S', location: 'T!B3' },
        { link: 'B3', sheet: 'S', url: 'https://example.org/doc', location: 'part', tooltip: 'Tip' },
        { link: 'B4', sheet: 'S', url: 'mailto:bob@example.org' },
        { link: 'C5', sheet: 'S', location: "'T2'!A1:B2", tooltip: 'an empty cell' },
      ],
    },
    {
      id: 'made-frozen',
      book: 'plain',
      make: [
        { freeze: 'S', rows: 2, columns: 1 },
        { freeze: 'T', rows: 1, columns: 0 },
        { freeze: 'T2', rows: 0, columns: 3 },
      ],
    },
  ]
  return { books, cases }
}
