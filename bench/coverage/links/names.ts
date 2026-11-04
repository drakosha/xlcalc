/**
 * What Excel writes as a link's target for a file's name: the book beside this one, opened, a formula typed over it and the
 * book saved. On Windows alone (`bench/windows/links.ts`), which puts a linked book beside the opened one.
 *
 * Predicted before Excel ran the set, in the saved `xl/externalLinks/_rels/externalLink2.xml.rels`: a space `%20`; `%` and
 * `#`, which an address gives another meaning, `%25` and `%23`; Cyrillic and `&` as they are (`&amp;` in the XML).
 */
import type { LinksCase, LinksCases } from './cases.js'

const FILES: [string, string][] = [
  ['spaces', 'a b c.xlsx'],
  ['percent', '100% a.xlsx'],
  ['hash', 'a#b.xlsx'],
  ['cyrillic', 'Отчёт 2025.xlsx'],
  ['ampersand', 'P&L.xlsx'],
  ['escape-like', 'a%20b.xlsx'],
]

const over = ([id, file]: [string, string]): LinksCase => ({
  id,
  about: `a formula typed over ${file}, open beside this book`,
  sheets: [{ name: 'S', cells: {} }],
  linked: { S: { A1: '[1]Sheet1!A1' } },
  links: [{ target: 'Book2.xlsx', sheets: [{ name: 'Sheet1', cells: { A1: 1 } }] }],
  read: [{ sheet: 'S', range: 'A1:A2' }],
  ops: [
    { op: 'source', target: file, open: true, sheets: { Sheet1: { A1: 5 } } },
    { op: 'enter', sheet: 'S', cell: 'A2', formula: `='[${file}]Sheet1'!A1` },
    { op: 'dropSource', target: file },
  ],
  predicted: { 'S!A2': 5 },
})

export function buildLinksCases(): LinksCases {
  return { cases: FILES.map(over) }
}
