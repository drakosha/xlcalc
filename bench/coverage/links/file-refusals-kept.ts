/**
 * What Excel kept of the sets file-refusals: a file it opened may have lost a rule, a name or a font without a word, and neither
 * `Workbooks.Open` nor the repair window tells. `npx tsx bench/coverage/links/file-refusals-kept.ts <set> [win/]` looks for each
 * case's `holds` in the parts of the file Excel saved, in `test/excel/[win/]links/<set>/`, and writes the verdicts to
 * `test/excel/[win/]file-refusals-kept/<set>.json`. The saved files themselves stay out of the repository.
 */
import { existsSync, readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { openArchive } from '../../../src/xlsx/zip/read.js'
import { writeAnswers } from '../stored.js'
import type { LinksCases } from './cases.js'

/** What the cases predicted refused and opened after all are about, so they have no `holds`. */
const SURVIVORS: Record<string, string> = {
  'cf-3d': 'S:T',
  'cf-external': '[1]',
  'x14-3d': 'S:T',
  'dv-3d': 'S:T',
  'count-cf-8193': 'duplicateValues',
  'count-dv-32768': 'type="custom"',
  'count-select-8193': '<selection',
  'ext-name-bad': 'x y',
  'ext-name-ref': 'garbage',
  'link-url-2085': 'https://e.com/?qqq',
  'link-url-3000': 'https://e.com/?qqq',
  'note-author-65536': '<author>aaa',
  'name-backslash-a': 'name="\\a"',
  'dxf-one-code-207': 'cfRule',
  'mixed-same-207': 'cfRule',
  'slicer-taken': 'Slicer_Region',
  'slicer-items': 'Slicer_Region',
  'text-cf-8193': 'type="expression"',
  'text-dv-8193': 'type="custom"',
}

const [set, root = ''] = process.argv.slice(2)
const at = (path: string) => fileURLToPath(new URL(`../../../test/excel/${root}${path}`, import.meta.url))
const { buildLinksCases } = (await import(`./${set}.js`)) as { buildLinksCases(): LinksCases }

const kept: Record<string, boolean> = {}
for (const c of buildLinksCases().cases) {
  const file = at(`links/${set}/${c.id}.xlsx`)
  const holds = c.holds ?? SURVIVORS[c.id]
  if (!holds || !existsSync(file)) continue
  const archive = openArchive(new Uint8Array(readFileSync(file)))
  const texts = await Promise.all(archive.names.filter((n) => n.endsWith('.xml') || n.endsWith('.rels')).map((n) => archive.text(n)))
  kept[c.id] = texts.some((t) => t!.includes(holds))
}
writeAnswers(at(`file-refusals-kept/${set}.json`), kept)
console.log(`${set}: ${Object.values(kept).filter(Boolean).length} of ${Object.keys(kept).length} kept, lost: ${Object.keys(kept).filter((k) => !kept[k]).join(' ')}`)
