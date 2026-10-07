/**
 * The xlsx subpath's table, written from data that live elsewhere: `npx tsx bench/coverage/xlsx/generate.ts`
 * writes `src/xlsx/formulas/prefixes.ts` (the `_xlfn.` prefixes: the functions the measurement found prefixed,
 * `signatures.ts` and `sets/arrays/common.ts` with `_xlfn._xlws.` for SORT and FILTER, and every function
 * Microsoft lists as newer than Excel 2007, `excel-functions.json`). `test/xlsx-generated.test.ts` checks that it
 * is current.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { xl } from '../sets/arrays/common.js'
import { XLFN } from '../signatures.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
export const PREFIXES = fileURLToPath(new URL('../../../src/xlsx/formulas/prefixes.ts', import.meta.url))

/** Function names by the prefix a file writes before them. */
function prefixes(): { xlfn: string[]; xlws: string[] } {
  const listed = (JSON.parse(readFileSync(`${HERE}../excel-functions.json`, 'utf8')) as { name: string; since: string | null }[]).filter((f) => f.since).map((f) => f.name)
  const names = [...new Set([...Object.keys(XLFN), ...listed])].sort()
  // the harness's own spelling where it has one (`_xlfn._xlws.FILTER`), `_xlfn.` otherwise
  const spelled = names.map((n) => [n, xl(`${n}(`).slice(0, -1)] as const)
  return { xlfn: spelled.filter(([, s]) => !s.startsWith('_xlfn._xlws.')).map(([n]) => n), xlws: spelled.filter(([, s]) => s.startsWith('_xlfn._xlws.')).map(([n]) => n) }
}

/** Items followed by commas, wrapped at about 120 characters. */
function wrap(items: string[]): string {
  const lines: string[] = []
  for (const x of items.map((n) => `${n},`)) {
    if (!lines.length || lines.at(-1)!.length + x.length > 118) lines.push(`  ${x}`)
    else lines[lines.length - 1] += ` ${x}`
  }
  return `\n${lines.join('\n')}\n`
}

const quoted = (xs: string[]) => wrap(xs.map((x) => `'${x}'`))

export function prefixesSource(): string {
  const { xlfn, xlws } = prefixes()
  return `/**
 * Functions newer than Excel 2007, which a file writes with a prefix: \`_xlfn.XLOOKUP\`, \`_xlfn._xlws.FILTER\`. Written by
 * \`npx tsx bench/coverage/xlsx/generate.ts\` from the harness's data (the prefixes the measurement found, the functions
 * Microsoft lists as newer than 2007); do not edit.
 */
export const XLFN: ReadonlySet<string> = new Set([${quoted(xlfn)}])

/** Written \`_xlfn._xlws.\`. */
export const XLWS: ReadonlySet<string> = new Set([${quoted(xlws)}])
`
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  writeFileSync(PREFIXES, prefixesSource())
  console.log(`wrote ${PREFIXES}`)
}
