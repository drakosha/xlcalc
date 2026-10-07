/**
 * The xlsx subpath's tables, written from data that live elsewhere: `npx tsx bench/coverage/xlsx/generate.ts`
 * writes `src/xlsx/formulas/prefixes.ts` (the `_xlfn.` prefixes: the functions the measurement found prefixed,
 * `signatures.ts` and `sets/arrays/common.ts` with `_xlfn._xlws.` for SORT and FILTER, and every function
 * Microsoft lists as newer than Excel 2007, `excel-functions.json`) and `src/xlsx/formulas/modes.ts` (how each
 * whitelisted function takes its arguments, from the engine's definitions, so the reader need not load the
 * engine). `test/xlsx-generated.test.ts` checks that both are current.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { DEFS } from '../../../src/functions.js'
import { xl } from '../sets/arrays/common.js'
import { XLFN } from '../signatures.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
export const PREFIXES = fileURLToPath(new URL('../../../src/xlsx/formulas/prefixes.ts', import.meta.url))
export const MODES = fileURLToPath(new URL('../../../src/xlsx/formulas/modes.ts', import.meta.url))

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

/** An argument kind's mode: lifted (`l`), an array or reference kind lifted by arrays only (`a`), passed on (`p`), taken whole (`w`). */
const mode = (k: string) => (k === 'r' ? 'a' : k === 'p' || k === 'P' ? 'p' : 'NAYZVGMRD'.includes(k) ? 'w' : 'l')

/** A function's modes by declared argument, `/n` the repeated last n, `!` when it can answer an array or a range. */
export function modes(): Record<string, string> {
  const out: Record<string, string> = {}
  for (const [name, def] of Object.entries(DEFS).sort(([a], [b]) => (a < b ? -1 : 1))) {
    const kinds = def.args.split(' ').filter(Boolean).map((t) => t[0])
    // N and T take a range's top-left cell in both kinds of formula
    const m = kinds.map((k) => (def.ranges === 'first' && mode(k) === 'l' ? 'w' : mode(k))).join('')
    out[name] = `${m}${def.rep ? `/${def.rep}` : ''}${def.array ? '!' : ''}`
  }
  return out
}

export function modesSource(): string {
  return `/**
 * How each whitelisted function takes its arguments, by declared argument: \`l\` one value (a range of several cells
 * is lifted), \`a\` a reference or one value (an array is lifted), \`p\` passed on to the function's answer, \`w\` taken
 * whole; \`/n\` the last n repeat, \`!\` the function can answer an array or a range. Written by
 * \`npx tsx bench/coverage/xlsx/generate.ts\` from the engine's definitions, so the reader need not load them; do not edit.
 */
export const MODES: Readonly<Record<string, string>> = {${wrap(Object.entries(modes()).map(([n, m]) => `${/^[A-Z][A-Z0-9]*$/.test(n) ? n : `'${n}'`}: '${m}'`))}}
`
}

if (process.argv[1] === fileURLToPath(import.meta.url)) {
  writeFileSync(PREFIXES, prefixesSource())
  writeFileSync(MODES, modesSource())
  console.log(`wrote ${PREFIXES} and ${MODES}`)
}
