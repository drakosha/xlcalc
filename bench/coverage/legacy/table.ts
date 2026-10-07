/**
 * Writes `src/legacy/params.ts`, the class of each whitelisted function's arguments in a formula of the old kind, from
 * Excel's own text of the `params` and `nesting` sets (`npx tsx bench/coverage/legacy/table.ts`):
 * - `v`: one value; a range there gets `@`, an array too where one value is wanted; an array evaluation goes on into it;
 * - `V`: the same, but an array evaluation stops there (IF's condition, VLOOKUP's key);
 * - `w`: one value, an array evaluation going on into a range there but not into an expression (VLOOKUP's column);
 * - `r`: a reference or an array; a range stays, an expression is computed by values (`@` on its ranges);
 * - `R`: the same, an array evaluation stops there (CHOOSE's values);
 * - `a`: an array; nothing in it gets `@`.
 * A position the sets did not probe (a size argument, an optional one the typical call leaves out) takes the class
 * of the function's other arguments of the same kind, else the usual class of its kind.
 */
import { readFileSync, writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { DEFS } from '../../../src/functions.js'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const load = (set: string) => (JSON.parse(readFileSync(`${HERE}../../../test/excel/legacy/${set}.json`, 'utf8')) as { formulas: Record<string, string> }).formulas
const top = load('params')
const nested = load('nesting')

const ranged = (f?: string) => (f === undefined ? undefined : /@D!B1:B3/.test(f))
const inner = (f?: string) => (f === undefined ? undefined : /\(@D!B1:B3\+0\)/.test(f))

/** The usual class of an argument kind (`args.ts`). */
const BY_KIND: Record<string, string> = { n: 'v', i: 'v', w: 'v', b: 'v', s: 'v', d: 'v', t: 'v', a: 'r', x: 'r', v: 'r', N: 'r', A: 'r', Y: 'r', Z: 'r', D: 'r', V: 'r', G: 'a', M: 'a', R: 'r', r: 'r', p: 'R', P: 'R' }

function classes(fn: string): string {
  const def = DEFS[fn]
  const kinds = def.args.split(' ').filter(Boolean).map((t) => t[0])
  const rep = def.rep ?? 0
  const kindAt = (i: number) => (i < kinds.length ? kinds[i] : kinds[kinds.length - rep + ((i - kinds.length + rep) % rep)])
  const probed: (string | undefined)[] = []
  for (let i = 1; top[`${fn}#r${i}`] !== undefined || top[`${fn}#e${i}`] !== undefined || i <= kinds.length; i++) {
    const r = ranged(top[`${fn}#r${i}`])
    const e = inner(top[`${fn}#e${i}`])
    if (r === undefined) probed.push(undefined)
    else if (r) probed.push(ranged(nested[`${fn}#array-r${i}`]) ? 'V' : inner(nested[`${fn}#array-e${i}`]) ? 'w' : 'v')
    else if (e === undefined) probed.push('r')
    else probed.push(e ? (inner(nested[`${fn}#array-e${i}`]) ? 'R' : 'r') : 'a')
  }
  // the declared kinds and, for repeated ones, every position probed
  const n = Math.max(kinds.length, probed.length)
  return Array.from({ length: n }, (_, i) => {
    if (probed[i]) return probed[i]
    const same = probed.find((c, j) => c && kindAt(j) === kindAt(i))
    return same ?? BY_KIND[kindAt(i)]
  }).join('')
}

const entries = Object.keys(DEFS)
  .sort()
  .map((fn) => `${/^[A-Z][A-Z0-9_]*$/.test(fn) ? fn : `'${fn}'`}: '${classes(fn)}'`)
const lines: string[] = []
for (let i = 0; i < entries.length; i += 6) lines.push(`  ${entries.slice(i, i + 6).join(', ')},`)
writeFileSync(
  `${HERE}../../../src/legacy/params.ts`,
  `/**
 * The class of each argument of the whitelisted functions in a formula of the old kind, as Excel 365 reads it (see
 * \`bench/coverage/legacy/table.ts\`, which writes this file from Excel's answers; do not edit): \`v\` one value, \`V\` one
 * value where an array evaluation stops, \`w\` one value where it goes on into a range alone, \`r\` a reference or an
 * array, \`R\` the same where an array evaluation stops, \`a\` an array. Repeated arguments take the class of their
 * place in the last group.
 */
export const PARAMS: Readonly<Record<string, string>> = {
${lines.join('\n')}
}
`,
)
console.log(`${entries.length} functions`)
