/**
 * `src/function-arity.ts`, each function's declared arguments without its implementation, `src/groups/manifest.ts`,
 * those of the groups' functions, and `src/function-words.ts`, the functions' names and argument counts alone, made from the engine's
 * definitions (`DEFS`, every group loaded): `test/function-arity.test.ts` checks that the files are the ones this makes;
 * `npx tsx test/function-arity.ts --write` writes them again after the functions change.
 */
import { writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { arityOf } from '../src/arity.js'
import { DEFS } from '../src/functions.js'
import { allFunctions } from '../src/groups/index.js'
import { loadGroups } from '../src/groups/registry.js'

loadGroups(allFunctions)

/** A value as a TypeScript literal in the repository's style: `{ args: 'n', rep: 1 }`, `[0, 2]`. */
const literal = (v: unknown): string =>
  Array.isArray(v) ? `[${v.map(literal).join(', ')}]` : typeof v === 'object' && v ? `{ ${Object.entries(v).map(([k, x]) => `${k}: ${literal(x)}`).join(', ')} }` : typeof v === 'string' ? `'${v}'` : String(v)

/** The rows of the declared arguments of `names`, sorted, indented by `indent`. */
function rows(names: readonly string[], indent: string): string[] {
  return [...names].sort().map((name) => {
    const { args, rep, max, notOmitted } = DEFS[name]
    const def = { args, ...(rep !== undefined ? { rep } : {}), ...(max !== undefined ? { max } : {}), ...(notOmitted ? { notOmitted } : {}) }
    return `${indent}${/^[A-Z][A-Z0-9_]*$/.test(name) ? name : `'${name}'`}: ${literal(def)},`
  })
}

const grouped = new Set(allFunctions.flatMap((g) => g.functions))

export function functionArity(): string {
  const groups = allFunctions.map((g) => `  ...GROUP_ARITY.${g.name},`)
  return `/**
 * Each function's declared arguments (\`ArityDef\`), without its implementation, for the check of a formula's calls where
 * the engine is not loaded (\`xlcalc/entry\`, \`useArities\`): the core's, then the groups' (\`groups/manifest.ts\`). Made
 * from the engine's definitions by \`npx tsx test/function-arity.ts --write\`; \`test/function-arity.test.ts\` keeps it the
 * same as them.
 */
import type { ArityDef } from './arity.js'
import { GROUP_ARITY } from './groups/manifest.js'

export const FUNCTION_ARITY: Readonly<Record<string, ArityDef>> = {
${rows(Object.keys(DEFS).filter((n) => !grouped.has(n)), '  ').join('\n')}
${groups.join('\n')}
}
`
}

export function groupManifest(): string {
  const groups = allFunctions.map((g) => `  ${g.name}: {\n${rows(g.functions, '    ').join('\n')}\n  },`)
  return `/**
 * The declared arguments (\`ArityDef\`) of each group's functions, without their implementations: what the light entry
 * checks a call of a group not loaded by (\`registry.ts\`, \`useManifest\`). Made from the engine's definitions by
 * \`npx tsx test/function-arity.ts --write\`; \`test/function-arity.test.ts\` keeps it the same as them.
 */
import type { Manifest } from './registry.js'

export const GROUP_ARITY: Manifest = {
${groups.join('\n')}
}
`
}

export function functionWords(): string {
  const groups = new Map<string, string[]>()
  for (const name of Object.keys(DEFS).sort()) {
    const { min, max, takes } = arityOf(DEFS[name])
    const next = Array.from({ length: 256 - min }, (_, i) => min + i + 1).find((k) => k <= max && takes(k))
    const key = `${min} ${max} ${next === undefined ? 1 : next - min}`
    groups.set(key, [...(groups.get(key) ?? []), name])
  }
  const row = ([key, names]: [string, string[]]) => {
    const lines: string[] = []
    let line = key
    for (const name of names) {
      if (line.length + name.length + 1 > 116) {
        lines.push(line)
        line = ''
      }
      line += (line ? ' ' : '') + name
    }
    return [...lines, line].join('\n  ')
  }
  return `/**
 * The whitelisted functions' names and how many arguments each takes, without the rest of their arguments: the xlsx writers tell a
 * function's name from a defined name's by them without the engine (\`syntax/check.ts\`, \`bindCalls\`), and what Excel refuses in a
 * file (\`xlsx/write/excel-loads.ts\`). Each row of the table starts with the least, the most and the step between them. Made from
 * the engine's definitions by \`npx tsx test/function-arity.ts --write\`; \`test/function-arity.test.ts\` keeps it the same as them.
 */
const parse = (table: string) => {
  const out = new Map<string, readonly [min: number, max: number, step: number]>()
  for (const row of table.split('|')) {
    const [min, max, step, ...names] = row.trim().split(/\\s+/)
    for (const name of names) out.set(name, [Number(min), Number(max), Number(step)])
  }
  return out
}

export const FUNCTION_COUNTS: ReadonlyMap<string, readonly [min: number, max: number, step: number]> = /* @__PURE__ */ parse(
  \`${[...groups].sort(([a], [b]) => (a < b ? -1 : 1)).map(row).join('\n  |')}\`,
)
`
}

if (process.argv.includes('--write')) {
  writeFileSync(resolve(import.meta.dirname, '../src/function-arity.ts'), functionArity())
  writeFileSync(resolve(import.meta.dirname, '../src/function-words.ts'), functionWords())
  writeFileSync(resolve(import.meta.dirname, '../src/groups/manifest.ts'), groupManifest())
}
