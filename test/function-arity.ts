/**
 * `src/function-arity.ts`, each function's declared arguments without its implementation, made from the engine's
 * definitions (`DEFS`): `test/function-arity.test.ts` checks that the file is the one this makes; `npx tsx
 * test/function-arity.ts --write` writes it again after a function's arguments change.
 */
import { writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { DEFS } from '../src/functions.js'

/** A value as a TypeScript literal in the repository's style: `{ args: 'n', rep: 1 }`, `[0, 2]`. */
const literal = (v: unknown): string =>
  Array.isArray(v) ? `[${v.map(literal).join(', ')}]` : typeof v === 'object' && v ? `{ ${Object.entries(v).map(([k, x]) => `${k}: ${literal(x)}`).join(', ')} }` : typeof v === 'string' ? `'${v}'` : String(v)

export function functionArity(): string {
  const rows = Object.keys(DEFS)
    .sort()
    .map((name) => {
      const { args, rep, max, notOmitted } = DEFS[name]
      const def = { args, ...(rep !== undefined ? { rep } : {}), ...(max !== undefined ? { max } : {}), ...(notOmitted ? { notOmitted } : {}) }
      return `  ${/^[A-Z][A-Z0-9_]*$/.test(name) ? name : `'${name}'`}: ${literal(def)},`
    })
  return `/**
 * Each function's declared arguments (\`ArityDef\`), without its implementation, for the check of a formula's calls where
 * the engine is not loaded (\`xlcalc/entry\`, \`useArities\`). Made from the engine's definitions by
 * \`npx tsx test/function-arity.ts --write\`; \`test/function-arity.test.ts\` keeps it the same as them.
 */
import type { ArityDef } from './arity.js'

export const FUNCTION_ARITY: Readonly<Record<string, ArityDef>> = {
${rows.join('\n')}
}
`
}

if (process.argv.includes('--write')) writeFileSync(resolve(import.meta.dirname, '../src/function-arity.ts'), functionArity())
