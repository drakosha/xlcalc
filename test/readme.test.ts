/**
 * The README's code runs, and the examples of docs/ that import the package and say what they print: every ```js block runs as an ES module that imports the built package by name
 * (`npm test` builds first), and each `// → text` comment must equal what the corresponding console.log
 * printed, in order; comment lines starting with `//   ` right after it continue a multi-line output.
 * The list of functions in docs/functions.md is the whitelist.
 */
import { execFileSync } from 'node:child_process'
import { readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { FUNCTION_NAMES } from '../src/index.js'

const root = fileURLToPath(new URL('..', import.meta.url))
const readme = readFileSync(new URL('../README.md', import.meta.url), 'utf8')
const blocks = [...readme.matchAll(/```js\n([\s\S]*?)```/g)].map((m) => m[1])
/** The examples of docs/ that run: ```js blocks that import the package and print what their `// →` comments say. */
const docBlocks = readdirSync(join(root, 'docs'))
  .filter((f) => f.endsWith('.md'))
  .sort()
  .flatMap((f) => [...readFileSync(join(root, 'docs', f), 'utf8').matchAll(/```js\n([\s\S]*?)```/g)].map((m): [string, string] => [f, m[1]]))
  .filter(([, b]) => /from 'xlcalc/.test(b) && b.includes('// → ') && b.includes('console.log('))

function expected(code: string): string[] {
  const out: string[] = []
  let open = false
  for (const line of code.split('\n')) {
    const arrow = line.indexOf('// → ')
    if (arrow >= 0) out.push(line.slice(arrow + 5))
    else if (open && line.startsWith('//   ')) out[out.length - 1] += '\n' + line.slice(5)
    open = arrow >= 0 || (open && line.startsWith('//   '))
  }
  return out
}

/** Runs a block with console.log captured; what each call printed. */
function run(code: string): string[] {
  const prelude = "import { format as __format } from 'node:util'\nconst __out = []\nconsole.log = (...a) => __out.push(__format(...a))\nprocess.on('exit', () => process.stdout.write(JSON.stringify(__out)))\n"
  return JSON.parse(execFileSync(process.execPath, ['--input-type=module'], { cwd: root, input: prelude + code, encoding: 'utf8' })) as string[]
}

describe('README', () => {
  it('has code blocks with expected output', () => {
    expect(blocks.length).toBeGreaterThanOrEqual(5)
    for (const b of blocks) expect(expected(b).length, b.slice(0, 80)).toBeGreaterThan(0)
  })

  it('lists the whitelist by category in docs/functions.md', () => {
    const functions = readFileSync(join(root, 'docs/functions.md'), 'utf8')
    const section = functions.slice(0, functions.indexOf('\nStatus:'))
    const listed = [...section.matchAll(/^- [^:]+: (.+)$/gm)].flatMap((m) => [...m[1].matchAll(/`([^`]+)`/g)].map((x) => x[1]))
    expect([...listed].sort()).toEqual(FUNCTION_NAMES)
  })

  for (const [i, [file, code]] of docBlocks.entries()) {
    it(`docs/${file} block ${i + 1}: ${code.split('\n').find((l) => l.startsWith('import'))}`, () => {
      expect(run(code)).toEqual(expected(code))
    }, 30_000)
  }

  for (const [i, code] of blocks.entries()) {
    it(`block ${i + 1}: ${code.split('\n').find((l) => l.startsWith('import'))}`, () => {
      expect(run(code)).toEqual(expected(code))
    }, 30_000)
  }
})
