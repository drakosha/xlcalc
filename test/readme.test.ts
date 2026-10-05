/**
 * The README's code runs: every ```js block runs as an ES module that imports the built package by name
 * (`npm test` builds first), and each `// → text` comment must equal what the corresponding console.log
 * printed, in order; comment lines starting with `//   ` right after it continue a multi-line output.
 */
import { execFileSync } from 'node:child_process'
import { readFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

const root = fileURLToPath(new URL('..', import.meta.url))
const blocks = [...readFileSync(new URL('../README.md', import.meta.url), 'utf8').matchAll(/```js\n([\s\S]*?)```/g)].map((m) => m[1])

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

  for (const [i, code] of blocks.entries()) {
    it(`block ${i + 1}: ${code.split('\n').find((l) => l.startsWith('import'))}`, () => {
      expect(run(code)).toEqual(expected(code))
    }, 30_000)
  }
})
