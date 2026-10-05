/**
 * Smoke test of the built package (`npm run build` first): every `exports` subpath exists in `dist` with
 * `.d.ts`, plain node imports them by package name, the synthetic hash and the session match the sources,
 * and `dist` cannot be reached around `exports`.
 */
import { execFileSync } from 'node:child_process'
import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { compute, openSession } from '../src/index.js'
import { classAEdits, synthBook } from '../src/synth.js'

const root = fileURLToPath(new URL('..', import.meta.url))
const pkg = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')) as {
  exports: Record<string, string | { types: string; import: string }>
  files: string[]
}
const child = fileURLToPath(new URL('./fixtures/dist-child.mjs', import.meta.url))

describe('built package', () => {
  it('exports subpaths: JS and .d.ts in dist', () => {
    const entries = Object.entries(pkg.exports).filter(([k]) => k !== './package.json')
    expect(entries.map(([k]) => k)).toEqual(['.', './contract', './synth', './internal'])
    for (const [, target] of entries) {
      if (typeof target === 'string') throw new Error('expected { types, import }')
      for (const file of [target.types, target.import]) {
        expect(file.startsWith('./dist/'), file).toBe(true)
        expect(existsSync(join(root, file)), `${file}: missing, run npm run build first`).toBe(true)
      }
    }
    expect(pkg.files).toContain('dist')
  })

  it('plain node imports the package by name, values match the sources', () => {
    const out = JSON.parse(execFileSync(process.execPath, [child], { cwd: root, encoding: 'utf8' }).trim())
    const clock = { now: '2026-10-04T12:00:00Z', seed: 7 }
    const synth = synthBook({ seed: 1, rows: 340 })
    const session = openSession(synth.book, { clock })
    for (const e of classAEdits(synth, 3, 10)) session.setCells([e])
    expect(out.hash).toBe(compute(synth.book, { clock }).hash)
    expect(out.edited).toBe(session.result().hash)
    session.close()
    expect(out.functions).toBe(17)
    expect(out.engine).toEqual(['compute', 'openSession', 'validateFormula', 'validateName'])
    expect(out.validate).toBe(true)
    expect(out.contract).toEqual(['MAX_FORMULA_DEPTH', 'MAX_FORMULA_LENGTH'])
    expect(out.internal).toEqual(['function', 'function', 'function'])
    expect(out.deep).toBe('ERR_PACKAGE_PATH_NOT_EXPORTED')
  }, 60_000)
})
