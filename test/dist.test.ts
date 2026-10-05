/**
 * Smoke test of the built package (`npm run build` first): every `exports` subpath exists in `dist` with
 * `.d.ts`, plain node imports them by package name, the synthetic hash and the session match the sources,
 * `dist` cannot be reached around `exports`, and the tarball holds only `dist` and the documents.
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
  exports: Record<string, string | { types: string; default: string }>
  files: string[]
}
const child = fileURLToPath(new URL('./fixtures/dist-child.mjs', import.meta.url))

describe('built package', () => {
  it('exports subpaths: JS and .d.ts in dist', () => {
    const entries = Object.entries(pkg.exports).filter(([k]) => k !== './package.json')
    expect(entries.map(([k]) => k)).toEqual(['.', './contract', './unstable/internal', './unstable/synth'])
    for (const [, target] of entries) {
      if (typeof target === 'string') throw new Error('expected { types, default }')
      for (const file of [target.types, target.default]) {
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
    expect(out.inputError).toBe('function')
    expect(out.validate).toBe(true)
    expect(out.contract).toEqual(['MAX_FORMULA_DEPTH', 'MAX_FORMULA_LENGTH'])
    expect(out.internal).toEqual(['function', 'function', 'function'])
    expect(out.closed).toEqual(Array(3).fill('ERR_PACKAGE_PATH_NOT_EXPORTED'))
  }, 60_000)

  it('the tarball holds dist, the documents and package.json only', () => {
    const [packed] = JSON.parse(execFileSync('npm', ['pack', '--dry-run', '--json', '--ignore-scripts'], { cwd: root, encoding: 'utf8' })) as { files: { path: string }[] }[]
    const files = packed.files.map((f) => f.path).sort()
    expect(files.filter((f) => !f.startsWith('dist/'))).toEqual(['CHANGELOG.md', 'LICENSE', 'README.md', 'THIRD_PARTY.md', 'package.json'])
    expect(files.filter((f) => f.startsWith('dist/') && !/\.(js|d\.ts)$/.test(f))).toEqual([])
    expect(files).toContain('dist/index.js')
    expect(files).toContain('dist/index.d.ts')
  }, 60_000)
})
