/**
 * Дымовой тест собранного пакета (`npm run build` перед ним): все подпути `exports` есть в `dist` с
 * `.d.ts`, чистый node импортирует их по имени пакета, хэш синтетики и сессия совпадают с исходниками,
 * внутрь `dist` мимо `exports` не попасть.
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

describe('собранный пакет', () => {
  it('подпути exports: JS и .d.ts в dist', () => {
    const entries = Object.entries(pkg.exports).filter(([k]) => k !== './package.json')
    expect(entries.map(([k]) => k)).toEqual(['.', './contract', './synth', './internal'])
    for (const [, target] of entries) {
      if (typeof target === 'string') throw new Error('ожидается { types, import }')
      for (const file of [target.types, target.import]) {
        expect(file.startsWith('./dist/'), file).toBe(true)
        expect(existsSync(join(root, file)), `${file}: нет, сначала npm run build`).toBe(true)
      }
    }
    expect(pkg.files).toContain('dist')
  })

  it('чистый node импортирует пакет по имени, значения как у исходников', () => {
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
