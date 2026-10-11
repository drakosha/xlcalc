/**
 * The `./a1` subpath takes none of the engine: bundled by esbuild for the browser from the build under test, it holds
 * only the address modules and the check of its functions' arguments (no parser, locale, number format or computation),
 * stays small, and runs in an empty `node:vm` context.
 */
import { build } from 'esbuild'
import { fileURLToPath } from 'node:url'
import { createContext, runInContext } from 'node:vm'
import { describe, expect, it } from 'vitest'
import { libUrl } from './fixtures/target.js'

async function bundle() {
  const r = await build({ entryPoints: [fileURLToPath(libUrl('a1.js'))], bundle: true, platform: 'browser', format: 'iife', globalName: 'A1', target: 'es2020', write: false, metafile: true, minify: true, logLevel: 'silent' })
  const used = Object.entries(Object.values(r.metafile.outputs)[0].inputs).flatMap(([file, v]) => (v.bytesInOutput > 0 ? [file.replace(/^(?:.*\/)?(src|dist)\//, '')] : []))
  return { code: r.outputFiles[0].text, used }
}

describe('a1 without the engine', () => {
  it('the bundle holds the address modules alone, in a few kilobytes', async () => {
    const { code, used } = await bundle()
    expect(used.map((f) => f.replace(/\.(ts|js)$/, '')).sort()).toEqual(['a1', 'collate', 'input-error', 'layout/spans', 'model', 'sheet-pattern', 'strict-args', 'structure/area'])
    expect(code.length).toBeLessThan(8_000)
  }, 60_000)

  it('runs in an empty context', async () => {
    const { code } = await bundle()
    const ctx = createContext({})
    runInContext(code, ctx)
    expect(JSON.parse(runInContext(`JSON.stringify([A1.parseRange('$b$2:D9'), A1.quoteSheetIfNeeded('Q1 data'), A1.rangeAddress({ top: 1, left: 2, bottom: 1048576, right: 4 })])`, ctx) as string)).toEqual([
      { kind: 'cells', top: 2, left: 2, bottom: 9, right: 4 },
      "'Q1 data'",
      'B:D',
    ])
  }, 60_000)
})
