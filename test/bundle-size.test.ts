/**
 * What a host's bundle takes of the package, bundled by esbuild for the browser and minified: `compute` alone stays
 * under 310 KB, as no locale module but en-US comes with it, nor the functions' declared arguments; the root takes no
 * other locale either, and the ru-RU module takes nothing of the computation. A budget, so that the computation does not grow past it unnoticed; the README's table of
 * what to import gives the sizes.
 */
import { build } from 'esbuild'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { libUrl } from './fixtures/target.js'

/**
 * The minified bundle of `contents` (imports of the build under test) and the files it took code from: its entry's
 * chunk, as a host's bundler splits it, a locale module that `loadLocale` imports going to a chunk of its own.
 */
async function minified(contents: string): Promise<{ bytes: number; used: string[] }> {
  const dir = mkdtempSync(join(tmpdir(), 'xlcalc-size-'))
  try {
    const entry = join(dir, 'entry.ts')
    writeFileSync(entry, contents)
    const r = await build({ entryPoints: [entry], bundle: true, splitting: true, outdir: join(dir, 'out'), minify: true, format: 'esm', platform: 'browser', target: 'es2020', write: false, metafile: true, logLevel: 'silent' })
    const [path, out] = Object.entries(r.metafile.outputs).find(([, o]) => o.entryPoint !== undefined)!
    const used = Object.entries(out.inputs).flatMap(([f, v]) => (v.bytesInOutput > 0 ? [f] : []))
    return { bytes: r.outputFiles.find((f) => f.path.endsWith(path.replace(/^.*\/out\//, '/out/')))!.contents.length, used }
  } finally {
    rmSync(dir, { recursive: true, force: true })
  }
}

const lib = (file: string) => JSON.stringify(fileURLToPath(libUrl(file)))
const locales = (used: string[]) => used.filter((f) => /\/locales\//.test(f)).map((f) => f.replace(/^.*\/locales\//, ''))

describe('bundle size', () => {
  it('compute alone: under 310 KB minified, with en-US alone of the locales, the functions\' arguments by their specs', async () => {
    const { bytes, used } = await minified(`export { compute } from ${lib('index.js')}`)
    expect(locales(used)).toEqual([expect.stringMatching(/^en-US\.(ts|js)$/)])
    // the table of declared arguments is for the calls that check formulas without the engine (`declared-arity.ts`)
    expect(used.filter((f) => /function-arity/.test(f))).toEqual([])
    expect(bytes).toBeLessThan(310 * 1024)
  }, 60_000)

  it('the root takes en-US alone of the locales; the ru-RU module none of the computation', async () => {
    const root = (await minified(`export * from ${lib('index.js')}`)).used
    expect(locales(root)).toEqual([expect.stringMatching(/^en-US\.(ts|js)$/)])
    const ru = await minified(`export * from ${lib('locales/ru-RU.js')}`)
    expect(ru.used.map((f) => f.replace(/^(.*\/)?(src|dist)\//, '')).sort()).toEqual([expect.stringMatching(/^locale\/lcid\.(ts|js)$/), expect.stringMatching(/^locales\/ru-RU\.(ts|js)$/)])
  }, 60_000)
})
