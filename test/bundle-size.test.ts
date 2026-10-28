/**
 * What a host's bundle takes of the package, bundled by esbuild for the browser and minified: `compute` alone stays
 * under 325 KB, as no locale module but en-US comes with it, nor the functions' declared arguments; the root takes no
 * other locale either, and the ru-RU module takes nothing of the computation. A budget, so that the computation does not grow past it unnoticed; the README's table of
 * what to import gives the sizes.
 */
import { build } from 'esbuild'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { libUrl } from './fixtures/target.js'

/** The minified bundle of `contents` (imports of the build under test) and the files it took code from. */
async function minified(contents: string): Promise<{ bytes: number; used: string[] }> {
  const r = await build({ stdin: { contents, loader: 'ts', resolveDir: fileURLToPath(new URL('..', import.meta.url)) }, bundle: true, minify: true, format: 'esm', platform: 'browser', target: 'es2020', write: false, metafile: true, logLevel: 'silent' })
  const used = Object.entries(Object.values(r.metafile.outputs)[0].inputs).flatMap(([f, v]) => (v.bytesInOutput > 0 ? [f] : []))
  return { bytes: r.outputFiles[0].contents.length, used }
}

const lib = (file: string) => JSON.stringify(fileURLToPath(libUrl(file)))
const locales = (used: string[]) => used.filter((f) => /\/locales\//.test(f)).map((f) => f.replace(/^.*\/locales\//, ''))

describe('bundle size', () => {
  it('compute alone: under 325 KB minified, with en-US alone of the locales, the functions\' arguments by their specs', async () => {
    const { bytes, used } = await minified(`export { compute } from ${lib('index.js')}`)
    expect(locales(used)).toEqual([expect.stringMatching(/^en-US\.(ts|js)$/)])
    // the table of declared arguments is for the calls that check formulas without the engine (`declared-arity.ts`)
    expect(used.filter((f) => /function-arity/.test(f))).toEqual([])
    expect(bytes).toBeLessThan(325 * 1024)
  }, 60_000)

  it('the root takes en-US alone of the locales; the ru-RU module none of the computation', async () => {
    const root = (await minified(`export * from ${lib('index.js')}`)).used
    expect(locales(root)).toEqual([expect.stringMatching(/^en-US\.(ts|js)$/)])
    const ru = await minified(`export * from ${lib('locales/ru-RU.js')}`)
    expect(ru.used.map((f) => f.replace(/^(.*\/)?(src|dist)\//, '')).sort()).toEqual([expect.stringMatching(/^locale\/lcid\.(ts|js)$/), expect.stringMatching(/^locales\/ru-RU\.(ts|js)$/)])
  }, 60_000)
})
