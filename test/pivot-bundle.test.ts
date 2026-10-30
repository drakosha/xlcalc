/**
 * `xlcalc/pivot` stays out of the root: `compute` and the root's whole bundle take none of the PivotTable engine but its
 * hook, and the subpath adds to a host's bundle of the root no more than its budget, minified for the browser.
 */
import { build } from 'esbuild'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { libUrl } from './fixtures/target.js'

async function minified(contents: string): Promise<{ bytes: number; used: string[] }> {
  const r = await build({ stdin: { contents, loader: 'ts', resolveDir: fileURLToPath(new URL('..', import.meta.url)) }, bundle: true, minify: true, format: 'esm', platform: 'browser', target: 'es2020', write: false, metafile: true, logLevel: 'silent' })
  const used = Object.entries(Object.values(r.metafile.outputs)[0].inputs).flatMap(([f, v]) => (v.bytesInOutput > 0 ? [f] : []))
  return { bytes: r.outputFiles[0].contents.length, used }
}

const lib = (file: string) => JSON.stringify(fileURLToPath(libUrl(file)))
/** The engine's modules: the calculation, its layout and refresh, all but the hook the root calls through. */
const engine = (used: string[]) => used.filter((f) => /\/pivot\/(?!hook\.)[^/]+\.(ts|js)$/.test(f))

describe('the bundle of xlcalc/pivot', () => {
  it('the root takes none of the engine, the subpath adds under its budget to the root', async () => {
    const root = await minified(`export * from ${lib('index.js')}`)
    expect(engine(root.used)).toEqual([])
    expect(engine((await minified(`export { compute } from ${lib('index.js')}`)).used)).toEqual([])
    const both = await minified(`export * from ${lib('index.js')}\nexport * from ${lib('pivot/index.js')}`)
    expect(both.bytes - root.bytes).toBeLessThan(BUDGET * 1024)
  }, 60_000)
})

/** The kilobytes the subpath may add to the root's bundle. */
const BUDGET = 120
