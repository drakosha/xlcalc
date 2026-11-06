/**
 * What a host's bundle takes of the package, bundled by esbuild for the browser and minified: `compute` alone stays
 * under 402 KB, as no locale module but en-US comes with it, nor the functions' declared arguments; the root,
 * `xlcalc/xlsx` and `xlcalc/csv` take no other locale either (`loadLocale`, which imports the others, is `xlcalc/locales`'s alone), and the ru-RU module takes nothing of the computation. A budget, so that the computation does not grow past it unnoticed; the README's table of
 * what to import gives the sizes. The light entry `xlcalc/core` takes none of the groups of functions, its `compute`
 * under 331 KB.
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
  it('compute alone: under 402 KB minified, with en-US alone of the locales, the functions\' arguments by their specs', async () => {
    const { bytes, used } = await minified(`export { compute } from ${lib('index.js')}`)
    expect(locales(used)).toEqual([expect.stringMatching(/^en-US\.(ts|js)$/)])
    // the table of declared arguments is for the calls that check formulas without the engine (`declared-arity.ts`)
    expect(used.filter((f) => /function-arity/.test(f))).toEqual([])
    expect(bytes).toBeLessThan(402 * 1024)
  }, 60_000)

  it('the root, xlcalc/xlsx and xlcalc/csv take en-US alone of the locales, xlcalc/csv and xlcalc/locales none of the computation, xlcalc/locales the modules it loads; the ru-RU module none of the computation, nor the registry its import loads it into', async () => {
    const root = (await minified(`export * from ${lib('index.js')}`)).used
    expect(locales(root)).toEqual([expect.stringMatching(/^en-US\.(ts|js)$/)])
    expect(locales((await minified(`export * from ${lib('xlsx.js')}`)).used)).toEqual([expect.stringMatching(/^en-US\.(ts|js)$/)])
    const loader = (await minified(`export * from ${lib('locales.js')}`)).used
    expect(locales(loader).sort()).toEqual([expect.stringMatching(/^en-US\.(ts|js)$/), expect.stringMatching(/^ru-RU\.(ts|js)$/)])
    expect(loader.filter((f) => /\/(functions|exec|session|spill|arrays|xlsx)\/|\/(engine|evaluate|compile|runtime|functions|session|result|graph)\.(ts|js)$/.test(f))).toEqual([])
    // CSV reads and writes without the computation: the entry's parser and the display of values (functions/common: the helpers of numbers the General format rounds by)
    const csv = (await minified(`export * from ${lib('csv.js')}`)).used
    expect(locales(csv)).toEqual([expect.stringMatching(/^en-US\.(ts|js)$/)])
    expect(csv.filter((f) => !/functions\/common\.(ts|js)$/.test(f) && /\/(functions|exec|session|spill|arrays)\/|\/(engine|evaluate|compile|runtime|functions|session|result|graph)\.(ts|js)$/.test(f))).toEqual([])
    const ru = await minified(`export * from ${lib('locales/ru-RU.js')}`)
    expect(ru.used.map((f) => f.replace(/^(.*\/)?(src|dist)\//, '')).sort()).toEqual([expect.stringMatching(/^locale\/imported\.(ts|js)$/), expect.stringMatching(/^locale\/lcid\.(ts|js)$/), expect.stringMatching(/^locales\/ru-RU\.(ts|js)$/)])
  }, 60_000)

  it("the light entry takes none of the groups' modules, its compute under 331 KB; the groups take the core's", async () => {
    // the modules of the groups alone (`src/groups/`): each is in the bundle of xlcalc/functions
    const GROUPS = [/\/groups\/(index|all|engineering|finance|lambda|statistics|xml)\./, /\/functions\/(engineering|xml)\//, /\/functions\/text\/filterxml\./, /\/functions\/financial\/(bond|coupon|odd|discount|depreciation)\./, /\/functions\/statistical\/(distributions|inverse|tests|fit|special|old-distributions)\./, /\/functions\/(jstat|lambda)\./, /\/functions\/lookup\/groupby\./]
    const grouped = (used: string[]) => used.filter((f) => GROUPS.some((g) => g.test(f)))
    const groups = (await minified(`export * from ${lib('groups/index.js')}`)).used
    expect(GROUPS.filter((g) => !groups.some((f) => g.test(f)))).toEqual([])
    const compute = await minified(`export { compute } from ${lib('core.js')}`)
    expect(grouped(compute.used)).toEqual([])
    expect(grouped((await minified(`export * from ${lib('core.js')}`)).used)).toEqual([])
    expect(compute.bytes).toBeLessThan(331 * 1024)
  }, 60_000)
})
