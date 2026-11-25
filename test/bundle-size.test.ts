/**
 * What a host's bundle takes of the package, bundled by esbuild for the browser and minified: `compute` alone stays
 * under 416 KB, as no locale module but en-US comes with it, nor the functions' declared arguments; the root,
 * `xlcalc/xlsx` and `xlcalc/csv` take no other locale either (`loadLocale`, which imports the others, is `xlcalc/locales`'s alone), and the ru-RU module takes nothing of the computation. A budget, so that the computation does not grow past it unnoticed; the README's table of
 * what to import gives the sizes. The light entry `xlcalc/core` takes none of the groups of functions, its `compute`
 * under 345 KB.
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
  it('keeps the XLS converter in its own entry point', async () => {
    const xls = await minified(`export * from ${lib('xls.js')}`)
    // The reader with its saved OLAP cache projection is about 386,300 minified bytes.
    expect(xls.bytes).toBeLessThan(378 * 1024)
    const print = await minified(`export * from ${lib('xls-print.js')}`)
    const isPrint = (file: string) => /\/(?:xls\/print-markup\.|xlsx\/print\/)/.test(file)
    expect(print.bytes).toBeLessThan(390 * 1024)
    expect(print.used.filter(isPrint)).toEqual(expect.arrayContaining([
      expect.stringMatching(/\/xls\/print-markup\./), expect.stringMatching(/\/xlsx\/print\/read\./),
    ]))
    expect(xls.used.filter(isPrint)).toEqual([])
    const connections = await minified(`export * from ${lib('xls-connections.js')}`)
    const isConnectionDetails = (file: string) => /\/xls\/connection-details\./.test(file)
    expect(connections.bytes).toBeLessThan(390 * 1024)
    expect(connections.used.some(isConnectionDetails)).toBe(true)
    expect(xls.used.some(isConnectionDetails) || print.used.some(isConnectionDetails)).toBe(false)
    for (const entry of ['index.js', 'xlsx.js', 'xlsx-keep.js', 'xlsx-print.js', 'xlsx-print-keep.js', 'xlsb.js']) {
      const { used } = await minified(`export * from ${lib(entry)}`)
      expect(used.filter((file) => /\/xls(?:\/|\.(?:js|ts)$)/.test(file))).toEqual([])
    }
  }, 60_000)

  it('print codecs and pagination belong to the opt-in xlsx/print entry, under 400 KB, and with keep under 464 KB', async () => {
    const isPrint = (file: string) => /\/xlsx(?:\/print\/|-print(?:-keep)?\.(?:ts|js)$)/.test(file)
    const { bytes, used } = await minified('export * from ' + lib('xlsx-print.js'))
    // 399,545 bytes of dist at 0.90.0: the xlsx reader and writer with print settings, pagination, notes and drawings
    expect(bytes).toBeLessThan(400 * 1024)
    expect(used.filter(isPrint)).toEqual(expect.arrayContaining([
      expect.stringMatching(/\/print\/layout\./), expect.stringMatching(/\/print\/read\./), expect.stringMatching(/\/print\/write\./),
    ]))
    expect(used.filter((file) => /\/xlsx\/keep\/book\./.test(file))).toEqual([])
    const kept = await minified('export * from ' + lib('xlsx-print-keep.js'))
    // 465,657 bytes of dist at 0.90.0: xlcalc/xlsx/keep with the print settings' reader and writer
    expect(kept.bytes).toBeLessThan(464 * 1024)
    expect(kept.used.filter(isPrint)).toEqual(expect.arrayContaining([expect.stringMatching(/\/print\/keep\./)]))
    for (const entry of ['index.js', 'core.js', 'xlsx.js', 'xlsx-keep.js', 'xlsb.js']) {
      const { used: inputs } = await minified('export * from ' + lib(entry))
      expect(inputs.filter(isPrint), entry).toEqual([])
    }
  }, 60_000)

  it('ChartML and chart source resolution are opt-in, including when using XLSX print, and under 52 KB', async () => {
    const isChart = (file: string) => /\/xlsx(?:\/charts\/|-charts\.(?:ts|js)$)/.test(file)
    const charts = await minified('export * from ' + lib('xlsx-charts.js'))
    // 48,831 bytes of dist at 0.90.0: the ChartML reader with the archive and drawing anchors it reads through
    expect(charts.bytes).toBeLessThan(52 * 1024)
    expect(charts.used.filter(isChart)).toEqual(expect.arrayContaining([
      expect.stringMatching(/\/charts\/read\./), expect.stringMatching(/\/charts\/values\./),
    ]))
    expect(charts.used.filter((file) => /\/(?:engine|session|evaluate|compile)\.(?:ts|js)$/.test(file))).toEqual([])
    for (const entry of ['index.js', 'core.js', 'xlsx.js', 'xlsx-keep.js', 'xlsx-print.js', 'xlsx-print-keep.js', 'xlsb.js']) {
      const { used } = await minified('export * from ' + lib(entry))
      expect(used.filter(isChart), entry).toEqual([])
    }
  }, 60_000)

  it('XLSB has its own budget and its converters stay outside the other entry points', async () => {
    const xlsb = await minified(`export * from ${lib('xlsb.js')}`)
    // 331,153 bytes of dist with sheet default widths in the shared reader; the room over it lets the xlsx reader grow by
    // a release or two before the converter is made smaller.
    expect(xlsb.bytes).toBeLessThan(336 * 1024)
    for (const entry of ['index.js', 'xlsx.js', 'xlsx-keep.js']) {
      const { used } = await minified(`export * from ${lib(entry)}`)
      expect(used.filter((file) => /\/xlsb(?:\/|\.(?:js|ts)$)/.test(file))).toEqual([])
    }
  }, 60_000)

  it('compute alone: under 416 KB minified, with en-US alone of the locales, the functions\' arguments by their specs', async () => {
    const { bytes, used } = await minified(`export { compute } from ${lib('index.js')}`)
    expect(locales(used)).toEqual([expect.stringMatching(/^en-US\.(ts|js)$/)])
    // the table of declared arguments is for the calls that check formulas without the engine (`declared-arity.ts`)
    expect(used.filter((f) => /function-arity/.test(f))).toEqual([])
    // CSE selector contexts: 413,845 bytes of dist, 410,601 at 0.89.0; ordinary generated formula code is unchanged.
    // Arrays of references (SUBTOTAL, SUMIF over OFFSET over an array): 415,109 bytes, 414,310 before.
    // Arrays of references into RANK, ROW, CELL and the database functions, #VALUE! for AGGREGATE and INDEX: 415,543 bytes, 415,376 before.
    // IRR by Excel's secant search with restarts: 416,103 bytes, 415,543 before.
    // Normal geometry and regular Liberation advances: 420,990 bytes.
    // Patterned fills checked on input and table styles' number formats in a cell's text: 422,087 bytes, 421,419 before.
    expect(bytes).toBeLessThan(416 * 1024)
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

  it("the light entry takes none of the groups' modules, its compute under 345 KB; the groups take the core's", async () => {
    // the modules of the groups alone (`src/groups/`): each is in the bundle of xlcalc/functions
    const GROUPS = [/\/groups\/(index|all|engineering|finance|lambda|statistics|xml)\./, /\/functions\/(engineering|xml)\//, /\/functions\/text\/filterxml\./, /\/functions\/financial\/(bond|coupon|odd|discount|depreciation)\./, /\/functions\/statistical\/(distributions|inverse|tests|fit|special|old-distributions)\./, /\/functions\/(jstat|lambda)\./, /\/functions\/lookup\/groupby\./]
    const grouped = (used: string[]) => used.filter((f) => GROUPS.some((g) => g.test(f)))
    const groups = (await minified(`export * from ${lib('groups/index.js')}`)).used
    expect(GROUPS.filter((g) => !groups.some((f) => g.test(f)))).toEqual([])
    const compute = await minified(`export { compute } from ${lib('core.js')}`)
    expect(grouped(compute.used)).toEqual([])
    expect(grouped((await minified(`export * from ${lib('core.js')}`)).used)).toEqual([])
    // CSE selector contexts: 340,986 bytes of dist, 337,717 at 0.89.0. Arrays of references: 342,250, 341,451 before.
    // IRR by Excel's secant search with restarts: 343,244.
    // Normal geometry and regular Liberation advances: 348,131 bytes.
    // Patterned fills and table styles' number formats: 349,228 bytes, 348,560 before.
    // Text order copying a text only when it holds a sharp s, 26 % faster sorting: 352,296 bytes, 352,254 before.
    expect(compute.bytes).toBeLessThan(345 * 1024)
  }, 60_000)
})
