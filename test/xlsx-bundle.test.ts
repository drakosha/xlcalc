/**
 * The `./xlsx` subpath outside Node: bundled by esbuild for the browser it takes no Node built-ins and none of the
 * engine's computation (functions, the evaluator, sessions), and it runs in an empty `node:vm` context given only
 * the web platform's codecs (deflate streams, UTF-8), writing a book and reading it back.
 */
import { build } from 'esbuild'
import { fileURLToPath } from 'node:url'
import { createContext, runInContext } from 'node:vm'
import { describe, expect, it } from 'vitest'
import { libUrl } from './fixtures/target.js'

/** The subpath, and with `locale` the root's `registerLocale` and the ru-RU module beside it, as a host of ru-RU books takes them. */
async function bundle(locale = false) {
  const lib = (f: string) => JSON.stringify(fileURLToPath(libUrl(f)))
  const contents = [`export * from ${lib('xlsx.js')}`, ...(locale ? [`export { registerLocale } from ${lib('index.js')}`, `export { ruRU } from ${lib('locales/ru-RU.js')}`] : [])].join('\n')
  const r = await build({
    stdin: { contents, loader: 'js', resolveDir: fileURLToPath(new URL('..', import.meta.url)) },
    bundle: true,
    platform: 'browser',
    format: 'iife',
    globalName: 'Xlsx',
    target: 'es2020',
    write: false,
    metafile: true,
    logLevel: 'silent',
  })
  const used = Object.entries(Object.values(r.metafile.outputs)[0].inputs).flatMap(([file, v]) => (v.bytesInOutput > 0 ? [file] : []))
  return { code: r.outputFiles[0].text, used }
}

describe('xlsx outside Node', () => {
  it('the bundle takes no Node built-ins and no part of the engine that computes', async () => {
    const { code, used } = await bundle()
    // Buffer in use, not named in a message (readXlsx takes a Uint8Array, an ArrayBuffer or a Buffer)
    expect(code).not.toMatch(/\brequire\s*\(|\bprocess\.|\bBuffer\s*[.(]|\b(typeof|instanceof)\s+Buffer\b|["']node:/)
    // the budget's check (`spend`, a few bytes) comes with copyRange, which pastes shared formulas: a session's copy stops at its deadline
    const engine = used.filter((f) => /\/(functions|exec|session|spill|arrays)\/|\/(engine|evaluate|compile|runtime|functions|session|result|graph)\.(ts|js)$/.test(f) && !/\/exec\/budget\.(ts|js)$/.test(f))
    expect(engine).toEqual([])
  }, 60_000)

  it('runs in an empty context given the web codecs: writes a book and reads it back', async () => {
    const { code, used } = await bundle(true)
    expect(used.filter((f) => /\/(functions|session|spill|arrays)\/|\/(engine|evaluate|compile|runtime|functions|session|result|graph)\.(ts|js)$/.test(f))).toEqual([])
    const ctx = createContext({ CompressionStream, DecompressionStream, TextEncoder, TextDecoder })
    expect(runInContext('[typeof process, typeof require, typeof Buffer]', ctx)).toEqual(['undefined', 'undefined', 'undefined'])
    runInContext(code, ctx)
    const out = (await runInContext(
      `(async () => {
        Xlsx.registerLocale(Xlsx.ruRU)
        const book = { sheets: [{ name: 'Лист', cells: { A1: { value: 2 }, B1: { formula: 'A1*2' } } }], locale: 'ru-RU' }
        const { data } = await Xlsx.writeXlsx(book)
        const r = await Xlsx.readXlsx(data)
        return JSON.stringify(r.book)
      })()`,
      ctx,
    )) as string
    expect(JSON.parse(out)).toEqual({ sheets: [{ name: 'Лист', cells: { A1: { value: 2 }, B1: { formula: 'A1*2' } } }], locale: 'ru-RU' })
  }, 60_000)

  it('without the codecs a clear error', async () => {
    const { code } = await bundle()
    const ctx = createContext({})
    runInContext(code, ctx)
    await expect(runInContext(`Xlsx.writeXlsx({ sheets: [{ name: 'S', cells: {} }] })`, ctx)).rejects.toThrow(/no TextEncoder|no CompressionStream/)
  }, 60_000)
})
