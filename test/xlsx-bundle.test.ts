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

async function bundle() {
  const r = await build({
    entryPoints: [fileURLToPath(libUrl('xlsx.js'))],
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
    expect(code).not.toMatch(/\brequire\s*\(|\bprocess\.|\bBuffer\b|["']node:/)
    expect(used.filter((f) => /\/(functions|exec|session|spill|arrays)\/|\/(engine|evaluate|compile|runtime|functions|session|result|graph)\.(ts|js)$/.test(f))).toEqual([])
  }, 60_000)

  it('runs in an empty context given the web codecs: writes a book and reads it back', async () => {
    const { code } = await bundle()
    const ctx = createContext({ CompressionStream, DecompressionStream, TextEncoder, TextDecoder })
    expect(runInContext('[typeof process, typeof require, typeof Buffer]', ctx)).toEqual(['undefined', 'undefined', 'undefined'])
    runInContext(code, ctx)
    const out = (await runInContext(
      `(async () => {
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
