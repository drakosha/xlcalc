/**
 * The `./xlsx/keep` subpath outside Node, as `test/xlsx-bundle.test.ts` checks `./xlsx`: bundled by esbuild for the
 * browser it takes no Node built-ins and none of the engine's computation, stays under its budget (the README's table
 * of what to import gives its size), and runs in an empty `node:vm` context given only the web platform's codecs.
 */
import { build } from 'esbuild'
import { fileURLToPath } from 'node:url'
import { createContext, runInContext } from 'node:vm'
import { describe, expect, it } from 'vitest'
import { libUrl } from './fixtures/target.js'

async function bundle(minify = false) {
  const lib = (f: string) => JSON.stringify(fileURLToPath(libUrl(f)))
  const r = await build({
    stdin: { contents: `export * from ${lib('xlsx-keep.js')}\nexport { writeXlsx } from ${lib('xlsx.js')}`, loader: 'js', resolveDir: fileURLToPath(new URL('..', import.meta.url)) },
    bundle: true,
    minify,
    platform: 'browser',
    format: 'iife',
    globalName: 'Keep',
    target: 'es2020',
    write: false,
    metafile: true,
    logLevel: 'silent',
  })
  const used = Object.entries(Object.values(r.metafile.outputs)[0].inputs).flatMap(([file, v]) => (v.bytesInOutput > 0 ? [file] : []))
  return { code: r.outputFiles[0].text, bytes: r.outputFiles[0].contents.length, used }
}

describe('xlsx/keep outside Node', () => {
  it('takes no Node built-ins and no part of the engine that computes, in under 290 KB minified', async () => {
    const { code, used, bytes } = await bundle(true)
    expect(code).not.toMatch(/\brequire\s*\(|\bprocess\.|\bBuffer\s*[.(]|\b(typeof|instanceof)\s+Buffer\b|["']node:/)
    const engine = used.filter((f) => /\/(functions|exec|session|spill|arrays)\/|\/(engine|evaluate|compile|runtime|functions|session|result|graph)\.(ts|js)$/.test(f) && !/\/exec\/budget\.(ts|js)$/.test(f))
    expect(engine).toEqual([])
    expect(bytes).toBeLessThan(290 * 1024)
  }, 60_000)

  it('runs in an empty context given the web codecs: a file opened, a value changed, written over it', async () => {
    const { code } = await bundle()
    const ctx = createContext({ CompressionStream, DecompressionStream, TextEncoder, TextDecoder })
    runInContext(code, ctx)
    const out = (await runInContext(
      `(async () => {
        const { data } = await Keep.writeXlsx({ book: { sheets: [{ name: 'S', cells: { A1: { value: 2 }, A2: { value: 'x' } } }] } })
        const r = await Keep.openKept(data)
        const same = await Keep.writeKept(data, { book: r.book, result: r.cached })
        r.book.sheets[0].cells.A1 = { value: 3 }
        const w = await Keep.writeKept(data, { book: r.book, result: r.cached })
        const back = await Keep.openKept(w.data)
        return JSON.stringify([same.data === data, back.book.sheets[0].cells])
      })()`,
      ctx,
    )) as string
    expect(JSON.parse(out)).toEqual([true, { A1: { value: 3 }, A2: { value: 'x' } }])
  }, 60_000)
})
