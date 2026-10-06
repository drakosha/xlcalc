/** The library (and the synthetic book) bundled by esbuild for the browser: an IIFE that sets `FormulaEngine`, and the files it took in. */
import { build } from 'esbuild'
import { fileURLToPath } from 'node:url'
import { libUrl } from './target.js'

export async function bundle(): Promise<{ code: string; inputs: string[] }> {
  const r = await build({
    stdin: { contents: `export * from ${JSON.stringify(fileURLToPath(libUrl('index.js')))}\nexport { synthBook } from ${JSON.stringify(fileURLToPath(libUrl('synth.js')))}`, loader: 'ts', resolveDir: fileURLToPath(new URL('../..', import.meta.url)) },
    bundle: true,
    platform: 'browser',
    format: 'iife',
    globalName: 'FormulaEngine',
    target: 'es2020',
    write: false,
    metafile: true,
    logLevel: 'silent',
  })
  return { code: r.outputFiles[0].text, inputs: Object.keys(r.metafile.inputs) }
}
