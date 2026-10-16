/** A subpath of the build under test bundled by esbuild for the browser, minified, as a script defining `globalName`: its code and the library's modules in it, without their extensions, in order. */
import { build } from 'esbuild'
import { fileURLToPath } from 'node:url'
import { libUrl } from './target.js'

export async function browserBundle(file: string, globalName: string): Promise<{ code: string; used: string[] }> {
  const r = await build({ entryPoints: [fileURLToPath(libUrl(file))], bundle: true, platform: 'browser', format: 'iife', globalName, target: 'es2020', write: false, metafile: true, minify: true, logLevel: 'silent' })
  const used = Object.entries(Object.values(r.metafile.outputs)[0].inputs).flatMap(([f, v]) => (v.bytesInOutput > 0 ? [f.replace(/^(?:.*\/)?(src|dist)\//, '').replace(/\.(ts|js)$/, '')] : []))
  return { code: r.outputFiles[0].text, used: used.sort() }
}
