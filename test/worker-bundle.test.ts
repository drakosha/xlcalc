/**
 * The threads' code stays out of the root entry: bundled by esbuild for the browser, `xlcalc` takes no file of
 * `worker/`, while it takes the computation in portions. The subpath `xlcalc/worker` and the thread's module bundle
 * for the browser too: no Node module is imported, and they reach Node's `worker_threads` (and the thread its
 * `crypto`, for the hashes) only through `process.getBuiltinModule` when they run in Node.
 */
import { build } from 'esbuild'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import { libUrl } from './fixtures/target.js'

async function bundle(file: string) {
  const r = await build({ entryPoints: [fileURLToPath(libUrl(file))], bundle: true, platform: 'browser', format: 'esm', target: 'es2020', write: false, metafile: true, logLevel: 'silent' })
  const used = Object.entries(Object.values(r.metafile.outputs)[0].inputs).flatMap(([f, v]) => (v.bytesInOutput > 0 ? [f] : []))
  return { code: r.outputFiles[0].text, used }
}

describe('bundles', () => {
  it('the root entry takes nothing of the threads, and takes the computation in portions', async () => {
    const { used } = await bundle('index.js')
    expect(used.filter((f) => /\/worker\//.test(f))).toEqual([])
    expect(used.some((f) => /\/async\/drive\.(ts|js)$/.test(f))).toBe(true)
  }, 60_000)

  it('the subpath and the thread bundle for the browser without Node modules', async () => {
    for (const [file, builtins] of [['worker/index.js', ['"node:worker_threads"']], ['worker/thread.js', ['"node:crypto"', '"node:worker_threads"']]] as const) {
      const { code, used } = await bundle(file)
      expect(used.filter((f) => !/^(src|dist)\//.test(f)), file).toEqual([])
      expect(code, file).not.toMatch(/\brequire\s*\(|from\s*["']node:|import\(\s*["']node:/)
      expect(code.match(/["']node:[a-z_]+["']/g) ?? [], file).toEqual(builtins)
    }
  }, 60_000)
})
