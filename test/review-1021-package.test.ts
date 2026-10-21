/**
 * The package as a host's bundler takes it (`"sideEffects": false`): each root call that checks a formula's calls,
 * bundled alone by esbuild for the browser from the build under test and run in an empty `node:vm` context, gives what
 * the package gives, by the functions' declared arguments and without the engine.
 */
import { build } from 'esbuild'
import { fileURLToPath } from 'node:url'
import { createContext, runInContext } from 'node:vm'
import { describe, expect, it } from 'vitest'
import { entryText, needsApostrophe, parseEntry, validateFormula, validateLocalFormula } from '../src/index.js'
import { libUrl } from './fixtures/target.js'

/** `names` of the root bundled alone, as a script defining `X`, and the library's modules it took code from. */
async function alone(names: string): Promise<{ code: string; used: string[] }> {
  const r = await build({ stdin: { contents: `export { ${names} } from ${JSON.stringify(fileURLToPath(libUrl('index.js')))}`, loader: 'ts', resolveDir: fileURLToPath(new URL('..', import.meta.url)) }, bundle: true, minify: true, format: 'iife', globalName: 'X', platform: 'browser', target: 'es2020', write: false, metafile: true, logLevel: 'silent' })
  const used = Object.entries(Object.values(r.metafile.outputs)[0].inputs).flatMap(([f, v]) => (v.bytesInOutput > 0 ? [f.replace(/^(?:.*\/)?(src|dist)\//, '').replace(/\.(ts|js)$/, '')] : []))
  return { code: r.outputFiles[0].text, used }
}

const run = (code: string, js: string): unknown => {
  const ctx = createContext({})
  runInContext(code, ctx)
  return JSON.parse(runInContext(`JSON.stringify(${js})`, ctx) as string)
}

describe('root calls that check formulas, bundled without the engine', () => {
  it.each([
    ['validateFormula', `X.validateFormula('SUM(1,2)')`, () => validateFormula('SUM(1,2)')],
    ['validateFormula', `X.validateFormula('FOO(1)')`, () => validateFormula('FOO(1)')],
    ['validateLocalFormula', `X.validateLocalFormula('=ROUND(1.5)', 'en-US')`, () => validateLocalFormula('=ROUND(1.5)', 'en-US')],
    ['parseEntry', `X.parseEntry('=ROUND(1.5)')`, () => parseEntry('=ROUND(1.5)')],
    ['needsApostrophe', `X.needsApostrophe('=SUM(1)')`, () => needsApostrophe('=SUM(1)')],
    ['entryText', `X.entryText('=SUM(1)')`, () => entryText('=SUM(1)')],
  ])('%s: %s gives the package\'s result', async (name, js, expected) => {
    const { code, used } = await alone(name)
    expect(used).not.toContain('functions')
    expect(run(code, js)).toEqual(JSON.parse(JSON.stringify(expected())))
  }, 60_000)
})
