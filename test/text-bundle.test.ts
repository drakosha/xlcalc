/**
 * The `./text` subpath takes none of the engine: bundled by esbuild for the browser from the build under test, it holds
 * the tokenizer, the translation of a locale's formula text and the check of its functions' arguments (no parser, number
 * format or computation), en-US's messages alone of the locales' data, stays under 50 KB, and runs in an empty `node:vm`
 * context with ru-RU given as an object, or the part of it a formula's text needs.
 */
import { build } from 'esbuild'
import { fileURLToPath } from 'node:url'
import { createContext, runInContext } from 'node:vm'
import { describe, expect, it } from 'vitest'
import { libUrl } from './fixtures/target.js'
import { ruRU } from '../src/locales/ru-RU.js'

async function bundle() {
  const r = await build({ entryPoints: [fileURLToPath(libUrl('text.js'))], bundle: true, platform: 'browser', format: 'iife', globalName: 'TEXT', target: 'es2020', write: false, metafile: true, minify: true, logLevel: 'silent' })
  const used = Object.entries(Object.values(r.metafile.outputs)[0].inputs).flatMap(([file, v]) => (v.bytesInOutput > 0 ? [file.replace(/^(?:.*\/)?(src|dist)\//, '').replace(/\.(ts|js)$/, '')] : []))
  return { code: r.outputFiles[0].text, used: used.sort() }
}

describe('text without the engine', () => {
  it('the bundle holds the modules of formula text alone, under 50 KB', async () => {
    const { code, used } = await bundle()
    // criteria, dates and numfmt/style: the lexer takes the escape of a regular expression's characters from criteria
    expect(used).toEqual(['a1', 'collate', 'criteria', 'dates', 'entry/clipboard', 'input', 'input-error', 'locale/english', 'locale/index', 'locale/lex', 'locale/load', 'locale/registry', 'locale/translate', 'locales/en-US', 'model', 'numfmt/style', 'scan', 'sheet-pattern', 'strict-args', 'syntax/ast', 'syntax/errors', 'syntax/scan', 'syntax/table', 'syntax/tokens', 'text'])
    expect(code.length).toBeLessThan(50 * 1024)
  }, 60_000)

  it('runs in an empty context, a locale given as an object or its formula text alone', async () => {
    const { code } = await bundle()
    const { tag, decimalSeparator, listSeparator, arrayColumnSeparator, arrayRowSeparator, booleans, errors, functions, tableItems } = ruRU
    const ctx = createContext({ ruRU, part: { tag, decimalSeparator, listSeparator, arrayColumnSeparator, arrayRowSeparator, booleans, errors, functions, tableItems } })
    runInContext(code, ctx)
    const out = (js: string) => JSON.parse(runInContext(`JSON.stringify(${js})`, ctx) as string)
    const brief = '.map((r) => r.at + "-" + r.end + " " + r.text)'
    expect(out(`TEXT.scanReferences('=СУММ(A1:B2;1.5)+Лист1!C3', { locale: ruRU })${brief}`)).toEqual(['6-11 A1:B2', '17-25 Лист1!C3'])
    expect(out(`TEXT.scanReferences('=СУММ(A1:B2;{1;2:3})', { locale: part })${brief}`)).toEqual(['6-11 A1:B2'])
    expect(out(`TEXT.scanReferences('=SUM(A1:B2,{1,2;3})+)C3')${brief}`)).toEqual(['5-10 A1:B2', '21-23 C3'])
    expect(out(`[TEXT.quoteSheetIfNeeded('Q1 data'), TEXT.quoteSheetIfNeeded('ИСТИНА', part), TEXT.escapeColumn('@x#y')]`)).toEqual(["'Q1 data'", "'ИСТИНА'", "'@x'#y"])
  }, 60_000)
})
