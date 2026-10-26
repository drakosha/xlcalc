/**
 * The `./text` subpath takes none of the engine: bundled by esbuild for the browser from the build under test, it holds
 * the tokenizer, the translation of a locale's formula text, the check of its functions' arguments and the category of a
 * number format's code (no parser, formatting or computation), en-US's messages alone of the locales' data, stays under
 * 55 KB, and runs in an empty `node:vm`
 * context with ru-RU given as an object, or the part of it a formula's text needs. It holds no `xlcalc`, which a host
 * looks for in its bundle as the sign of the engine: a tag not loaded asks for the locale as an object, and only the root
 * names the module to import and `registerLocale`.
 */
import { createContext, runInContext } from 'node:vm'
import { describe, expect, it } from 'vitest'
import { browserBundle } from './fixtures/browser-bundle.js'
import { ruRU } from '../src/locales/ru-RU.js'

const bundle = () => browserBundle('text.js', 'TEXT')

describe('text without the engine', () => {
  it('the bundle holds the modules of formula text alone, under 55 KB', async () => {
    const { code, used } = await bundle()
    // wildcards: the lexer takes the escape of a regular expression's characters from it;
    // numfmt/builtin, category and parse: a number format's category, by the built-in codes and the split into sections;
    // decimals, general, glyphs and functions/common (its rounding): Increase and Decrease Decimal, by what General shows in a column
    expect(used).toEqual(['a1', 'call-at', 'collate', 'column-width', 'entry/clipboard', 'entry/close', 'error-type', 'find-text', 'function-categories', 'functions/common', 'input', 'input-error', 'input/basic', 'input/styles', 'layout/spans', 'locale/english', 'locale/index', 'locale/lex', 'locale/load', 'locale/outline-messages', 'locale/registry', 'locale/translate', 'locales/en-US', 'model', 'numfmt/builtin', 'numfmt/category', 'numfmt/decimals', 'numfmt/general', 'numfmt/glyphs', 'numfmt/parse', 'point', 'scan', 'sheet-pattern', 'sort-order', 'strict-args', 'syntax/ast', 'syntax/errors', 'syntax/parse', 'syntax/scan', 'syntax/table', 'syntax/tokens', 'text', 'wildcards'])
    // 46.1 KB with nameAt, Find, the sort order, the categories and the column's characters (16.10); 51.8 with stepDecimals
    expect(code.length).toBeLessThan(55 * 1024)
    expect(code).not.toContain('xlcalc')
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
    expect(out(`[TEXT.numberFormatCategory('0.000'), TEXT.numberFormatCategory('0.000', ruRU), TEXT.numberFormatChoices(ruRU)[2]]`)).toEqual([
      { category: 'number', name: 'Number' },
      { category: 'number', name: 'Числовой' },
      { key: 'currency', name: 'Денежный', format: '#,##0.00\\ "₽"', code: '#,##0.00\\ "₽"' },
    ])
    expect(out(`(() => { try { TEXT.scanReferences('=A1', { locale: 'ru-RU' }) } catch (e) { return e.message } })()`)).toBe(
      "opts.locale: the locale ru-RU is not loaded: pass the locale as an object: ruRU from the package's module locales/ru-RU (loaded: en-US)",
    )
  }, 60_000)
})
