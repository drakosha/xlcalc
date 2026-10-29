/**
 * The `./csv` subpath takes none of the engine: bundled by esbuild for the browser from the build under test, it runs in
 * an empty `node:vm` context, reads formulas by the functions' declared arguments and writes values as Excel shows them.
 */
import { createContext, runInContext } from 'node:vm'
import { describe, expect, it } from 'vitest'
import { browserBundle } from './fixtures/browser-bundle.js'
import { ruRU } from '../src/locales/ru-RU.js'

describe('CSV without the engine', () => {
  it('runs in an empty context, ru-RU given as an object', async () => {
    const { code, used } = await browserBundle('csv.js', 'CSV')
    expect(used.filter((m) => /^(functions|exec|compile|engine|evaluate|runtime|graph|session)\b/.test(m) && m !== 'functions/common')).toEqual([])
    const ctx = createContext({ ruRU, TextDecoder, TextEncoder })
    runInContext(code, ctx)
    const out = (js: string) => JSON.parse(runInContext(`JSON.stringify(${js})`, ctx) as string)
    expect(out(`CSV.readCsv('1 234,5;"=СУММ(1;2)";=ОКРУГЛ(1,5)', { locale: ruRU, formulas: true }).book.sheets[0].cells`)).toEqual({ A1: { value: 1234.5, numberFormat: 4 }, B1: { formula: 'SUM(1,2)' }, C1: { value: '=ОКРУГЛ(1,5)' } })
    expect(out(`CSV.writeCsv({ book: { sheets: [{ name: 'S', cells: { A1: { value: 46096, numberFormat: 14 }, B1: { value: 'a,b' } } }] } }, { bom: false }).text`)).toBe('3/15/2026,"a,b"')
  }, 60_000)
})
