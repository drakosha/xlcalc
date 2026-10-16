/**
 * The `./entry` subpath takes none of the engine: bundled by esbuild for the browser from the build under test, it holds
 * the reading of typed text, the parser and the check of a formula's calls by the functions' declared arguments
 * (`function-arity.ts`), and the formatting of a value for editing, but no function's implementation, no compiler and
 * no computation; it stays under 110 KB, and runs in an empty `node:vm` context with ru-RU given as an object.
 */
import { createContext, runInContext } from 'node:vm'
import { describe, expect, it } from 'vitest'
import { browserBundle } from './fixtures/browser-bundle.js'
import { ruRU } from '../src/locales/ru-RU.js'

const bundle = () => browserBundle('typed-entry.js', 'ENTRY')

describe('typed entry without the engine', () => {
  it('the bundle holds no function, compiler or computation, under 110 KB', async () => {
    const { code, used } = await bundle()
    // functions/common: the helpers of numbers the General format rounds by, no function's implementation
    expect(used.filter((m) => /^(functions|exec|compile|engine|evaluate|runtime|graph|session)\b/.test(m) && m !== 'functions/common')).toEqual([])
    expect(used).toContain('function-arity')
    expect(code.length).toBeLessThan(110 * 1024)
  }, 60_000)

  it('runs in an empty context, ru-RU given as an object', async () => {
    const { code } = await bundle()
    const ctx = createContext({ ruRU })
    runInContext(code, ctx)
    const out = (js: string) => JSON.parse(runInContext(`JSON.stringify(${js})`, ctx) as string)
    expect(out(`ENTRY.parseEntry('1 234,5', { locale: ruRU }).input`)).toEqual({ value: 1234.5, numberFormat: 4 })
    expect(out(`ENTRY.parseEntry('=СУММ(A1;2', { locale: ruRU }).input`)).toEqual({ formula: 'SUM(A1,2)' })
    expect(out(`ENTRY.parseEntry('=ОКРУГЛ(1,5)', { locale: ruRU }).message`)).toBe('Функция ОКРУГЛ принимает 2 аргумента, а указано 1')
    expect(out(`[ENTRY.entryText(0.125, { numberFormat: '0%', locale: ruRU }), ENTRY.entryText(46096, { numberFormat: 14, locale: ruRU }), ENTRY.needsApostrophe('007'), ENTRY.needsApostrophe('abc')]`)).toEqual(['12,5%', '15.03.2026', true, false])
  }, 60_000)
})
