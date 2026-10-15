/**
 * `functionSignature`: every whitelisted function has one, and it agrees with the function's measured arguments (how
 * many a call needs, how many it takes, whether they repeat); ru-RU has a word for every English one but single letters;
 * no tip names two arguments alike; the tip's text, a name in the locale, and the input checked.
 */
import { describe, expect, it } from 'vitest'
import { FUNCTION_CATEGORIES, FUNCTION_NAMES, functionSignature, InputError, registerLocale } from '../src/index.js'
import { FUNCTIONS } from '../src/functions.js'
import { SIGNATURES } from '../src/signatures.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { functionDescription } from '../src/function-text/index.js'
import { ruRUFunctionText } from '../src/function-text/ru-RU.js'

registerLocale(ruRU)

/** ru-RU with its words of the tips joined, as a host takes them. */
const ru = { ...ruRU, ...ruRUFunctionText }

/** The English words of the signatures' names, numbers taken off. */
const WORDS = new Set(Object.values(SIGNATURES).flatMap((s) => s.split(' ').flatMap((t) => (t && t !== '...' ? [t.replace(/^\[|\]$/g, '').replace(/\d+$/, '')] : []))))

describe('functionSignature', () => {
  it('has a signature for every whitelisted function and none else', () => {
    expect(Object.keys(SIGNATURES).sort()).toEqual([...FUNCTION_NAMES])
  })

  it('agrees with the arguments a call needs and takes', () => {
    for (const name of FUNCTION_NAMES) {
      const sig = functionSignature(name)!
      const required = sig.args.filter((a) => !a.optional).length
      // LET is syntax: a name, its value and the calculation at least, then pairs
      const spec = FUNCTIONS[name] ?? { min: 3, max: 253 }
      expect(required, name).toBe(spec.min)
      if (sig.repeat) expect(spec.max, name).toBeGreaterThan(sig.args.length)
      else expect(sig.args.length, name).toBe(spec.max)
      // the repeated arguments come twice, the second time optional
      if (sig.repeat) expect(sig.args.slice(-sig.repeat).every((a) => a.optional), name).toBe(true)
    }
  })

  it('names no two arguments of a tip alike, in en-US and ru-RU', () => {
    for (const name of FUNCTION_NAMES)
      for (const locale of ['en-US', ru]) {
        const names = functionSignature(name, locale)!.args.map((a) => a.name)
        expect(new Set(names).size, `${typeof locale === 'string' ? locale : 'ru'} ${name}: ${names.join(' ')}`).toBe(names.length)
      }
  })

  it('has ru-RU words for every English word but single letters', () => {
    expect([...WORDS].filter((w) => !(w in ruRUFunctionText.argumentWords) && !/^[a-z]$/.test(w))).toEqual([])
    expect(Object.keys(ruRUFunctionText.argumentWords).filter((w) => !WORDS.has(w))).toEqual([])
  })

  it('gives every function a category of the window and a line in en-US and ru-RU', () => {
    expect([...FUNCTION_CATEGORIES]).toEqual(['financial', 'date', 'math', 'statistical', 'lookup', 'database', 'text', 'logical', 'information', 'engineering', 'compatibility'])
    for (const name of FUNCTION_NAMES) {
      expect(FUNCTION_CATEGORIES, name).toContain(functionSignature(name)!.category)
      for (const locale of [undefined, ru]) expect(functionDescription(name, locale), `${locale ? 'ru' : 'en'} ${name}`).toMatch(/^\S.*\.$/)
    }
    // ru-RU's module names every category and describes every function in Russian
    const own = ruRUFunctionText.functionText
    expect(Object.keys(own.categories).sort()).toEqual([...FUNCTION_CATEGORIES].sort())
    expect(Object.keys(own.descriptions).sort()).toEqual([...FUNCTION_NAMES])
    expect(FUNCTION_NAMES.filter((n) => !/[а-яё]/i.test(own.descriptions[n]))).toEqual([])
    expect(functionSignature('NORMDIST')).toMatchObject({ fn: 'NORMDIST', category: 'compatibility', categoryName: 'Compatibility' })
    expect(functionSignature('СЧЁТЕСЛИ', ru)).toMatchObject({ fn: 'COUNTIF', category: 'statistical', categoryName: 'Статистические' })
    expect(functionDescription('countif', ru)).toBe('Сколько ячеек отвечают условию.')
    expect(functionDescription('VLOOKUP')).toBe('A value from a column of a table, found in its first column.')
    expect(functionDescription('CELL')).toBeNull()
    // the ru-RU locale alone, as a host that only computes takes it, has none of the tips' words
    expect(ruRU.argumentWords).toBeUndefined()
    expect(functionSignature('ВПР', ruRU)!.text).toBe('ВПР(lookup_value;table;column_index;[approximate])')
  })

  it('writes the tip as Excel lists it, with the locale list separator', () => {
    expect(functionSignature('SUM')).toEqual({
      name: 'SUM',
      fn: 'SUM',
      args: [
        { name: 'number1', optional: false },
        { name: 'number2', optional: true },
      ],
      repeat: 1,
      text: 'SUM(number1, [number2], ...)',
      category: 'math',
      categoryName: 'Math & Trig',
    })
    expect(functionSignature('sumifs')!.text).toBe('SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)')
    expect(functionSignature('ВПР', ru)!.text).toBe('ВПР(искомое;таблица;номер_столбца;[приближённо])')
    expect(functionSignature('СуммЕслиМн', ru)).toMatchObject({ name: 'СУММЕСЛИМН', repeat: 2, text: 'СУММЕСЛИМН(диапазон_суммирования;диапазон_условия1;условие1;[диапазон_условия2;условие2];...)' })
    expect(functionSignature('SUM', ru)!.text).toBe('СУММ(число1;[число2];...)')
    // a function ru-RU names in English
    expect(functionSignature('ABS', ru)!.text).toBe('ABS(число)')
    expect(functionSignature('TODAY')!.text).toBe('TODAY()')
    expect(functionSignature('ROW')!.text).toBe('ROW([reference])')
  })

  it('answers null for a name that is no function of the whitelist', () => {
    for (const name of ['', 'SUMM', 'CELL', 'СУММ']) expect(functionSignature(name), name).toBeNull()
  })

  it('checks its input', () => {
    expect(() => functionSignature(1 as unknown as string)).toThrow(InputError)
    expect(() => functionSignature('SUM', 'de-XX')).toThrow(InputError)
  })
})
