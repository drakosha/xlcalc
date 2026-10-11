/** The texts of error values and the sets of them each reader takes (`src/syntax/errors.ts`), from one table. */
import { describe, expect, it } from 'vitest'
import { toInvariant } from '../src/index.js'
import { CELL_ERRORS, ERROR_CODE, ERROR_TEXT, FORMULA_ERRORS, LOCALE_ERRORS, TYPED_ERRORS } from '../src/syntax/errors.js'
import { EN_US } from '../src/locales/en-US.js'

describe('error texts', () => {
  it('codes by text are the table turned round', () => {
    for (const [code, text] of Object.entries(ERROR_TEXT)) expect(ERROR_CODE[text]).toBe(code)
    expect(Object.keys(ERROR_CODE)).toHaveLength(Object.keys(ERROR_TEXT).length)
  })

  it('each set holds the one before it: typed, read by a formula, held by a file’s cell, translated by a locale', () => {
    expect(TYPED_ERRORS).toEqual(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'])
    expect(FORMULA_ERRORS).toEqual([...TYPED_ERRORS, '#SPILL!'])
    expect(CELL_ERRORS).toEqual([...FORMULA_ERRORS, '#CALC!'])
    expect(CELL_ERRORS.every((e) => LOCALE_ERRORS.includes(e))).toBe(true)
  })

  it('a locale’s check lists the translated ones in their order', () => {
    const locale = { ...EN_US, tag: 'xx-XX', errors: { '#CYCLE!': '#ЦИКЛ!' } }
    expect(() => toInvariant('1', locale)).toThrow(`unknown key "#CYCLE!" (expected ${LOCALE_ERRORS.join(', ')})`)
    expect(LOCALE_ERRORS.join(' ')).toBe('#NULL! #DIV/0! #VALUE! #REF! #NAME? #NUM! #N/A #GETTING_DATA #SPILL! #CALC! #FIELD! #BLOCKED! #CONNECT! #BUSY! #UNKNOWN! #PYTHON!')
  })
})
