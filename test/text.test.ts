/**
 * `xlcalc/text`: the root's functions of formula text and of number formats' categories without the engine, the same
 * functions; `escapeColumn` writes a column's name in a structured reference as Excel writes it, which the engine reads
 * back as the name.
 */
import { describe, expect, it } from 'vitest'
import * as root from '../src/index.js'
import { escapeColumn, InputError, validateFormula } from '../src/index.js'
import * as text from '../src/text.js'

describe('xlcalc/text', () => {
  it("holds the root's functions themselves", () => {
    expect(Object.keys(text).sort()).toEqual(['DEFAULT_ROW_HEIGHT', 'FUNCTION_CATEGORIES', 'MAX_OUTLINE', 'MAX_ROW_HEIGHT', 'callAt', 'columnWidthFromChars', 'columnWidthFromPoints', 'columnWidthToChars', 'columnWidthToPoints', 'compareValues', 'escapeColumn', 'findMatcher', 'findReplacer', 'formatClipboard', 'nameAt', 'numberFormatCategory', 'numberFormatChoices', 'parseClipboard', 'quoteSheetIfNeeded', 'refAt', 'scanReferences'])
    for (const [k, f] of Object.entries(text)) expect(root[k as keyof typeof root], k).toBe(f)
  })
})

describe('escapeColumn', () => {
  it("puts ' before [ ] # ' and a leading @, as Excel writes them", () => {
    expect(['x#y', '[br]', "it's", '@at', 'a@b', 'Total $', 'Unit Price', '2024', 'Сумма'].map(escapeColumn)).toEqual(["x'#y", "'[br']", "it''s", "'@at", 'a@b', 'Total $', 'Unit Price', '2024', 'Сумма'])
  })

  it('a name escaped reads back as the name, alone and in a span', () => {
    const names = ['x#y', '[br]', "it's", '@at', 'a]#[b', "''"]
    const tables = [{ name: 'T', columns: names }]
    for (const n of names) {
      expect(validateFormula(`T[${escapeColumn(n)}]`, { tables }), n).toMatchObject({ ok: true, canonical: `T[${escapeColumn(n)}]` })
      expect(validateFormula(`T[[#Headers],[${escapeColumn(n)}]:[${escapeColumn(n)}]]`, { tables }), n).toMatchObject({ ok: true })
    }
  })

  it('wrong input is an InputError', () => {
    expect(() => escapeColumn(1 as never)).toThrow(new InputError('name', 'expected a string, got number 1'))
  })
})
