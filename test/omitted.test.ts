/**
 * Arguments left out between commas beyond Excel's answers (`test/excel/cases/omit*.json`): the parse and the
 * canonical text, the counts, the calls Excel does not load, what each function gets by the rules of
 * `src/functions/omitted.ts`, ru-RU text and structure operations keeping them.
 */
import { describe, expect, it } from 'vitest'
import { REFUSED as HELD_REFUSED } from '../bench/coverage/sets/omit-held.js'
import { REFUSED } from '../bench/coverage/sets/omit.js'
import { compute, restructure, toInvariant, toLocal, validateFormula, validateLocalFormula } from '../src/index.js'
import { ruRU } from './locales.js'
import { FUNCTIONS } from '../src/functions.js'
import { canonical, parse } from '../src/syntax.js'
import { E } from '../src/values.js'
import { engineFormula } from './excel-cases.js'

const one = (formula: string) => compute({ sheets: [{ name: 'S', cells: { A1: { formula } } }] }).sheets[0].cells.A1.value

describe('arguments left out', () => {
  it('parse as left-out nodes; a call of none has no arguments', () => {
    const a = parse('IF(A1,,1)')
    expect(a.k === 'fn' && a.args.map((x) => x.k)).toEqual(['ref', 'missing', 'num'])
    const b = parse('SUM(,)')
    expect(b.k === 'fn' && b.args.map((x) => x.k)).toEqual(['missing', 'missing'])
    const c = parse('NA()')
    expect(c.k === 'fn' && c.args).toEqual([])
  })

  it('keep their place in the canonical text', () => {
    for (const f of ['IF(A1,,1)', 'SUM(,)', 'TEXTSPLIT("a,b",",",,,1)', 'ROUND(1.5,)', 'CONCAT(,"a",)']) expect(canonical(parse(f))).toBe(f)
    expect(validateFormula('if( a1 , , 1 )').ok && validateFormula('if( a1 , , 1 )')).toMatchObject({ canonical: 'IF(A1,,1)' })
  })

  it('count as arguments: a count the function does not take is a syntax error', () => {
    expect(validateFormula('ROUND(1.5,)').ok).toBe(true)
    expect(validateFormula('ROUND(1.5,,)')).toMatchObject({ ok: false, code: 'syntax', message: 'ROUND takes 2 arguments, got 3' })
    expect(validateFormula('ABS(,)')).toMatchObject({ ok: false, code: 'syntax' })
  })

  it('Excel does not load one where it takes only a reference (the omit sets)', () => {
    for (const f of [...REFUSED, ...HELD_REFUSED]) expect(validateFormula(engineFormula(f)), f).toMatchObject({ ok: false, code: 'syntax' })
    expect(validateFormula('SUMIF(A1:A3,">1",)')).toMatchObject({ message: 'SUMIF takes a reference as its argument 3, which cannot be left out' })
    // AGGREGATE's references and the database functions' ranges load: #VALUE!
    for (const f of ['AGGREGATE(9,0,)', 'DSUM(A1:B3,"x",)']) expect(validateFormula(f).ok, f).toBe(true)
  })

  it('a blank read by kind, #VALUE! for an array or a reference, the default, #N/A in the Analysis ToolPak', () => {
    expect([FUNCTIONS.ROUND.omitted(1), FUNCTIONS.IF.omitted(1), FUNCTIONS.CORREL.omitted(0), FUNCTIONS.INDEX.omitted(0)]).toEqual([null, null, E.VALUE, E.VALUE])
    expect([FUNCTIONS.SEQUENCE.omitted(1), FUNCTIONS.XLOOKUP.omitted(3), FUNCTIONS.TEXTJOIN.omitted(1), FUNCTIONS.VLOOKUP.omitted(1)]).toEqual([undefined, undefined, true, null])
    // the former Analysis ToolPak: a required argument #N/A, an optional one not given; ERF's upper limit #N/A
    expect([FUNCTIONS.EDATE.omitted(0), FUNCTIONS.DEC2BIN.omitted(1), FUNCTIONS.ACCRINT.omitted(4), FUNCTIONS.ERF.omitted(1)]).toEqual([E.NA, undefined, undefined, E.NA])
    // repeated arguments: PRODUCT's and LCM's are not given, HSTACK's #VALUE!, SORTBY's by-arrays blanks and orders not given
    expect([FUNCTIONS.PRODUCT.omitted(3), FUNCTIONS.LCM.omitted(2), FUNCTIONS.HSTACK.omitted(1), FUNCTIONS.SORTBY.omitted(3), FUNCTIONS.SORTBY.omitted(4)]).toEqual([undefined, undefined, E.VALUE, null, undefined])
  })

  it('compute as Excel does, in both execution modes', () => {
    const cells = { A1: { formula: 'IF(TRUE,)&"x"' }, A2: { formula: 'SUM(1,,2)*10+COUNT(1,,2)' }, A3: { formula: 'PRODUCT(2,,3)' }, A4: { formula: 'LEFT("abc",)&"|"&LEFT("abc")' }, A5: { formula: 'SUM(SEQUENCE(,3))' } }
    for (const execution of ['closures', 'eval'] as const) {
      const out = compute({ sheets: [{ name: 'S', cells }] }, { execution }).sheets[0].cells
      expect([out.A1.value, out.A2.value, out.A3.value, out.A4.value, out.A5.value]).toEqual(['x', 33, 6, '|a', 6])
    }
    expect(one('IF(TRUE,)')).toBe(0)
  })

  it('ru-RU text: the list separator between them, both ways', () => {
    expect(toInvariant('=ЕСЛИ(A1;;1)', ruRU)).toMatchObject({ ok: true, formula: 'IF(A1,,1)' })
    expect(toLocal('TEXTSPLIT("a,b",",",,,1)', ruRU)).toMatchObject({ ok: true, formula: 'ТЕКСТРАЗД("a,b";",";;;1)' })
    expect(validateLocalFormula('=ОКРУГЛ(1,5;)', ruRU)).toMatchObject({ ok: true, canonical: 'ROUND(1.5,)' })
  })

  it('structure operations move the references around them', () => {
    const book = { sheets: [{ name: 'S', cells: { A1: { value: 1 }, B1: { formula: 'IF(A1,,A1)' } } }] }
    const r = restructure(book, { op: 'insertRows', sheet: 'S', row: 1 })
    expect(r.ok && r.book.sheets[0].cells.B2.formula).toBe('IF(A2,,A2)')
  })
})
