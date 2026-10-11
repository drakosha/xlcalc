/**
 * The harnesses' shared helpers on the inputs their copies took before (`stored.ts`, `case-book.ts`, `book.ts`, `halving`,
 * Windows's `again`), without Excel.
 */
import { afterEach, describe, expect, it } from 'vitest'
import { functionParts, inParts } from './book.js'
import { editOf, eachRectCell, mirrorFormulas, refusedOf } from './case-book.js'
import { halving } from './excel.js'
import { answeredLine, chosen, keptAnswers, keptSame, runCases, setModule, unpredicted } from './stored.js'
import { columnCells, rowCells } from './struct.js'
import { again } from '../windows/remote.js'

afterEach(() => {
  delete process.env.FHELPERS_ONLY
})

describe('stored answers', () => {
  it('keeps an answer while its case is the same by the key, a prediction aside', () => {
    const prev = { cases: [{ id: 'a', f: 1, predicted: 1 }, { id: 'b', f: 2 }, { id: 'c', f: 3 }], answers: { a: 'A', b: 'B' } }
    const now = { cases: [{ id: 'a', f: 1, predicted: 2 }, { id: 'b', f: 9 }, { id: 'c', f: 3 }] }
    expect(keptSame(now, prev, unpredicted)).toEqual({ a: 'A' })
    expect(keptSame(now, prev)).toEqual({})
    expect(keptSame(now, null)).toEqual({})
  })

  it('keeps a case sent as a file while the case and the file are the same, an answer only with its saved file', () => {
    const prev = { cases: [{ id: 'a' }, { id: 'b' }, { id: 'c' }], sent: { a: 'h', b: 'h', c: 'h' }, answers: { a: 1, b: 2 }, rejected: { c: 'no' } }
    expect(keptAnswers(prev.cases, { a: 'h', b: 'h', c: 'h' }, prev, (id) => id === 'a')).toEqual({ answers: { a: 1 }, rejected: { c: 'no' } })
    expect(keptAnswers(prev.cases, { a: 'x', b: 'h', c: 'h' }, prev, () => true)).toEqual({ answers: { b: 2 }, rejected: { c: 'no' } })
  })

  it('runs the cases F*_ONLY names, and ends however the run ends', async () => {
    const cases = [{ id: 'a' }, { id: 'b' }, { id: 'c' }]
    process.env.FHELPERS_ONLY = 'a,c'
    expect(cases.filter(chosen('FHELPERS_ONLY')).map((c) => c.id)).toEqual(['a', 'c'])
    const ran: string[] = []
    let ended = -1
    await runCases(cases, 'FHELPERS_ONLY', async (c) => void ran.push(c.id), (n) => (ended = n))
    expect([ran, ended]).toEqual([['a', 'c'], 2])
    await expect(runCases(cases, 'FHELPERS_NONE', async (c) => {
      if (c.id === 'b') throw new Error('stop')
    }, (n) => (ended = n))).rejects.toThrow('stop')
    expect(ended).toBe(1)
  })

  it('names the cases run and answered', () => {
    expect(answeredLine('cond x', 2, { a: 1 }, [1, 2, 3])).toBe('cond x: 2 cases run, 1 of 3 answered')
    expect(answeredLine('cond x', 2, { a: 1 }, [1, 2, 3], {})).toBe('cond x: 2 cases run, 1 of 3 answered, 0 not loaded')
  })

  it('takes a set by a name that is a file name only', async () => {
    await expect(setModule('struct', '../stored')).rejects.toThrow('set name')
    await expect(setModule('struct', undefined)).rejects.toThrow('set name')
  })
})

describe('a case as Excel takes it', () => {
  it('types a value as the scripts did', () => {
    expect([null, 5, -0.5, 'a\tb', { formula: 'SUM(A1)' }, { local: 'СУММ(A1)' }].map(editOf)).toEqual([['x', ''], ['n', '5'], ['n', '-0.5'], ['s', 'a\tb'], ['f', '=SUM(A1)'], ['l', 'СУММ(A1)']])
  })

  it('lists the steps refused by index, none when all were taken', () => {
    expect(refusedOf(['', 'e1', undefined, null, 'e4'])).toEqual({ 1: 'e1', 4: 'e4' })
    expect(refusedOf(['', ''])).toBeUndefined()
  })

  it('walks the rectangles row by row', () => {
    const seen: string[] = []
    eachRectCell([{ sheet: 'S', range: 'B2:C3' }, { sheet: 'T', range: 'A1' }], (key, i, r, j) => seen.push(`${key}:${i}${r}${j}`))
    expect(seen).toEqual(['S!B2:000', 'S!C2:001', 'S!B3:010', 'S!C3:011', 'T!A1:100'])
  })

  it('writes the mirror formulas excel-tables.applescript wrote, over the top-left cell of a quoted sheet', () => {
    const t = ['IF(ISBLANK(#),-5,IF(ISERROR(#),ERROR.TYPE(#),IF(ISNUMBER(#),0,IF(ISLOGICAL(#),IF(#,-1,-2),IF(ISTEXT(#),-3,-4)))))', 'IF(ISNUMBER(#),#,0)', 'IF(ISTEXT(#),#,"")']
    expect(mirrorFormulas({ sheet: "It's", range: 'B2:C4' })).toEqual(t.map((f) => '=' + f.split('#').join("'It''s'!B2")))
  })

  it('spells rows and columns to insert or delete as cells', () => {
    expect([rowCells(5), rowCells(5, 3), columnCells('Z'), columnCells('Z', 2), columnCells('XFC', 2)]).toEqual(['A5:A5', 'A5:A7', 'Z1:Z1', 'Z1:AA1', 'XFC1:XFD1'])
  })
})

describe('parts and halves', () => {
  it('cuts items in parts, and the measurement by whole functions', () => {
    expect(inParts([1, 2, 3, 4, 5], 2)).toEqual([[1, 2], [3, 4], [5]])
    expect(inParts([], 2)).toEqual([])
    const fns = [...Array(300).fill('A'), ...Array(200).fill('B'), ...Array(500).fill('C'), 'D'].map((fn) => ({ fn }))
    expect(functionParts(fns).map((p) => p.length)).toEqual([300, 200, 500, 1])
  })

  it('halves a book Excel refuses down to the items it refuses', async () => {
    const refused: number[] = []
    const labels: string[] = []
    await halving([1, 2, 3, 4, 5], '1', async (part, label) => (labels.push(label), !part.includes(4)), (x) => refused.push(x))
    expect([refused, labels]).toEqual([[4], ['1', '1a', '1b', '1ba', '1bb', '1bba', '1bbb']])
  })

  it('gives a part back to its jobs whole, halved or not at all', () => {
    const refused: number[] = []
    expect(again([1, 2, 3], undefined, true, (x) => refused.push(x))).toEqual([[1, 2, 3]])
    expect(again([1, 2, 3], { rejected: 'r', seconds: 1 }, false, (x) => refused.push(x))).toEqual([[1], [2, 3]])
    expect(again([7], undefined, false, (x) => refused.push(x))).toEqual([])
    expect(refused).toEqual([7])
  })
})
