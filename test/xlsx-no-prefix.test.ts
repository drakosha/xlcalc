/**
 * A newer function in a file without its prefix (`noPrefix`), against Excel for Mac and for Windows
 * (`bench/coverage/links/no-prefix.ts`): the reader reports a call exactly where Excel does not take it as the function, and the
 * values and the formula's spelling stay what they were.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredLinks } from '../bench/coverage/links.js'
import { linkFile } from '../bench/coverage/links/cases.js'
import { buildLinksCases } from '../bench/coverage/links/no-prefix.js'
import { compute } from '../src/index.js'
import { noPrefixCalls } from '../src/xlsx/formulas/from-file.js'
import { readXlsx } from '../src/xlsx.js'
import { sheet } from './xlsx-fixtures.js'

const stored = (dir: string) => JSON.parse(readFileSync(new URL(`./excel/${dir}links/no-prefix.json`, import.meta.url), 'utf8')) as StoredLinks
const PLATFORMS = [
  ['Excel for Mac', stored('')],
  ['Excel for Windows', stored('win/')],
] as const

const cases = buildLinksCases().cases
/** Calls of functions Excel takes, one prefixed and one inside quotes among them, and one hidden by `IFERROR`, which Excel answers with its fallback. */
const TAKEN = new Set(['with-prefix', 'old-function', 'in-text'])
const HIDDEN = new Set(['in-iferror'])

describe('noPrefix against Excel', () => {
  const reports = Promise.all(
    cases.map(async (c) => {
      const r = await readXlsx(await linkFile(c))
      return [c.id, r.issues.filter((i) => i.code === 'noPrefix')] as const
    }),
  ).then((x) => new Map(x))

  describe.each(PLATFORMS)('%s', (_, set) => {
    it('has answered every case, a refused file among them', () => {
      expect(Object.keys(set.answers).length + Object.keys(set.rejected ?? {}).length).toBe(cases.length)
    })

    it.each(cases.map((c) => c.id))('%s: reported where Excel does not take the call', async (id) => {
      const got = (await reports).get(id)!
      const refused = set.rejected?.[id] !== undefined
      const value = set.answers[id]?.reads[1].values['S!A1']
      const name = value?.t === 'e' && value.v === 'NAME'
      expect(got.length > 0).toBe(!TAKEN.has(id) && (refused || name || HIDDEN.has(id)))
      if (refused) expect(got[0].message).toContain('does not open')
    })
  })
})

describe('noPrefixCalls', () => {
  it('finds a call without its prefix, once, upper-case, in the order of the calls', () => {
    expect(noPrefixCalls('xlookup(1,A1:A2,B1:B2)+SUM(DAYS(2,1),XLOOKUP(1,A1:A2,B1:B2))')).toEqual(['XLOOKUP', 'DAYS'])
  })

  it.each([
    ['_xlfn.XLOOKUP(1,A1:A2,B1:B2)'],
    ['_xlfn._xlws.FILTER(A1:A2,B1:B2)'],
    ['"XLOOKUP(1)"'],
    ["'XLOOKUP(1'!A1"],
    ['Sales[XLOOKUP(1)]'],
    ['S!DAYS(1)'],
    ['яDAYS(1)'],
    ['DAYS'],
    ['SUM(A1)'],
    ['NETWORKDAYS.INTL(1,2)'],
  ])('leaves %s alone', (f) => {
    expect(noPrefixCalls(f)).toBeUndefined()
  })

  it('sees a call after text in quotes', () => {
    expect(noPrefixCalls('"(a"&DAYS(2,1)')).toEqual(['DAYS'])
  })
})

describe('reading a call without its prefix', () => {
  it('reports each function on its sheet with its cells, keeps the spelling and calculates', async () => {
    const rows = '<row r="1"><c r="A1"><f>XLOOKUP(2,{1,2},{3,4})</f></c><c r="B1"><f>DAYS(10,1)</f></c></row><row r="2"><c r="A2"><f>XLOOKUP(1,{1,2},{3,4})</f></c><c r="B2"><f>_xlfn.DAYS(10,1)</f></c></row>'
    const r = await sheet(rows)
    expect(r.issues.filter((i) => i.code === 'noPrefix').map(({ fn, cells, sheet }) => ({ fn, cells, sheet }))).toEqual([
      { fn: 'XLOOKUP', cells: ['A1', 'A2'], sheet: 'S' },
      { fn: 'DAYS', cells: ['B1'], sheet: 'S' },
    ])
    expect(r.book.sheets[0].cells.A1).toEqual({ formula: '@XLOOKUP(2,{1,2},{3,4})' })
    expect(compute(r.book).sheets[0].cells.A1.value).toBe(4)
  })

  it('names a refused file in the message of FILTER and SORT', async () => {
    const r = await sheet('<row r="1"><c r="A1"><f>SORT({2,1})</f></c></row>')
    expect(r.issues.find((i) => i.code === 'noPrefix')?.message).toContain('does not open')
  })
})
