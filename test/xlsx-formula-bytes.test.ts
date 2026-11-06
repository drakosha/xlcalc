/**
 * The writer's margin over a formula's tokens and arguments, against Excel for Windows (`bench/coverage/links/formula-bytes-held.ts`,
 * `formula-bytes-args.ts`): whatever Excel refused the writer leaves out, as its value with an issue `formula`. What Excel opened and the
 * writer leaves out all the same is listed: the price of a margin, on formulas no corpus has. The two sets' predictions keep their score.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredLinks } from '../bench/coverage/links.js'
import type { BookInput, CellInput } from '../src/contract.js'
import { fromFile } from '../src/xlsx/formulas/from-file.js'
import { parseTokens } from '../src/syntax/parse.js'
import { fileTree, type FileNames } from '../src/xlsx/write/formula.js'
import { formulaWeight, SHORT_FORMULA } from '../src/xlsx/write/excel-loads.js'
import { writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'

const SETS = ['formula-bytes-held', 'formula-bytes-args', 'file-refusals-4'] as const
const stored = (set: string) => JSON.parse(readFileSync(new URL(`./excel/win/links/${set}.json`, import.meta.url), 'utf8')) as StoredLinks

const NAMES: FileNames = {
  sheets: new Set(['S', 'T', 'U']),
  tables: [{ name: 'Sales', columns: ['Qty', 'F'] }],
  externalBooks: [{ target: 'other.xlsx', sheets: [{ name: 'T', cells: {} }], names: [{ name: 'ext', ref: 'T!$B$1' }] }],
  defined: new Map([['rate', 'rate'], ['sn', 'sn']]),
}
const formulaOf = (c: StoredLinks['cases'][number]) => {
  const f = c.linked.S?.A1
  return f === undefined ? undefined : typeof f === 'string' ? f : f.formula
}

describe.each(SETS)('%s', (set) => {
  const s = stored(set)
  const cases = s.cases.filter((c) => formulaOf(c) !== undefined && !c.id.startsWith('defined') && !c.id.startsWith('few-') && !c.id.startsWith('many-'))
  const written = (c: StoredLinks['cases'][number]) => fileTree(fromFile(formulaOf(c)!).formula, NAMES) !== null

  it('the writer writes nothing Excel refused over its bytes', () => {
    expect(cases.filter((c) => s.rejected?.[c.id] && written(c)).map((c) => c.id)).toEqual([])
  })

  it('what Excel opened and the margin leaves out', () => {
    expect(cases.filter((c) => s.answers[c.id] && !written(c)).map((c) => c.id)).toEqual(LEFT_OUT[set])
  })
})

/** What Excel opened and the writer leaves out, as a margin. */
const LEFT_OUT: Record<(typeof SETS)[number], string[]> = {
  // Excel opened these on both sides, each with a function newer than Excel 2003 and whatever its bytes: why was not looked for
  'formula-bytes-held': ['future-open', 'future-over', 'future-ws-open', 'future-ws-over', 'iferror-open', 'iferror-over', 'ifna-open', 'ifna-over', 'ifs-open', 'ifs-over', 'switch-open', 'switch-over', 'single-open', 'single-over', 'spill-open', 'spill-over', 'let-open', 'let-over'],
  // built on the edge of a hypothesis that was not taken
  'formula-bytes-args': ['cell-open', 'cell-over', 'range-open', 'range-over', 'number-open', 'number-over', 'integer-open', 'integer-over', 'text-short-open', 'text-short-over', 'text-long-open', 'text-long-over', 'sheet-open', 'sheet-over', 'max-open', 'max-over', 'abs-open', 'abs-over', 'nested-open', 'nested-over', 'if-open', 'if-over'],
  // `cells-2048` by the parser, of over 2,000 operators; `sum-cells` by the margin
  'file-refusals-4': ['cells-2048', 'sum-cells'],
}

describe('the predictions, committed before Excel ran each set', () => {
  /** The rule of `formula-bytes-args` was set before it ran: every family as predicted, or a byte for an argument is not taken. */
  it.each([
    ['formula-bytes-held', 71, 69],
    ['formula-bytes-args', 13, 13],
  ] as const)('%s scored what it scored', (set, hit, miss) => {
    const s = stored(set)
    const got = s.cases.filter((c) => !!c.refused === !!s.rejected?.[c.id]).length
    expect([got, s.cases.length - got]).toEqual([hit, miss])
  })
})

describe('a long formula', () => {
  const S = (cells: Record<string, CellInput>): BookInput => ({ sheets: [{ name: 'S', cells: { B1: { value: 1 }, ...cells } }] })
  const of = (n: number, unit: string) => {
    const groups: string[] = []
    for (let i = 0; i < n; i += 250) groups.push(`SUM(${Array.from({ length: Math.min(250, n - i) }, () => unit).join(',')})`)
    return groups.join('+')
  }
  const sheet = async (book: BookInput) => {
    const r = await writeXlsx({ book }, { values: 'omit' })
    return { xml: (await openArchive(r.data).text('xl/worksheets/sheet1.xml'))!, issues: r.issues }
  }

  it('over the margin it goes as its value, with an issue, under it stays a formula', async () => {
    const heavy = of(1200, 'B1:B2')
    const r = await sheet(S({ A1: { formula: heavy, value: 7 }, A2: { formula: 'B1+1', value: 2 } }))
    expect(r.xml).not.toMatch(/<c r="A1"[^>]*><f/)
    expect(r.xml).toMatch(/<c r="A2"[^>]*><f/)
    expect(r.issues.filter((i) => i.code === 'formula').length).toBe(1)
    const light = await sheet(S({ A1: { formula: of(900, 'B1:B2'), value: 7 } }))
    expect(light.xml).toMatch(/<c r="A1"[^>]*><f/)
  })

  it('formulas of one shape are counted each, as a reference is a cell or a range', async () => {
    // the same shape: a cell or a range is a hole of it, and two of them weigh differently
    const light = of(1200, 'B1')
    const heavy = of(1200, 'B1:B2')
    const r = await sheet(S({ A1: { formula: light, value: 1 }, A2: { formula: heavy, value: 2 } }))
    expect(r.xml).toMatch(/<c r="A1"[^>]*><f/)
    expect(r.xml).not.toMatch(/<c r="A2"[^>]*><f/)
  })

  it('a text of the length not counted cannot reach the margin, in the forms that weigh the most for their text', () => {
    for (const unit of ['a', '@a', '.5', '1:1', '{1}', '[a]', 'a b', '@a @a', 'a()', '@a()', 'A1#', '@1', 'IF(a,a)', 'N(@a)', 'MAX(a,a)', 'IF(,)', 'CHOOSE(,)']) {
      const text = Array.from({ length: Math.floor(SHORT_FORMULA / (unit.length + 1)) }, () => unit).join('+').slice(0, SHORT_FORMULA)
      const cut = text.slice(0, text.lastIndexOf('+'))
      let tree
      try {
        tree = parseTokens(cut, null, {})
      } catch {
        continue
      }
      expect(formulaWeight(tree), unit).toBeLessThanOrEqual(15_000)
    }
  })
})
