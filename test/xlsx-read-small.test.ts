/**
 * Small readings of a file against Excel for Mac and for Windows, each set's prediction, committed before Excel ran it, scored with the
 * misses named, then the reader and the engine on what Excel opened: colours of several kinds and the first two fills (`xlsx/fills-read.ts`);
 * a 2010 data bar's lengths (`xlsx/bars-held.ts`); `_xlop.`, a LAMBDA's optional parameter, and `_xlfn.LONGTEXT`
 * (`links/longtext-typed.ts`, `links/longtext-file.ts`); the date codes `e` and `ee` in TEXT (`links/date-e-text.ts`).
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredLinks } from '../bench/coverage/links.js'
import { linkFile } from '../bench/coverage/links/cases.js'
import { buildLinksCases as dateCases } from '../bench/coverage/links/date-e-text.js'
import { REFUSED, buildLinksCases as longCases } from '../bench/coverage/links/longtext-file.js'
import { BAR_CASES, barFile, barRule } from '../bench/coverage/xlsx/bars-held.js'
import { FILL_FILES, fillBook } from '../bench/coverage/xlsx/fills-read.js'
import { compute, displayValue, type ConditionalFormat } from '../src/index.js'
import { readXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { cellsOf } from './looks-book.js'

const PLATFORMS = [
  ['Excel for Mac', ''],
  ['Excel for Windows', 'win/'],
] as const
const excel = (dir: string, path: string) => new URL(`./excel/${dir}${path}`, import.meta.url)
const json = <T>(url: URL) => JSON.parse(readFileSync(url, 'utf8')) as T
const predicted = <T>(path: string) => json<T>(new URL(`../bench/coverage/xlsx/${path}`, import.meta.url))
const wanted = (v: unknown) => (typeof v === 'object' && v !== null && 'error' in v ? { t: 'e', v: (v as { error: string }).error } : typeof v === 'number' ? { t: 'n', v } : { t: 's', v })
const value = (v: { t: string; v: unknown } | undefined) => (v?.t === 'e' ? { error: v.v } : v?.v)

/** Excel's names of a pattern: COM's numbers, AppleScript's words. */
const PATTERN: Record<string | number, string> = { [-4142]: 'none', 1: 'solid', 17: 'gray125', 'pattern none': 'none', 'pattern solid': 'solid', 'pattern gray 16': 'gray125' }

describe.each(PLATFORMS)('%s', (_, dir) => {
  describe('colours of several kinds and the first two fills', async () => {
    interface Shown { fill: string; pattern: string | number; font: string }
    const { cells } = json<{ cells: Record<string, Shown> }>(excel(dir, 'xlsx/fills-read.json'))
    const rule = predicted<{ cells: Record<string, { pattern: string; fill?: string; font?: string }> }>('fills-read.predicted.json').cells

    it('as predicted', () => {
      for (const [id, p] of Object.entries(rule)) {
        expect(PATTERN[cells[id].pattern], id).toBe(p.pattern)
        if (p.fill) expect(cells[id].fill, id).toBe(p.fill)
        if (p.font) expect(cells[id].font, id).toBe(p.font)
      }
    })

    it.each(FILL_FILES.map((f) => [f.id, f] as const))('the reader on %s', async (_id, f) => {
      const r = await readXlsx(await fillBook(f))
      const at = cellsOf(r.book.sheets[0])
      for (const c of f.cells) {
        const style = at[c.cell]?.style === undefined ? undefined : r.book.styles![at[c.cell].style!]
        expect(style?.fill?.color ?? 'none', c.id).toBe(PATTERN[cells[c.id].pattern] === 'solid' ? cells[c.id].fill : 'none')
        expect(style?.font?.color ?? '#000000', c.id).toBe(cells[c.id].font)
      }
    })
  })

  describe('a 2010 data bar\'s lengths', () => {
    interface Bar { old: Record<string, string>; x14: Record<string, string> | null; read: { minLength: number; maxLength: number } }
    const cases = predicted<{ cases: Record<string, Bar> }>('bars-held.predicted.json').cases
    /** Predictions Excel did not confirm: a bar Excel makes through COM has lengths on the 2007 part alone, and no 2010 part. */
    const MISSED: Record<string, Bar['x14']> = { made: null }
    const lengths = (tag: string, xml: string) => {
      const m = new RegExp(`<${tag}\\b([^>]*)>`).exec(xml)
      return m && Object.fromEntries([...m[1].matchAll(/(minLength|maxLength)="(\d+)"/g)].map((x) => [x[1], x[2]]))
    }
    const ids = [...BAR_CASES.map((c) => c.id), ...(dir ? ['made'] : [])]

    it.each(ids)('%s: the parts Excel saved, and the reader on them', async (id) => {
      const file = readFileSync(excel(dir, id === 'made' ? 'xlsx/bars-held-made/made.xlsx' : `xlsx/bars-held/${id}.xlsx`))
      const xml = new TextDecoder().decode((await openArchive(file).bytes('xl/worksheets/sheet1.xml'))!)
      expect(lengths('dataBar', xml)).toEqual(cases[id].old)
      expect(lengths('x14:dataBar', xml)).toEqual(id in MISSED ? MISSED[id] : cases[id].x14)
      const rule = (await readXlsx(file)).book.sheets[0].conditionalFormats?.[0] as Extract<ConditionalFormat, { type: 'dataBar' }>
      const old = rule.excel2007 === true
      expect([rule.minLength ?? (old ? 10 : 0), rule.maxLength ?? (old ? 90 : 100)]).toEqual([cases[id].read.minLength, cases[id].read.maxLength])
    })
  })

  describe('_xlop. and _xlfn.LONGTEXT in a file', () => {
    const set = json<StoredLinks>(excel(dir, 'links/longtext-file.json'))
    /**
     * Predictions Excel did not confirm. `_xlfn.LONGTEXT` is no function of either Excel, whatever its arguments: #NAME?. `_xlop.` on a
     * use of the parameter is #NAME? too. A literal of 300 characters: Excel for Mac 16.114.3 opens it, Excel for Windows 16.0.20430 refuses the file.
     */
    const NAME_NOT_VALUE = ['lt-two', 'lt-len', 'lt-empty', 'lt-quote', 'lt-seventeen', 'lt-long']
    const NAME_NOT_REFUSED = ['xlop-in-body', 'lt-one', 'lt-eighteen', 'lt-cell', 'lt-number']

    it.each(longCases().cases.map((c) => [c.id, c] as const))('%s', async (id, c) => {
      const got = set.answers[id]?.reads[1].values['S!A1']
      const refused = set.rejected?.[id] !== undefined
      if (NAME_NOT_VALUE.includes(id) || NAME_NOT_REFUSED.includes(id)) expect(got).toEqual({ t: 'e', v: 'NAME' })
      else if (id === 'literal-300') expect(refused).toBe(dir !== '')
      else if (REFUSED.includes(id)) expect(refused).toBe(true)
      else expect(got).toEqual(wanted(c.predicted!['S!A1']))
      // the engine on `_xlop.`: test/lambda.test.ts
      if (!refused && !id.startsWith('xlop-')) expect(compute((await readXlsx(await linkFile(c))).book).sheets[0].cells.A1?.value).toEqual(value(got))
    })
  })

  describe('the date codes e and ee', () => {
    const set = json<StoredLinks>(excel(dir, 'links/date-e-text.json'))
    /** A prediction Excel did not confirm: TEXT takes no upper-case `E`, which a cell's format refuses too (`display/date-e-held.ts`). */
    const MISSED = ['upper']

    it.each(dateCases().cases.map((c) => [c.id, c] as const))('TEXT with %s', async (id, c) => {
      const got = set.answers[id].reads[1].values['S!A1']
      expect(got).toEqual(MISSED.includes(id) ? { t: 'e', v: 'VALUE' } : wanted(c.predicted!['S!A1']))
      expect(compute((await readXlsx(await linkFile(c))).book).sheets[0].cells.A1?.value).toEqual(value(got))
    })
  })
})

describe.skipIf(!existsSync(excel('win/', 'links/longtext-typed.json')))('typed into Excel for Windows', () => {
  const set = json<StoredLinks>(excel('win/', 'links/longtext-typed.json'))
  const text = async (id: string) => new TextDecoder().decode((await openArchive(readFileSync(excel('win/', `links/longtext-typed/${id}.xlsx`))).bytes('xl/worksheets/sheet1.xml'))!)

  it('`[y]` is saved as `_xlop.y`, `_xlpm.y` in the calculation', async () => {
    expect(await text('optional-one')).toContain('_xlfn.LAMBDA(_xlpm.x,_xlop.y,IF(_xlfn.ISOMITTED(_xlpm.y),_xlpm.x,_xlpm.x+_xlpm.y))(1)')
  })

  it('values as predicted, and what it refuses: a required parameter after an optional one; a literal over 255, which was predicted taken', () => {
    const shown = Object.fromEntries(set.cases.map((c) => [c.id, set.answers[c.id].reads[1].values['S!A2']]))
    expect(shown).toMatchObject({ 'optional-one': { v: 1 }, 'optional-two': { v: 3 }, 'optional-three': { v: 'VALUE' }, 'optional-empty': { v: 1 }, 'literal-255': { v: 255 } })
    for (const id of ['required-after', 'literal-256', 'literal-4095', 'literal-4096']) expect(set.answers[id].errors, id).toBeDefined()
  })
})

describe('the date codes e and ee in a cell', () => {
  it.each([
    ['Excel for Mac', 'display/date-e-mac.json'],
    ['Excel for Windows', 'win/display/date-e-held.json'],
  ])('%s: as predicted, but for upper case, which Excel refuses a file for; the engine shows the same', (_, path) => {
    const set = json<{ cases: { id: string; cell: { value: number; format: string }; predicted: { text: string } }[]; answers: Record<string, { text: string }>; rejected?: string[] }>(new URL(`./excel/${path}`, import.meta.url))
    expect(set.rejected ?? []).toEqual(path.startsWith('win') ? ['e5', 'e6'] : [])
    for (const c of set.cases.filter((x) => !set.rejected?.includes(x.id))) {
      expect(set.answers[c.id].text, c.id).toBe(c.predicted.text)
      expect(displayValue(c.cell.value, { numberFormat: c.cell.format, width: 40 }).text, c.id).toBe(c.predicted.text)
    }
  })
})

describe('the writer of a 2010 data bar', () => {
  const written = async (c: { minLength: number; maxLength: number }) => {
    const r = await readXlsx(await barFile({ id: 'x', ...c }))
    return [(r.book.sheets[0].conditionalFormats?.[0] as ReturnType<typeof barRule>), new TextDecoder().decode((await openArchive(await barFile({ id: 'x', ...c })).bytes('xl/worksheets/sheet1.xml'))!)] as const
  }

  it('puts lengths other than 0 and 100 on both parts, leaving out what is a default', async () => {
    expect((await written({ minLength: 5, maxLength: 95 }))[1]).toContain('<dataBar minLength="5" maxLength="95">')
    expect((await written({ minLength: 10, maxLength: 95 }))[1]).toContain('<dataBar maxLength="95">')
    expect((await written({ minLength: 0, maxLength: 100 }))[1]).toContain('<dataBar>')
    expect((await written({ minLength: 5, maxLength: 95 }))[0]).toMatchObject({ minLength: 5, maxLength: 95 })
  })
})
