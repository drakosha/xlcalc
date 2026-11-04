/**
 * The two forms of a formula over a link against Excel for Mac and Excel for Windows (`bench/coverage/links/forms.ts`): each
 * formula of the old kind and as a dynamic array, in files made by hand. Where the writer writes the old kind, Excel computes
 * the two forms alike; where it keeps a dynamic array, they part or the conversion cannot tell. What the engine computes of
 * the old kind is what Excel computes of it, but for the cells a link's closed book answers otherwise in Excel (`OTHERWISE`).
 */
import { describe, expect, it } from 'vitest'
import type { StoredLinks } from '../bench/coverage/links.js'
import { buildLinksCases } from '../bench/coverage/links/forms.js'
import { linkFile, linksBook, type LinkedFormula, type LinksCase } from '../bench/coverage/links/cases.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import type { CellValue } from '../src/contract.js'
import { compute } from '../src/index.js'
import { fromLegacy, writeXlsx } from '../src/xlsx.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { storedSets } from './stored-sets.js'

const PLATFORMS = [
  ['Excel for Mac', ''],
  ['Excel for Windows', 'win/'],
] as const

/**
 * Excel's answers the engine does not give, whichever the form: over a closed book it counts the blanks of a range and takes a
 * marked sheet's values otherwise than when the rules were measured (`test/links-excel.test.ts`, `VERSIONED`), and Excel for
 * Mac and for Windows differ over a marked sheet.
 */
const OTHERWISE: Record<string, Record<string, CellValue>> = {
  '': { 'cells S!A25': 1, 'converted S!AG2': 5, 'converted S!AU2': 3, 'converted S!AW2': 3, 'converted S!AY2': 1 },
  'win/': { 'cells S!A25': 1, 'converted S!AG2': 5, 'converted S!AY2': 3 },
}

const cases = buildLinksCases().cases
const formula = (f: LinkedFormula) => (typeof f === 'string' ? f : f.formula)
const valueOf = (v: { t: string; v: unknown } | undefined): CellValue => (v === undefined ? null : v.t === 'e' ? ({ error: v.v } as CellValue) : (v.v as CellValue))
const groups = ['cells', 'ranges', 'converted', 'filled']

/** Whether the writer writes a case's formula of the old kind. */
async function writtenPlain(c: LinksCase): Promise<Record<string, boolean>> {
  const book = linksBook(c)
  const xml = (await openArchive((await writeXlsx({ book, result: compute(book) })).data).text('xl/worksheets/sheet1.xml'))!
  return Object.fromEntries(Object.keys(c.linked.S).map((a) => [a, new RegExp(`<c r="${a}"[^>]*><f>`).test(xml)]))
}

describe.each(PLATFORMS)('the forms of a formula over a link against %s', (_, root) => {
  const set = storedSets<StoredLinks>(`${root}links/`).find((s) => s.name === 'forms')!
  const answer = (id: string) => set.answers[id].reads

  it("the stored cases are the set's, sent as made now, and every file opened", async () => {
    expect(set.cases).toEqual(cases)
    for (const c of cases) expect(set.sent[c.id], c.id).toBe(await partsHash(await linkFile(c)))
    expect(set.rejected).toBeUndefined()
    expect(cases.filter((c) => !set.answers[c.id] || set.answers[c.id].errors).map((c) => c.id)).toEqual([])
  })

  it.each(groups.slice(0, 3))('%s: where the writer writes the old kind, Excel computes the two forms alike', async (group) => {
    const plain = cases.find((c) => c.id === `${group}-plain`)!
    const written = await writtenPlain(plain)
    const parted: string[] = []
    for (const a of Object.keys(plain.linked.S)) {
      const [p, d] = [`${group}-plain`, `${group}-dynamic`].map((id) => answer(id).map((r) => valueOf(r.values[`S!${a}`])))
      if (JSON.stringify(p) === JSON.stringify(d)) continue
      parted.push(formula(plain.linked.S[a]))
      expect(written[a], `${a} ${formula(plain.linked.S[a])}: ${JSON.stringify(p)} of the old kind, ${JSON.stringify(d)} as a dynamic array`).toBe(false)
    }
    expect(parted).toEqual(
      group === 'converted' ? ['[1]Sheet1!B1:B3', '[1]Sheet1!B1:B3*2', 'ABS([1]Sheet1!B1:B3)', 'SUM(IF([1]Sheet1!B1:B3>7,1,0))', 'MAX([1]Sheet1!B1:B3*2)', '[1]Marked!A1:A3'] : [],
    )
  })

  it('every formula with one cell in each reference is written of the old kind', async () => {
    const plain = cases.find((c) => c.id === 'cells-plain')!
    const written = await writtenPlain(plain)
    expect(Object.entries(written).filter(([, yes]) => !yes).map(([a]) => formula(plain.linked.S[a]))).toEqual([])
  })

  it('a fill down that marks the link leaves the two forms alike', () => {
    const [p, d] = ['filled-plain', 'filled-dynamic'].map((id) => answer(id).map((r) => Object.values(r.values).map((v) => valueOf(v))))
    expect(p).toEqual(d)
  })

  it.each(['cells', 'converted'])('%s: the engine computes of the old kind, through its conversion, what Excel computes of it', (group) => {
    const c = cases.find((x) => x.id === `${group}-plain`)!
    const book = linksBook(c)
    const converted = { ...book, sheets: book.sheets.map((s) => ({ ...s, cells: Object.fromEntries(Object.keys(c.linked.S).map((a) => [a, { formula: fromLegacy(formula(c.linked.S[a])) }])) })) }
    const result = compute({ ...converted, sheets: converted.sheets.map((s) => ({ ...s, cells: { ...s.cells, C1: { value: 2 } } })) })
    const otherwise: Record<string, CellValue> = {}
    for (const a of Object.keys(c.linked.S)) {
      const excel = valueOf(answer(c.id)[1].values[`S!${a}`])
      const got = result.sheets[0].cells[a].value ?? 0
      if (JSON.stringify(got) === JSON.stringify(excel)) continue
      otherwise[`${group} S!${a}`] = excel
    }
    expect(otherwise).toEqual(Object.fromEntries(Object.entries(OTHERWISE[root]).filter(([k]) => k.startsWith(group))))
  })
})
