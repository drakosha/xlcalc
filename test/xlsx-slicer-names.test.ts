/**
 * What the writer writes of a slicer and a timeline, against Excel for Windows and for Mac (`bench/coverage/links/slicer-names.ts`):
 * with a place, the name of its cache beside it, one of a name however the book came by it; without one neither it nor a cache only it
 * used, reported, the PivotTable and its cells kept. `writeKept` leaves a file's own alone.
 */
import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredLinks } from '../bench/coverage/links.js'
import { againFile, buildLinksCases, SAVED, sessionFile } from '../bench/coverage/links/slicer-names.js'
import { writtenBook, type LinksCase } from '../bench/coverage/links/cases.js'
import type { BookInput, StructureOp } from '../src/contract.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import { openKept, writeKept } from '../src/xlsx-keep.js'
import { openArchive } from '../src/xlsx/zip/read.js'
import { pivotBook } from './pivot-books.js'

const CASES = buildLinksCases().cases
const named = (id: string) => CASES.find((c) => c.id === id)!
const BOOK = (c: LinksCase) => writtenBook(c)

/** The file's slicer and timeline parts, and its names but the filter databases, the cache's own value by its name. */
async function held(data: Uint8Array) {
  const a = openArchive(data)
  const wb = (await a.text('xl/workbook.xml'))!
  const names = [...wb.matchAll(/<definedName name="([^"]+)"[^>]*>([^<]*)<\/definedName>/g)].filter((m) => !m[1].startsWith('_xlnm')).map((m) => [m[1], m[2]] as const)
  const count = (re: RegExp) => a.names.filter((n) => re.test(n)).length
  return { names, slicers: count(/^xl\/slicers\/slicer\d+\.xml$/), slicerCaches: count(/^xl\/slicerCaches\//), timelines: count(/^xl\/timelines\/timeline\d+\.xml$/), timelineCaches: count(/^xl\/timelineCaches\//), drawings: count(/^xl\/drawings\/drawing\d+\.xml$/) }
}
const nameList = (h: Awaited<ReturnType<typeof held>>) => h.names.map(([n]) => n).sort()

const SLICER: StructureOp = { op: 'addSlicer', sheet: 'P', slicer: {}, cache: { field: 'Region', pivots: [{ sheet: 'P', name: 'PT1' }] } }
const TIMELINE: StructureOp = { op: 'addTimeline', sheet: 'P', timeline: {}, cache: { field: 'Date', pivots: [{ sheet: 'P', name: 'PT1' }] } }

describe('a slicer and a timeline of a session', () => {
  it('a new slicer takes the name of its cache, with its frame', async () => {
    const r = await sessionFile(BOOK(named('session-slicer')), [SLICER])
    const h = await held(r.data)
    expect(h).toMatchObject({ names: [['Slicer_Region', '#N/A']], slicers: 1, slicerCaches: 1, drawings: 1 })
    expect(r.issues.filter((i) => i.code === 'pivotTable')).toEqual([])
  })

  it('a new timeline takes the name of its cache, with its frame', async () => {
    const h = await held((await sessionFile(BOOK(named('session-timeline')), [TIMELINE])).data)
    expect(h).toMatchObject({ names: [['NativeTimeline_Date', '#N/A']], timelines: 1, timelineCaches: 1, drawings: 1 })
  })

  it('read and written anew, it holds one of each name', async () => {
    const book = BOOK(named('session-both'))
    const first = await sessionFile(book, [SLICER, TIMELINE])
    const second = await againFile(first.data, book)
    const third = await againFile(second.data, book)
    for (const r of [first, second, third]) expect(nameList(await held(r.data))).toEqual(['NativeTimeline_Date', 'Slicer_Region'])
  })

  it('a cache beside a name of the book takes the next name, and each stands in the file', async () => {
    const book = BOOK(named('session-taken'))
    const h = await held((await sessionFile(book, [SLICER])).data)
    expect(h.names).toEqual([['Slicer_Region', "'D'!$A$1"], ['Slicer_Region1', '#N/A']])
  })
})

describe('a slicer and a timeline of a book', () => {
  const placed = (b: BookInput) => {
    b.sheets[1].slicers![0].place = { from: 'H2', to: 'K12' }
    return b
  }
  const written = async (b: BookInput) => {
    const r = await writeXlsx({ book: b }, { values: 'omit' })
    return { h: await held(r.data), issues: r.issues.filter((i) => i.code === 'pivotTable'), data: r.data }
  }

  it('without a place neither it nor its cache is written, reported by its name; the PivotTable and its cells stay', async () => {
    const { h, issues, data } = await written(pivotBook())
    expect(h).toEqual({ names: [], slicers: 0, slicerCaches: 0, timelines: 0, timelineCaches: 0, drawings: 0 })
    expect(issues.map((i) => [i.sheet, i.name])).toEqual([['Pivot', 'Region']])
    const back = (await readXlsx(data)).book
    expect(back.sheets[1].pivotTables![0].name).toBe('PivotTable1')
    expect(back.sheets[1].cells.A4).toEqual({ value: 'North' })
    expect([back.slicerCaches, back.sheets[1].slicers]).toEqual([undefined, undefined])
  })

  it('a timeline without a place goes as a slicer does', async () => {
    const b = pivotBook()
    b.slicerCaches = []
    b.sheets[1].slicers = []
    b.timelineCaches = [{ name: 'NativeTimeline_Month', field: 'Month', pivots: [{ sheet: 'Pivot', name: 'PivotTable1' }], cache: 1 }]
    b.sheets[1].timelines = [{ name: 'Month', cache: 'NativeTimeline_Month' }]
    const { h, issues } = await written(b)
    expect([h.timelines, h.timelineCaches, h.names]).toEqual([0, 0, []])
    expect(issues.map((i) => [i.sheet, i.name])).toEqual([['Pivot', 'Month']])
  })

  it('a cache another slicer shows stays, with its name', async () => {
    const b = placed(pivotBook())
    b.sheets[1].slicers!.push({ name: 'Region 1', cache: 'Slicer_Region' })
    const { h, issues } = await written(b)
    expect([h.slicers, h.slicerCaches, nameList(h)]).toEqual([1, 1, ['Slicer_Region']])
    expect(issues.map((i) => i.name)).toEqual(['Region 1'])
  })

  it('a cache nothing shows is not written, reported by its name', async () => {
    const b = pivotBook()
    b.sheets[1].slicers = []
    const { h, issues } = await written(b)
    expect([h.slicerCaches, h.names]).toEqual([0, []])
    expect(issues.map((i) => i.name)).toEqual(['Slicer_Region'])
  })

  it.each([
    ['the book’s own', 'Slicer_Region'],
    ['in another case', 'SLICER_REGION'],
  ])('a name the book holds, %s, is not added twice', async (_, name) => {
    const { h, issues } = await written({ ...placed(pivotBook()), names: [{ name, ref: '#N/A' }] })
    expect(h.names.map(([n]) => n)).toEqual([name])
    expect(issues).toEqual([])
  })

  it.each([
    ['a cell', 'Data!$A$1'],
    ['what the writer leaves out', '1+'],
  ])('a cache whose name the workbook holds for %s is not written, as Excel refuses that file', async (_, ref) => {
    const { h, issues } = await written({ ...placed(pivotBook()), names: [{ name: 'Slicer_Region', ref }] })
    expect([h.slicers, h.slicerCaches, h.drawings]).toEqual([0, 0, 0])
    expect(issues.map((i) => i.name)).toEqual(['Slicer_Region'])
  })

  it('a timeline’s cache of such a name goes too', async () => {
    const b = pivotBook()
    b.slicerCaches = []
    b.sheets[1].slicers = []
    b.timelineCaches = [{ name: 'NativeTimeline_Month', field: 'Month', pivots: [{ sheet: 'Pivot', name: 'PivotTable1' }], cache: 1 }]
    b.sheets[1].timelines = [{ name: 'Month', cache: 'NativeTimeline_Month', place: { from: 'H2', to: 'K12' } }]
    b.names = [{ name: 'NativeTimeline_Month', ref: 'Data!$A$1' }]
    const { h, issues } = await written(b)
    expect([h.timelines, h.timelineCaches]).toEqual([0, 0])
    expect(issues.map((i) => i.name)).toEqual(['NativeTimeline_Month'])
  })

  it('a sheet’s own name of the cache does not stand for the workbook’s', async () => {
    const { h } = await written({ ...placed(pivotBook()), names: [{ name: 'Slicer_Region', ref: 'Data!$A$1', scope: 'Pivot' }] })
    expect(h.names).toEqual([['Slicer_Region', "'Data'!$A$1"], ['Slicer_Region', '#N/A']])
  })

})

describe('Excel holds what the writer wrote', () => {
  it('a slicer and a timeline stay on saving, on both Excels, of a session', async () => {
    for (const dir of ['win/', '']) {
      const a = JSON.parse(readFileSync(new URL(`./excel/${dir}links/slicer-names.json`, import.meta.url), 'utf8')) as StoredLinks
      expect(Object.keys(a.rejected ?? {}).filter((id) => id.startsWith('session') || id.startsWith('no-place'))).toEqual([])
      for (const [id, want] of [['session-slicer', [1, 1, 0, 0]], ['session-timeline', [0, 0, 1, 1]], ['no-place-slicer', [0, 0, 0, 0]], ['no-place-timeline', [0, 0, 0, 0]]] as const) {
        const h = await held(new Uint8Array(readFileSync(new URL(`./excel/${dir}links/slicer-names/${id}.xlsx`, import.meta.url))))
        expect([h.slicers, h.slicerCaches, h.timelines, h.timelineCaches], `${dir}${id}`).toEqual(want)
      }
      expect(Object.keys(a.rejected ?? {})).toEqual(expect.arrayContaining(['no-name-slicer', 'no-name-timeline']))
    }
  })
})

describe('writeKept', () => {
  it('keeps the name of a file’s cache and the cache, with a cell changed beside them', async () => {
    const data = new Uint8Array(readFileSync(new URL('./excel/win/pivot/slicer/row-two.xlsx', import.meta.url)))
    const { book } = await openKept(data)
    book.sheets[0].cells.H1 = { value: 1 }
    const out = await writeKept(data, { book })
    expect(await held(out.data)).toMatchObject({ names: [['Slicer_Region', '#N/A']], slicers: 1, slicerCaches: 1, drawings: 1 })
  })
})

/** What Excel's saved file holds. */
async function savedBy(dir: string, set: StoredLinks, c: LinksCase) {
  const file = new URL(`./excel/${dir}links/slicer-names/${c.id}.xlsx`, import.meta.url)
  if (!set.answers[c.id] || !existsSync(file)) return null
  return held(new Uint8Array(readFileSync(file)))
}

describe.each([
  ['Windows', 'win/'],
  ['Mac', ''],
])('Excel for %s', (_, dir) => {
  const url = new URL(`./excel/${dir}links/slicer-names.json`, import.meta.url)
  const set = existsSync(url) ? (JSON.parse(readFileSync(url, 'utf8')) as StoredLinks) : null
  it.skipIf(!set)('the cases the prediction missed', async () => {
    const missed: string[] = []
    for (const c of CASES) {
      const asked = set!.answers[c.id] || set!.rejected?.[c.id]
      if (!asked || !!c.refused !== !!set!.rejected?.[c.id]) {
        if (asked) missed.push(c.id)
        continue
      }
      const h = await savedBy(dir, set!, c)
      const want = SAVED[c.id]
      if (h && want && JSON.stringify([h.slicers, h.timelines, [...new Set(nameList(h))]]) !== JSON.stringify([want.slicers, want.timelines, [...new Set(want.names)].sort()])) missed.push(c.id)
    }
    expect(missed).toEqual(MISSED[dir])
  })
})

/** The cases the committed prediction missed, by Excel: its folder of answers. */
const MISSED: Record<string, string[]> = {
  // a name of another value costs the file; Excel holds a sheet's slicers in one part, not one each, and takes the cache's spelling of the name
  'win/': ['session-two', 'book-name-case', 'book-name-other', 'placed-apart'],
  '': [],
}
