/**
 * Links to other workbooks in the files Excel saved (`test/excel/links/`, `test/excel/win/links/`, from the step
 * `links`): the reader reads every link of every file, with the targets Excel for Mac and Excel for Windows wrote,
 * Excel's alternate addresses beside them, the marks of failed updates and the stored values Excel kept (it drops the
 * cells no formula refers to), and no link is reported as one the model does not have. The set `written` is the files
 * `writeXlsx` wrote: Excel opened them and showed what it showed for the files made by hand.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredLinks } from '../bench/coverage/links.js'
import type { CachedValue } from '../bench/coverage/links/cases.js'
import type { ExternalBook } from '../src/contract.js'
import { readXlsx } from '../src/xlsx.js'
import { XLSX_EN_MESSAGES } from '../src/xlsx/messages.js'

const PLATFORMS = [
  ['Excel for Mac', ''],
  ['Excel for Windows', 'win/'],
] as const

/** The stored sets of a platform. */
function stored(root: string): (StoredLinks & { name: string })[] {
  const dir = new URL(`./excel/${root}links/`, import.meta.url)
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .filter((f) => f.endsWith('.json'))
    .map((f) => ({ name: f.slice(0, -5), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredLinks) }))
}

const saved = (root: string, set: string, id: string) => readXlsx(new Uint8Array(readFileSync(new URL(`./excel/${root}links/${set}/${id}.xlsx`, import.meta.url))))

/** A target's file name, its escapes and Windows' separators undone. */
const fileName = (target: string) => decodeURIComponent(target).replaceAll('\\', '/').split('/').at(-1)

const NOT_IN_MODEL = XLSX_EN_MESSAGES.externalLinks({})

/** A stored value as Excel keeps it among a link's values: #SPILL! and #CALC! as #VALUE!. */
const kept = (v: CachedValue | null | undefined) => (v && typeof v === 'object' && (v.error === 'SPILL' || v.error === 'CALC') ? { error: 'VALUE' } : v)

/** The targets Excel wrote for the case `targets` of `basics-held`: Excel for Mac keeps a relative path it could not find as an absolute one beside it (`alternateUrls`), Excel for Windows writes every path as an absolute one without the drive. */
const TARGETS: Record<string, string[]> = {
  '': ['Book2.xlsx', '/Users/user/Library/Containers/com.microsoft.Excel/Data/tmp/data/Book3.xlsx', 'file:///C:/data/Book4.xlsx', '/Users/x/data/Book5.xlsx', 'https://example.com/share/Book6.xlsx', 'Sub%20Folder/Книга%207.xlsx'],
  'win/': ['/xlcalc/work/links-basics-held-91631-1/Book2.xlsx', '/xlcalc/work/data/Book3.xlsx', '/data/Book4.xlsx', '/Users/x/data/Book5.xlsx', 'https://example.com/share/Book6.xlsx', '/xlcalc/work/links-basics-held-91631-1/Sub%20Folder/Книга%207.xlsx'],
}

/** Where Excel wrote that `Book2.xlsx` is: Excel for Mac as the file had it, Excel for Windows as an absolute path from the folder it opened the file in. */
const book2 = (root: string, folder: string) => (root ? `/xlcalc/work/${folder}/Book2.xlsx` : 'Book2.xlsx')

/** The link of the case `cache` of `basics-held` as Excel saved it after the fill downs and the copy: every sheet marked, a sheet for `[Book2.xlsx]` that `INDIRECT` names, the constant name lost. */
const CACHE: Omit<ExternalBook, 'target'> = {
  sheets: [
    {
      name: 'Sheet1',
      cells: { A1: 1, B1: 'abc', C1: 0.2, D1: 'x', E1: 10, F1: 20, A2: 2, B2: true, C2: 7, D2: 'y', A3: 3, B3: { error: 'NA' }, D3: 'x', B4: { error: 'DIV0' } },
      refreshError: true,
    },
    { name: 'My Sheet', cells: { A1: 11 }, refreshError: true },
    { name: 'Лист 1', cells: { A1: 12 }, refreshError: true },
    { name: 'Book2', cells: {}, refreshError: true },
  ],
  names: [
    { name: 'Rate', ref: "'Sheet1'!$C$1" },
    { name: 'Local', ref: "'Sheet1'!$C$2", sheet: 0 },
    { name: 'Konst', ref: '#REF!' },
    { name: 'Bad', ref: '#REF!' },
  ],
}

describe.each(PLATFORMS)('links in the files %s saved', (_, root) => {
  for (const set of stored(root)) {
    describe(set.name, () => {
      it.each(set.cases.filter((c) => set.answers[c.id]).map((c) => [c.id, c] as const))('%s: every link read, with the values Excel kept', async (_, c) => {
        const r = await saved(root, set.name, c.id)
        expect(r.issues.filter((i) => i.message === NOT_IN_MODEL)).toEqual([])
        const books = r.book.externalBooks!
        // Excel drops a link no formula refers to when it saves, the one without sheets
        const links = c.links.filter((l) => l.sheets.length)
        expect(books.length).toBe(links.length)
        books.forEach((b, i) => {
          const link = links[i]
          expect(fileName(b.target), b.target).toBe(fileName(link.target))
          // Excel keeps the sheets in order and may add one a formula names
          expect(b.sheets.slice(0, link.sheets.length).map((s) => s.name)).toEqual(link.sheets.map((s) => s.name))
          link.sheets.forEach((s, k) => {
            for (const [addr, v] of Object.entries(b.sheets[k].cells ?? {})) expect(v, `[${i + 1}]${s.name}!${addr}`).toEqual(kept(s.cells?.[addr]))
          })
          expect((b.names ?? []).map((n) => n.name)).toEqual((link.names ?? []).map((n) => n.name))
        })
      })
    })
  }

  it('the targets of the forms a link takes, as Excel wrote them', async () => {
    expect((await saved(root, 'basics-held', 'targets')).book.externalBooks!.map((b) => b.target)).toEqual(TARGETS[root])
  })

  it('the stored values, the marks of failed updates and the names of a link after edits', async () => {
    expect((await saved(root, 'basics-held', 'cache')).book.externalBooks).toEqual([{ target: book2(root, 'links-basics-held-91631-1'), ...CACHE }])
  })

  it('data without values, with the mark and without it; a sheet the link had no data of comes back marked', async () => {
    const r = await saved(root, 'refresh-held', 'empty')
    expect(r.book.externalBooks).toEqual([{ target: book2(root, 'links-refresh-held-25311-1'), sheets: [{ name: 'Empty', cells: {} }, { name: 'Marked', cells: {}, refreshError: true }, { name: 'None', cells: {}, refreshError: true }] }])
  })

  it("every formula and name over Excel's links reads without an issue", async () => {
    for (const set of stored(root))
      for (const c of set.cases.filter((x) => set.answers[x.id])) expect((await saved(root, set.name, c.id)).issues.filter((i) => i.code === 'externalLink'), `${set.name}/${c.id}`).toEqual([])
  })
})
