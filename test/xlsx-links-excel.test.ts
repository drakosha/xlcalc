import { withoutView } from './xlsx-fixtures.js'
/**
 * Links to other workbooks in the files Excel saved (`test/excel/links/`, `test/excel/win/links/`, from the step
 * `links`): the reader reads every link of every file, with the targets Excel for Mac and Excel for Windows wrote,
 * Excel's alternate addresses beside them, the marks of failed updates and the stored values Excel kept (it drops the
 * cells no formula refers to), and no link is reported as one the model does not have. The set `written` is the files
 * `writeXlsx` wrote: Excel opened them and showed what it showed for the files made by hand.
 */
import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredLinks } from '../bench/coverage/links.js'
import { linkFile, linksBook, writtenBook, type CachedValue, type LinksAnswer, type LinksRead } from '../bench/coverage/links/cases.js'
import { buildLinksCases, MISSED } from '../bench/coverage/links/written.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import type { BookInput, ExternalBook } from '../src/contract.js'
import { compute } from '../src/index.js'
import { canonical } from '../src/syntax/print.js'
import { parse } from '../src/syntax/parse.js'
import { readXlsx } from '../src/xlsx.js'
import { XLSX_EN_MESSAGES } from '../src/xlsx/messages.js'
import { storedSets } from './stored-sets.js'

const PLATFORMS = [
  ['Excel for Mac', ''],
  ['Excel for Windows', 'win/'],
] as const

/** The stored sets of a platform. */
const stored = (root: string) => storedSets<StoredLinks>(`${root}links/`)

const saved = (root: string, set: string, id: string) => readXlsx(new Uint8Array(readFileSync(new URL(`./excel/${root}links/${set}/${id}.xlsx`, import.meta.url))))

/** A target's file name, its escapes and Windows' separators undone. */
const fileName = (target: string) => decodeURIComponent(target).replaceAll('\\', '/').split('/').at(-1)

const NOT_IN_MODEL = XLSX_EN_MESSAGES.externalLinks({})

/** A stored value as Excel keeps it among a link's values: #SPILL! and #CALC! as #VALUE!. */
const kept = <T extends CachedValue | null | undefined>(v: T): T => (v && typeof v === 'object' && (v.error === 'SPILL' || v.error === 'CALC') ? ({ error: 'VALUE' } as T) : v)

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

/** A value of an answer as a stored value: an error by its code. */
const valueOf = (v: { t: string; v: unknown } | undefined) => (v === undefined ? null : v.t === 'e' ? { error: v.v } : v.v)

/** A read of an answer with the folder of the opened book taken out of its formulas, which name it. */
function unfoldered(a: LinksAnswer, r: LinksRead): LinksRead {
  const out = (m: Record<string, string>) => Object.fromEntries(Object.entries(m).map(([k, f]) => [k, f.replaceAll(a.folder, '<folder>')]))
  return { ...r, formulas: out(r.formulas), local: out(r.local) }
}

const written = buildLinksCases().cases

describe.each(PLATFORMS)('the links writeXlsx wrote against %s', (_, root) => {
  const set = stored(root).find((s) => s.name === 'written')!
  const held = stored(root).find((s) => s.name === 'basics-held')!

  it("the stored cases are the set's, sent as the writer writes them now, and every file opened", async () => {
    expect(set.cases).toEqual(written)
    for (const c of written) expect(set.sent[c.id], c.id).toBe(await partsHash(await linkFile(c)))
    expect(set.rejected).toBeUndefined()
    expect(written.filter((c) => !set.answers[c.id]).map((c) => c.id)).toEqual([])
  })

  it('Excel showed the predicted values, the missed ones as their alternatives', () => {
    const missed: string[] = []
    for (const c of written)
      for (const [key, p] of Object.entries(c.predicted ?? {})) {
        const got = valueOf(set.answers[c.id].reads[1].values[key])
        if (JSON.stringify(got) === JSON.stringify(p)) continue
        missed.push(`${c.id} ${key}`)
        expect(c.alternatives?.[key], `${c.id} ${key}`).toContainEqual(got)
      }
    expect(missed).toEqual(MISSED)
  })

  it.each(written.filter((c) => held.answers[c.id]).map((c) => [c.id, c] as const))('%s: Excel read and computed the written file as the file made by hand, before the edits and after', (_, c) => {
    const [mine, hand] = [set.answers[c.id], held.answers[c.id]]
    mine.reads.forEach((r, i) => expect(unfoldered(mine, r)).toEqual(unfoldered(hand, hand.reads[i])))
  })
})

/** A book with every formula in its canonical text, as the round trip compares formulas. */
function canonicalBook(b: BookInput): BookInput {
  const text = (f: string, union = false) => canonical(parse(f, { union }))
  return { ...b, sheets: b.sheets.map((s) => ({ ...s, cells: Object.fromEntries(Object.entries(s.cells).map(([a, c]) => [a, c.formula === undefined ? c : { ...c, formula: text(c.formula) }])) })), ...(b.names ? { names: b.names.map((n) => ({ ...n, ref: text(n.ref, true) })) } : {}) }
}

describe('the round trip of the books of the set written', () => {
  it.each(written.map((c) => [c.id, c] as const))('%s: the file reads back as the book written, with the same input hash', async (_, c) => {
    // the writer stores a link's #SPILL! and #CALC! as #VALUE!, as Excel keeps them
    const externalBooks: ExternalBook[] = linksBook(c).externalBooks!.map((b) => ({ ...b, sheets: b.sheets.map((s) => (s.cells ? { ...s, cells: Object.fromEntries(Object.entries(s.cells).map(([a, v]) => [a, kept(v)])) } : s)) }))
    const book = { ...writtenBook(c), externalBooks }
    const r = await readXlsx(await linkFile(c))
    expect(r.issues).toEqual([])
    expect(r.book.externalBooks).toEqual(externalBooks)
    expect(r.book.names).toEqual(book.names)
    // the writer quotes every sheet of this book (`'Sheet1'!A1`): formulas compare by their canonical text
    expect(canonicalBook(withoutView(r.book))).toEqual(canonicalBook(book))
    expect(compute(canonicalBook(r.book)).inputHash).toBe(compute(canonicalBook(book)).inputHash)
  })
})
