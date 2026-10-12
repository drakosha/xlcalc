/**
 * Links to other workbooks against Excel for Mac and Excel for Windows (`test/excel/links/`, `test/excel/win/links/`,
 * from the step `links`): each case's book computed from its links' stored values gives what Excel read after opening
 * the file without updating its links; after the case's operations in a session (a fill down, a copy, rows inserted, a
 * sheet renamed) it gives what Excel read after them, and its formulas are those Excel showed, shown as Excel shows a
 * closed book's references with the opened book's folder as the links' path. The predictions in the sets were
 * committed before Excel ran them: Excel's answer is the predicted value or one of its alternatives, but where `MISSED`
 * says why not.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredLinks } from '../bench/coverage/links.js'
import { linkFile, linksBook, type CachedValue, type LinksCase, type LinksCases, type LinksOp } from '../bench/coverage/links/cases.js'
import { samePin } from '../bench/coverage/compare.js'
import type { Val } from '../bench/coverage/formulajs.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import { compute, createSession, toLocal, type BookInput, type CellValue, type CellRef, type Session } from '../src/index.js'

const PLATFORMS = [
  ['Excel for Mac', ''],
  ['Excel for Windows', 'win/'],
] as const

/** Functions the engine does not have (`#NAME?` in any book): Excel's answer over a closed book is in the comment. */
const UNSUPPORTED: Record<string, string> = {
  'basics-held/cache/S!A22': 'INDIRECT into a closed book is #REF!',
  'basics-held/cache/S!A23': 'OFFSET over a closed book is #VALUE!',
  'basics-held/cache/S!A37': 'INDIRECT into a closed book is #REF!',
  'written/cache/S!A22': 'INDIRECT into a closed book is #REF!',
  'written/cache/S!A23': 'OFFSET over a closed book is #VALUE!',
  'written/cache/S!A37': 'INDIRECT into a closed book is #REF!',
  'functions-held/functions/S!A11': 'CELL over a closed book gives no text',
}

/**
 * Predictions Excel did not confirm, and why. A cell a link does not store was predicted blank: it is, until a formula
 * with that link is entered, which marks the link's sheets as Excel cannot read them (`refreshError`), and `#REF!` after.
 */
const MISSED: Record<string, string> = {
  'basics-held/cache/S!A3': 'SUM over a whole column of a link marked by the fill down: #REF!',
  'basics-held/cache/S!A24': 'SUM over a whole row of a link marked by the fill down: #REF!',
  'basics-held/cache/S!A28': 'COUNTA counts the #REF! of a cell the marked link does not store',
  'basics-held/cache/S!A35': 'AVERAGE over a whole column of a link marked by the fill down: #REF!',
  'written/cache/S!A3': 'SUM over a whole column of a link marked by the fill down: #REF!',
  'written/cache/S!A24': 'SUM over a whole row of a link marked by the fill down: #REF!',
  'written/cache/S!A28': 'COUNTA counts the #REF! of a cell the marked link does not store',
  'written/cache/S!A35': 'AVERAGE over a whole column of a link marked by the fill down: #REF!',
  ...Object.fromEntries(['A1', 'A2', 'A3', 'A4'].map((a) => [`written/errors/S!${a}`, 'a stored #SPILL! or #CALC! a file holds as #VALUE!, which Excel shows'])),
}

/** The stored sets of a platform. */
function stored(root: string): (StoredLinks & { name: string })[] {
  const dir = new URL(`./excel/${root}links/`, import.meta.url)
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .filter((f) => f.endsWith('.json'))
    .map((f) => ({ name: f.slice(0, -5), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredLinks) }))
}

const valOf = (v: CellValue | undefined): Val | null => (v === null || v === undefined ? null : ((typeof v === 'object' ? { t: 'e', v: v.error } : { t: typeof v === 'number' ? 'n' : typeof v === 'string' ? 's' : 'b', v }) as Val))
const cachedVal = (v: CachedValue | null): Val | null => valOf(v as CellValue)

/** The engine's value against Excel's: equal, or an empty cell against the 0 a mirror reads of it. */
const agrees = (excel: Val, got: Val | null) => (got ? samePin(excel, got) : excel.t === 'n' && excel.v === 0)

/** The value of `Sheet!A1` in a book's result. */
function valueIn(book: BookInput, result: { sheets: { cells: Record<string, { value: CellValue }> }[] }, key: string): Val | null {
  const [sheet, cell] = key.split('!')
  return valOf(result.sheets[book.sheets.findIndex((s) => s.name === sheet)]?.cells[cell]?.value)
}

/** A case's operations in a session, as Excel took them (a fill down is a copy of its first row over the rest); how rows inserted moved the cells. */
function applyOps(s: Session, ops: readonly LinksOp[] = []): ((ref: CellRef) => CellRef)[] {
  return ops.flatMap((o) => {
    const r =
      o.op === 'fillDown'
        ? s.editStructure({ op: 'copyRange', from: { sheet: o.sheet, range: o.to.split(':')[0] }, to: { sheet: o.sheet, range: `${o.to.replace(/\d+:/, (m) => `${Number(m.slice(0, -1)) + 1}:`)}` } })
        : o.op === 'copyRange'
          ? s.editStructure({ op: 'copyRange', from: { sheet: o.sheet, range: o.range }, to: { sheet: o.sheet, range: o.to } })
          : o.op === 'insertRows'
            ? s.editStructure({ op: 'insertRows', sheet: o.sheet, row: Number(o.range.replace(/^[A-Z]+/, '')) })
            : s.editStructure({ op: 'renameSheet', sheet: o.sheet, name: o.to })
    if (!r.ok) throw new Error(`${o.op}: ${r.message}`)
    return o.op === 'insertRows' ? [(ref: CellRef) => r.mapCell(ref) ?? ref] : []
  })
}

/** The links of a book with each target in the folder Excel opened it from (`sep` its separator), as Excel shows a closed book. */
const inFolder = (book: BookInput, folder: string, sep: string): BookInput => ({ ...book, externalBooks: book.externalBooks!.map((b) => ({ ...b, target: `${folder}${sep}${b.target}` })) })

describe.each(PLATFORMS)('links against %s', (_, root) => {
  it('every set is answered', () => expect(stored(root).map((s) => s.name)).toEqual(['basics-held', 'edits-held', 'first', 'functions-held', 'refresh-held', 'written']))
  for (const set of stored(root)) {
    describe(set.name, async () => {
      const built = ((await import(`../bench/coverage/links/${set.name}.ts`)) as { buildLinksCases(): LinksCases }).buildLinksCases()
      const at = (c: LinksCase, key: string) => `${set.name}/${c.id}/${key}`

      it('the stored cases are the generator’s, sent as `linkFile` writes them now, and every one opened', async () => {
        expect(set.cases).toEqual(built.cases)
        for (const c of built.cases) expect(set.sent[c.id]).toBe(await partsHash(await linkFile(c)))
        expect(set.rejected ?? {}).toEqual({})
      })

      it('after opening, the engine computes what Excel read', () => {
        for (const c of built.cases) {
          const book = linksBook(c)
          const result = compute(book, { display: false })
          for (const [key, excel] of Object.entries(set.answers[c.id].reads[0].values)) {
            const got = valueIn(book, result, key)
            if (UNSUPPORTED[at(c, key)]) expect(got, at(c, key)).toEqual({ t: 'e', v: 'NAME' })
            else expect(agrees(excel, got), `${at(c, key)}: engine ${JSON.stringify(got)}, Excel ${JSON.stringify(excel)}`).toBe(true)
          }
        }
      })

      it('after the case’s operations, a session gives what Excel read and the formulas it showed, and its hash is the full computation’s', () => {
        for (const c of built.cases) {
          const s = createSession(linksBook(c))
          const maps = applyOps(s, c.ops)
          const book = s.getBook()
          const result = s.getResult()
          expect(compute(book).hash, c.id).toBe(result.hash)
          const answer = set.answers[c.id].reads[1]
          for (const [key, excel] of Object.entries(answer.values)) {
            // the Mac reads through mirror formulas, which rows inserted move with the cells they read
            const [sheet, cell] = key.split('!')
            const moved = root === '' ? maps.reduce((ref, m) => m(ref), { sheet, cell }) : { sheet, cell }
            const got = valueIn(book, result, `${moved.sheet}!${moved.cell}`)
            if (UNSUPPORTED[at(c, key)]) expect(got, at(c, key)).toEqual({ t: 'e', v: 'NAME' })
            else expect(agrees(excel, got), `${at(c, key)}: session ${JSON.stringify(got)}, Excel ${JSON.stringify(excel)}`).toBe(true)
          }
          // targets as Excel resolves them from the opened book's folder are its file system's: `targets` is shown as stored
          if (c.id === 'targets') continue
          const shown = inFolder(book, set.answers[c.id].folder, root === '' ? '/' : '\\')
          for (const [key, excel] of Object.entries(answer.formulas)) {
            const [sheet, cell] = key.split('!')
            const formula = book.sheets.find((x) => x.name === sheet)!.cells[cell]!.formula!
            const t = toLocal(formula, 'en-US', { externalBooks: shown.externalBooks })
            expect(t.ok && `=${t.formula}`, at(c, key)).toBe(excel)
          }
        }
      })

      it('Excel gave the predicted value or one of its alternatives', () => {
        for (const c of built.cases) {
          for (const [key, p] of Object.entries(c.predicted ?? {})) {
            if (MISSED[at(c, key)] || (root === '' && c.id === 'inserted')) continue
            const excel = set.answers[c.id].reads[1].values[key]
            const ok = [p, ...(c.alternatives?.[key] ?? [])].some((v) => (cachedVal(v) ? excel && samePin(cachedVal(v)!, excel) : excel === undefined || (excel.t === 'n' && excel.v === 0)))
            expect(ok, `${at(c, key)}: predicted ${JSON.stringify(p)}, Excel ${JSON.stringify(excel)}`).toBe(true)
          }
        }
      })
    })
  }
})

describe('links shown as stored, where Excel resolves the target from the opened book’s folder', () => {
  it('a file name, a relative path, file:///, an absolute path, a web address, a folder and a file with a space', () => {
    const externalBooks = ['Book2.xlsx', '../data/Book3.xlsx', 'file:///C:\\data\\Book4.xlsx', '/Users/x/data/Book5.xlsx', 'https://example.com/share/Book6.xlsx', 'Sub Folder/Книга 7.xlsx'].map((target) => ({ target, sheets: [{ name: 'Sheet1' }] }))
    const shown = [1, 2, 3, 4, 5, 6].map((k) => toLocal(`'[${k}]Sheet1'!A1`, 'en-US', { externalBooks }))
    expect(shown.map((t) => t.ok && t.formula)).toEqual([
      '[Book2.xlsx]Sheet1!A1',
      "'../data/[Book3.xlsx]Sheet1'!A1",
      "'C:\\data\\[Book4.xlsx]Sheet1'!A1",
      "'/Users/x/data/[Book5.xlsx]Sheet1'!A1",
      "'https://example.com/share/[Book6.xlsx]Sheet1'!A1",
      "'Sub Folder/[Книга 7.xlsx]Sheet1'!A1",
    ])
  })
})
