/**
 * Excel's answers to conditional formats (`test/excel/cond/`, `test/excel/win/cond/`, from `run.ts cond <set>`) and their
 * comparison with the engine: what the read cells show after each case (the book as written, or as the reader reads a
 * file changed by hand, after the structure operation, in a session at the day of the run), the rules the reader reads from the file Excel saved against the
 * engine's rules after the case, and what the engine shows on that saved book. Known differences are pinned with
 * their reasons.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import type { StoredCond } from '../bench/coverage/cond.js'
import { condFile, engineCase, rulesOf, type CondCase, type CondCases, type CondSet, type Look } from '../bench/coverage/cond/cases.js'
import { readXlsx } from '../src/xlsx.js'
import './locales.js'

export type CondSetFile = StoredCond & { name: string; dir: URL }

/** The stored sets of a directory (`excel/cond/`, `excel/cond/ru-RU/`, `excel/win/cond/`) with the folders of their saved files. */
export function condSets(sub: string): CondSetFile[] {
  const dir = new URL(`./excel/${sub}`, import.meta.url)
  if (!existsSync(dir)) return []
  return readdirSync(dir)
    .filter((f) => f.endsWith('.json'))
    .sort()
    .map((f) => ({ name: f.slice(0, -5), dir: new URL(`${f.slice(0, -5)}/`, dir), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredCond) }))
}

/** Differences pinned by key: `<set>:<case>` for a case's rules, `<set>:<case>:<sheet>!<cell>` for a cell's look. */
export type Known = Record<string, string>

const show = (l: Look | undefined) => (l ? `${JSON.stringify(l.text)} fill ${l.fill ?? 'none'} font ${l.font}${l.bold ? ' bold' : ''}${l.italic ? ' italic' : ''}` : 'none')

/** Looks that differ only as a platform lays text out: Excel for Windows fits `#` across a column in pixels. */
export const PIXEL_HASHES = (excel: Look, engine: Look | undefined): boolean => !!engine && /^#+$/.test(excel.text) && /^#+$/.test(engine.text) && JSON.stringify({ ...excel, text: '' }) === JSON.stringify({ ...engine, text: '' })

/** The cells whose look differs (but those `alike` takes as the same): `<sheet>!<cell>: Excel …, engine …`. */
function lookDifferences(excel: Record<string, Record<string, Look>>, engine: Record<string, Record<string, Look>>, alike?: (excel: Look, engine: Look | undefined) => boolean): Map<string, string> {
  const out = new Map<string, string>()
  for (const [sheet, cells] of Object.entries(excel))
    for (const [addr, l] of Object.entries(cells)) {
      const e = engine[sheet]?.[addr]
      const x = { text: l.text, fill: l.fill, font: l.font, bold: l.bold, italic: l.italic }
      if (JSON.stringify(e) !== JSON.stringify(x) && !alike?.(x, e)) out.set(`${sheet}!${addr}`, `Excel ${show(l)}, engine ${show(e)}`)
    }
  return out
}

/** The book the engine holds for a case: its book, or for a raw case the book the reader reads from the file sent. */
const caseBook = async (set: CondCases, c: CondCase) => (c.raw ? (await readXlsx(await condFile(set, c))).book : set.books[c.book])

/** Every case of the sets against Excel's answers: looks after the case, rules Excel saved, looks on the saved book. */
export function checkCondSets(sets: CondSetFile[], known: Known, load: (name: string) => Promise<CondSet>, alike?: (excel: Look, engine: Look | undefined) => boolean) {
  for (const set of sets)
    describe(`${set.name}`, () => {
      const clock = { now: `${set.meta.day}T12:00:00Z`, seed: 1 }
      it('Excel loaded every book but those it repairs', () => expect(Object.keys(set.rejected ?? {})).toEqual(set.cases.filter((c) => COND_REPAIRED[`${set.name}:${c.id}`]).map((c) => c.id)))
      for (const c of set.cases as CondCase[]) {
        const key = `${set.name}:${c.id}`
        if (COND_REPAIRED[key])
          it(`${c.id}: the reader reads the file Excel does not load and reports what it leaves out`, async () => {
            const read = await readXlsx(await condFile(set, c))
            expect(read.issues.map((i) => `${i.code} ${i.sheet ?? ''} ${(i.cells ?? []).join(' ')}: ${i.message}`)).toEqual(COND_REPAIRED[key])
          })
        const a = set.answers[c.id]
        if (!a) continue
        // an operation Excel refused leaves the book as it was: the engine refuses it too, and the book shows as before
        const refused = c.op !== undefined && a.error?.startsWith('op') === true
        const after = refused ? { ...c, op: undefined } : c
        if (refused) it(`${c.id}: the engine refuses the operation Excel refused`, () => expect(engineCase(set, c, clock).refused, a.error).toBeDefined())
        it(`${c.id}: the cells show what Excel showed`, async () => {
          const mod = await load(set.name)
          const book = c.make && mod.expectedBook ? undefined : await caseBook(set, c)
          if (!book) return
          const got = lookDifferences(a.looks, engineCase(set, after, clock, book).looks, alike)
          const pinned = [...got].filter(([cell]) => known[`${key}:${cell}`])
          expect([...got].filter(([cell]) => !known[`${key}:${cell}`]).map(([cell, d]) => `${cell}: ${d}`)).toEqual([])
          for (const [cell] of pinned) expect(got.has(cell), `${key}:${cell} pinned`).toBe(true)
        })
        it(`${c.id}: the reader reads from Excel's file the rules the engine holds after the case`, async () => {
          const saved = await readXlsx(new Uint8Array(readFileSync(new URL(`${c.id}.xlsx`, set.dir))))
          if (c.make) return
          const same = JSON.stringify(rulesOf(saved.book)) === JSON.stringify(engineCase(set, after, clock, await caseBook(set, c)).rules)
          expect(same, known[key] ?? 'the rules differ').toBe(!known[key])
        })
        it(`${c.id}: on the book read from Excel's file the engine shows what Excel showed`, async () => {
          const saved = await readXlsx(new Uint8Array(readFileSync(new URL(`${c.id}.xlsx`, set.dir))))
          const got = lookDifferences(a.looks, engineCase({ books: { saved: saved.book }, cases: [] }, { ...c, op: undefined, book: 'saved' }, clock).looks, alike)
          expect([...got].filter(([cell]) => !known[`${key}:${cell}`] && !known[`${key}:saved:${cell}`]).map(([cell, d]) => `${cell}: ${d}`)).toEqual([])
        })
      }
    })
}

/**
 * `<set>:<case id>` → what the reader reports of a file changed by hand that neither Excel loads (the repair window on
 * the Mac, `Workbooks.Open` failing on Windows): a dxf with a gradient beside its pattern, read as its pattern.
 */
export const COND_REPAIRED: Readonly<Record<string, string[]>> = {
  'fills-held:dxf-gradient-pattern': ["conditionalFormat S B1:C1: a conditional format's gradient fill beside its pattern is not in the model"],
}

/** Every pinned difference names a case of the sets. */
export function pinnedCases(sets: CondSetFile[], known: Known): string[] {
  const ids = new Set(sets.flatMap((s) => s.cases.map((c) => `${s.name}:${c.id}`)))
  return Object.keys(known).filter((k) => !ids.has(k.split(':').slice(0, 2).join(':')))
}
