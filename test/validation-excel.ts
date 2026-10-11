/**
 * Excel's answers to data validation and outlines (`test/excel/valid/[<tag>/]<set>.json` and the files Excel saved beside
 * them, from `run.ts valid <set>`; Excel for Windows's in `test/excel/win/valid/`) and their comparison with the engine
 * and the reader: each cell's rule over a window with its formulas as at the cell, whether its value meets the rule, the
 * lines' outline levels, hidden and collapsed marks, the summary settings.
 */
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { validFile, type StoredValid } from '../bench/coverage/valid.js'
import { bookViews, excelView, made, viewDifferences, WINDOW, type SheetView, type ValidSet } from '../bench/coverage/valid/cases.js'
import { partsHash } from '../bench/coverage/xlsx/sent.js'
import { editStructure } from '../src/index.js'
import { readXlsx, writeXlsx } from '../src/xlsx.js'
import './locales.js'

export type ValidSetFile = StoredValid & { name: string; tag: string; dir: URL }

/** The stored sets of a directory (`excel/valid/`, `excel/win/valid/`), en-US's and each locale's. */
export function validSets(sub: string): ValidSetFile[] {
  const root = new URL(`./excel/${sub}`, import.meta.url)
  if (!existsSync(root)) return []
  const dirs: [string, URL][] = [['en-US', root], ...readdirSync(root).filter((d) => /^[a-z]{2}-[A-Z]{2}$/.test(d)).map((d): [string, URL] => [d, new URL(`${d}/`, root)])]
  return dirs.flatMap(([tag, dir]) =>
    readdirSync(dir)
      .filter((f) => f.endsWith('.json'))
      .sort()
      .map((f) => ({ name: f.slice(0, -5), tag, dir: new URL(`${f.slice(0, -5)}/`, dir), ...(JSON.parse(readFileSync(new URL(f, dir), 'utf8')) as StoredValid) })),
  )
}

/** What a comparison leaves out: the formulas of literal lists (`lists`), or every formula (`formulas`), where Excel's object model shows them otherwise than as text the engine reads. */
export type Loose = 'lists' | 'formulas'

/** Views without what a comparison leaves out. */
function loosened(views: SheetView[], loose: Loose | undefined): SheetView[] {
  if (!loose) return views
  return views.map((v) => ({
    ...v,
    rules: Object.fromEntries(
      Object.entries(v.rules).map(([addr, r]) => {
        if (loose === 'lists' && !(r.type === 'list' && r.formula1?.startsWith('"'))) return [addr, r]
        const { formula1: _, formula2: __, ...rest } = r
        return [addr, rest]
      }),
    ),
  }))
}

/** The engine's book after a case: the book with what Excel made, the operation applied; or its refusal. */
function engineAfter(set: ValidSetFile, id: string) {
  const c = set.cases.find((x) => x.id === id)!
  const book = made(set.books[c.book], c.make)
  return c.op ? editStructure(book, c.op) : { ok: true as const, book }
}

/** `<set>:<case id>` (a locale's with its tag first, `ru-RU/ru:typed`) → lines pinned, with their reasons where they are made. */
export type Pins = Readonly<Record<string, string[]>>

export interface ValidPins {
  /** Between Excel and the engine. */
  known: Pins
  /** Between Excel and the reader of the file Excel saved. */
  read: Pins
  /** What the reader reports of the file Excel saved (`code sheet cells`). */
  issues: Pins
  /** What a comparison leaves out, by locale. */
  loose?: Readonly<Record<string, Loose>>
}

const keyOf = (set: ValidSetFile, id: string) => `${set.tag === 'en-US' ? '' : `${set.tag}/`}${set.name}:${id}`

/**
 * The tests of stored sets: each case answered for the current books and writer, Excel loaded each book, Excel read what
 * the engine holds, and the reader reads Excel's file as Excel read it, reporting only what `pins.issues` names.
 */
export function checkValidSets(sets: ValidSetFile[], pins: ValidPins, from: (name: string) => Promise<ValidSet>) {
  for (const set of sets) {
    const window = set.window ?? WINDOW
    const loose = pins.loose?.[set.tag]
    describe(`valid set ${set.tag === 'en-US' ? '' : `${set.tag}/`}${set.name}`, () => {
      it('the stored answers are for the current cases and books', async () => {
        const built = (await from(set.name)).buildValidCases()
        expect(set.cases).toEqual(built.cases)
        expect(set.books).toEqual(built.books)
        expect(built.cases.filter((c) => !set.answers[c.id] && !set.rejected?.[c.id])).toEqual([])
        for (const c of built.cases) expect(set.sent[c.id], `${c.id}: the writer changed since Excel saw the file`).toBe(await partsHash((await writeXlsx({ book: built.books[c.book] })).data))
      })
      for (const c of set.cases) {
        const key = keyOf(set, c.id)
        it(`${c.id}: Excel loaded the book`, () => expect(set.rejected?.[c.id]).toBeUndefined())
        const a = set.answers[c.id]
        if (!a) continue
        const excel = () => loosened(a.sheets.map((s) => excelView(s, set.tag)), loose)
        it(`${c.id}: Excel read what the engine holds`, () => {
          const r = engineAfter(set, c.id)
          if (a.error && c.op && !a.error.startsWith('make')) return void expect(r.ok, `Excel refused: ${a.error}`).toBe(false)
          const lines = r.ok ? viewDifferences(excel(), loosened(bookViews(r.book, window), loose)) : [`engine refused (${r.code}): ${r.message}`]
          expect(lines).toEqual(pins.known[key] ?? [])
        })
        it(`${c.id}: the reader reads Excel's file as Excel read it`, async () => {
          const read = await readXlsx(new Uint8Array(readFileSync(new URL(`${c.id}.xlsx`, set.dir))))
          expect(viewDifferences(excel(), loosened(bookViews(read.book, window), loose))).toEqual(pins.read[key] ?? [])
          expect(read.issues.map((i) => `${i.code} ${i.sheet ?? ''} ${(i.cells ?? []).join(' ')}`.trim())).toEqual(pins.issues[key] ?? [])
        })
      }
    })
  }
}

const show = (v: unknown) => JSON.stringify(v)

/** The rule of `long` (`bench/coverage/valid/files.ts`) on A1:A3 as the book holds it and as Excel reads it, its texts cut to Excel's 32, 255, 32 and 225 characters. */
function cutTexts(): string[] {
  const long = (n: number) => 'abcdefghij'.repeat(n / 10)
  const base = { type: 'any', allowBlank: true, showPrompt: true, showError: true, errorStyle: 'stop' }
  const book = { ...base, promptTitle: long(40), prompt: long(300), errorTitle: long(40), error: long(300) }
  const excel = { ...base, promptTitle: book.promptTitle.slice(0, 32), prompt: book.prompt.slice(0, 255), errorTitle: book.errorTitle.slice(0, 32), error: book.error.slice(0, 225) }
  return ['A1', 'A2', 'A3'].map((a) => `S rule ${a}: Excel ${show(excel)}, book ${show(book)}`)
}

/** A rule's line where Excel's object model shows a formula otherwise than as its text: `shown`, the engine's `text`; `after`, the fields after the formula. */
const shownAs = (sheet: string, cells: string[], rule: Record<string, unknown>, shown: string, text: string, after: Record<string, unknown> = {}) =>
  cells.map((a) => `${sheet} rule ${a}: Excel ${show({ ...rule, formula1: shown, ...after })}, book ${show({ ...rule, formula1: text, ...after })}`)

const base = { allowBlank: true, showPrompt: true, showError: true, errorStyle: 'stop' }

/** The object model shows a custom rule's text constant `"TRUE"` as `TRUE`, without its quotes. */
const textConstant = shownAs('S', ['M2'], { type: 'custom', ...base }, 'TRUE', '"TRUE"')

/** The object model shows a date rule's bound 0 as `1/0/1900`, which the comparison does not read as a date. */
const dayZero = shownAs('S', ['B2', 'B3'], { type: 'date', ...base, allowBlank: false, operator: 'greaterThan' }, '1/0/1900', '0')

/**
 * Excel saves a list's reference to a deleted sheet as `#REF!` alone (`<xm:f>#REF!</xm:f>` in the sheet's extension), and
 * shows it as `#REF!$A$1:$A$4`, as the engine keeps it; both are lost references.
 */
const lostSheet = shownAs('S', ['E2', 'E3', 'E5'], { type: 'list', ...base }, '#REF!$A$1:$A$4', '#REF!', { dropdown: true })

/**
 * A list item holding a comma, typed through the object model in ru-RU, where the comma is no separator: Excel keeps
 * `p,q` as one item while the book is open (`q` fails, `p,q` meets), and saves the list as `"p,q"`, which a file reads as
 * two items, with the one item only in an extension (`x12ac:list`) the reader reports.
 */
const commaItem = ['S value D2: Excel fails, book meets', 'S value D3: Excel meets, book fails']

/** The pins both Excels share. */
export const VALID_PINS: ValidPins = {
  known: {
    // the writer cuts a rule's texts to Excel's limits: Excel does not load longer ones
    'files:long': cutTexts(),
    'entry-held:values': textConstant,
    'held-3:values': dayZero,
    'ru-RU/ru:typed': commaItem,
  },
  read: {
    'struct-held:delete-l': lostSheet,
    'entry-held:values': textConstant,
    'held-3:values': dayZero,
    'ru-RU/ru:typed': commaItem,
  },
  issues: {
    // Excel saves a moved rule over a cell beside the rule the cell had: the earlier one, which Excel shows, is kept
    'struct-held-2:move-rule-onto-rule': ['dataValidation S F2:F6'],
    'struct-held-2:move-rules-onto-rules': ['dataValidation S F2:F6'],
    'ru-RU/ru:typed': ['dataValidation S D2:D3'],
  },
}
