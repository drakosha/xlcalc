/**
 * A sheet's name in a formula shown to a person (`toLocal`, a window's `formulaLocal` and `editText`, the text of
 * `explain`) is in quotes only where Excel puts them, by `quoteSheetIfNeeded` with the locale's booleans: Excel's formula
 * bar shows `=Data!B2*2`, and `'ИСТИНА'!A1` in ru-RU. The stored text keeps every sheet quoted, and `toInvariant` reads a
 * name with quotes and without. In every kind of session, with closures and eval, and over a saved result (`viewportOf`).
 */
import { afterAll, describe, expect, it } from 'vitest'
import { enUS, toInvariant, toLocal, viewportOf, type BookInput, type Execution } from '../src/index.js'
import { ruRU } from '../src/locales/ru-RU.js'
import { canonical, parse } from '../src/syntax.js'
import { CLOCK } from './async-books.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'
import './locales.js'

afterAll(closeCalcs)

/** Stored text, as en-US and ru-RU show it. */
const SHOWN: [string, string, string][] = [
  ["'Data'!B2*2", 'Data!B2*2', 'Data!B2*2'],
  ["SUM('Список'!A1:A3,1.5)", 'SUM(Список!A1:A3,1.5)', 'СУММ(Список!A1:A3;1,5)'],
  // a name that is no plain name, a cell, an R1C1 reference: in quotes everywhere
  ["'Q1 data'!A1+'O''Brien'!A1+'A1'!A1+'R2D2'!A1", "'Q1 data'!A1+'O''Brien'!A1+'A1'!A1+'R2D2'!A1", "'Q1 data'!A1+'O''Brien'!A1+'A1'!A1+'R2D2'!A1"],
  // the locale's booleans: TRUE in en-US, ИСТИНА in ru-RU, an unquoted name of the stored text too
  ["'TRUE'!A1+'ИСТИНА'!A1+ИСТИНА!A2", "'TRUE'!A1+ИСТИНА!A1+ИСТИНА!A2", "TRUE!A1+'ИСТИНА'!A1+'ИСТИНА'!A2"],
  // a reference across sheets in quotes when either end needs them
  ["SUM('Jan:Mar'!A1)+SUM('S1:S3'!A1)+SUM('Jan:Q 4'!A1)", "SUM(Jan:Mar!A1)+SUM('S1:S3'!A1)+SUM('Jan:Q 4'!A1)", "СУММ(Jan:Mar!A1)+СУММ('S1:S3'!A1)+СУММ('Jan:Q 4'!A1)"],
  ["'S'!#REF!+'S'!rate", 'S!#REF!+S!rate', 'S!#ССЫЛКА!+S!rate'],
]

describe('a formula shown: a sheet in quotes where Excel puts them', () => {
  it.each(SHOWN)('%s', (stored, en, ru) => {
    expect([toLocal(stored, enUS), toLocal(stored, ruRU)]).toEqual([
      { ok: true, formula: en },
      { ok: true, formula: ru },
    ])
    // what is shown reads back as the stored formula, and so does the stored text itself
    const same = canonical(parse(stored))
    for (const [text, locale] of [[en, enUS], [ru, ruRU], [stored, enUS]] as const) {
      const back = toInvariant(`=${text}`, locale)
      expect(back.ok && canonical(parse(back.formula)), text).toBe(same)
    }
  })
})

/** S refers to sheets named plainly, by a boolean of ru-RU and with a space. */
const book = (locale: 'en-US' | 'ru-RU'): BookInput => ({
  locale,
  sheets: [
    { name: 'S', cells: { A1: { formula: 'Data!B2*2' }, A2: { formula: "SUM(ИСТИНА!A1,'Q 1'!A1)" } } },
    { name: 'Data', cells: { B2: { value: 2 } } },
    { name: 'ИСТИНА', cells: { A1: { value: 3 } } },
    { name: 'Q 1', cells: { A1: { value: 4 } } },
  ],
})

describe.each(['closures', 'eval'] as const)('a window and explain, execution %s', (execution: Execution) => {
  describe.each(KINDS)('%s', (kind) => {
    it.each([
      ['en-US', 'Data!B2*2', "SUM(ИСТИНА!A1,'Q 1'!A1)", 'S!A1 = 4  ← Data!B2*2\n  Data!B2 = 2'],
      ['ru-RU', 'Data!B2*2', "СУММ('ИСТИНА'!A1;'Q 1'!A1)", 'S!A1 = 4  ← Data!B2*2\n  Data!B2 = 2'],
    ] as const)('%s', async (tag, a1, a2, explained) => {
      const open = await opener(kind)
      const s = await open(book(tag), { clock: CLOCK, execution })
      const opts = { formulaLocal: true, editText: true }
      const want = {
        A1: { formula: "'Data'!B2*2", formulaLocal: a1, editText: `=${a1}` },
        A2: { formula: "SUM('ИСТИНА'!A1,'Q 1'!A1)", formulaLocal: a2, editText: `=${a2}` },
      }
      expect((await s.getViewport('S', 'A1:A2', opts)).cells).toMatchObject(want)
      expect(viewportOf(await s.getSnapshot(), 'S', 'A1:A2', opts).cells).toMatchObject(want)
      expect((await s.explain('S', 'A1')).text).toBe(explained)
      // the tree keeps the stored text
      expect((await s.explain('S', 'A1')).root).toMatchObject({ ref: "'S'!A1", formula: "'Data'!B2*2", children: [{ ref: "'Data'!B2" }] })
      await s.close()
    })
  })
})
