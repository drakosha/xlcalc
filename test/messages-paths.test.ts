/**
 * Every way a message reaches a person gives it in the language of the book or of the locale called with: a formula that
 * does not read (`validateLocalFormula`, `toInvariant`, `toLocal`, `parseEntry`), a refused structure operation
 * (`editStructure`, and a session's, synchronous, in the same thread and in threads, in both execution modes), and the
 * text of `explain`. Each case names the message of the catalog it gives: in ru-RU its Russian text, in a locale
 * without messages of its own (and in an en-US book) the English one, at the same place. Together the cases reach
 * every message of the catalog but the rules of names, which only `validateName` gives, in English.
 */
import { afterAll, describe, expect, it } from 'vitest'
import { closeTyped } from '../src/entry/close.js'
import { createInlineCalc, createSession, parseEntry, editStructure, toInvariant, toLocal, validateLocalFormula, type BookInput, type Locale, type SheetInput, type StructureOp, type ValidateContext } from '../src/index.js'
import { ruRU } from './locales.js'
import { EN_MESSAGES } from '../src/locales/en-US.js'
import { EDIT_MESSAGES } from '../src/locale/edit-messages.js'
import { say, type EngineMessage, type MessageKey } from '../src/locale/messages.js'
import { closedAfter, pool } from './fixtures/pool.js'
import { closeCalcs, KINDS, opener } from './session-kinds.js'

const reached = new Set<MessageKey>()
const ru = (m: EngineMessage) => say('ru-RU', m)
const en = (m: EngineMessage) => say('en-US', m)
/** ru-RU's formula text without messages of its own: English messages, the same places. */
const english: Locale = { ...ruRU, tag: 'xx-YY', messages: undefined }

const TABLES = { tables: [{ name: 'Т', columns: ['Кол'] }] }

/** Text typed in ru-RU, the message it gives, and the check's context. */
const TYPED: [string, EngineMessage, ValidateContext?][] = [
  ['=$A$0', { key: 'rowOutOfRange', row: '0' }],
  ['=$XFE$1', { key: 'columnOutOfRange', column: 'XFE' }],
  ['=1E400', { key: 'numberOutOfRange', number: '1E400' }],
  ['=A:B#', { key: 'spillAfterRange' }],
  ['=Лист1:Лист3!A1#', { key: 'spillAcrossSheets' }],
  ['="abc', { key: 'unterminatedString' }],
  ['=#ФУ', { key: 'unknownErrorLiteral' }],
  ['=[1]Лист!A1', { key: 'externalReference' }],
  ['=Лист!СУММ(1)', { key: 'referenceAfterSheet' }],
  ["='Лист'", { key: 'quotedSheetAlone' }],
  ['=1~2', { key: 'unexpectedCharacter', char: '~' }],
  ['=', { key: 'emptyFormula' }],
  [`=${'1+'.repeat(4096)}1`, { key: 'tooLong', max: 8192 }],
  [`=${'('.repeat(65)}1${')'.repeat(65)}`, { key: 'tooDeep', max: 64 }],
  ['=1+', { key: 'unexpectedEnd' }],
  ['=1)', { key: 'unexpected', found: { t: ')' } }],
  ['=1 СУММ(2)', { key: 'unexpected', found: { t: 'fn', v: 'SUM' } }],
  ['=СУММ(1', { key: 'expected', char: ')', found: { t: 'eof' } }],
  ['=1;2', { key: 'commaOutside' }],
  ['={1 2}', { key: 'arraySeparator', found: { t: 'num' } }],
  ['={1;2:3}', { key: 'arrayRows' }],
  ['={1;+}', { key: 'arrayValue', found: { t: 'op', v: '+' } }],
  ['=LET(x;1;X;2;x)', { key: 'letTwice', name: 'X' }],
  ['=LET(1;2;3)', { key: 'letName', position: 1 }],
  ['=LET(1)', { key: 'letEmpty' }],
  [`=LET(${Array.from({ length: 128 }, (_, i) => `x_${i};${i}`).join(';')};1)`, { key: 'letTooMany', max: 254 }],
  ['=Т[Кол', { key: 'specUnterminated' }],
  ['=Т[[Ко[л]]', { key: 'specBracketInColumn' }],
  ['=Т[#Нет]', { key: 'specUnknown', word: '#Нет' }],
  ['=Т[[А]:Б]', { key: 'specSpan' }],
  ['=Т[[#Данные];Кол]', { key: 'specOpen' }],
  ['=Т[[А];[Б]]', { key: 'specTwoColumns' }],
  ['=Т[[#Данные];[#Данные]]', { key: 'specTwice', word: '#Data' }],
  ['=Т[[А] Б]', { key: 'specClose' }],
  ['=Т[[#Все];[#Данные]]', { key: 'specTogether', words: ['#All', '#Data'] }],
  ['=BAL(A1)', { key: 'externalLiterals', fn: 'BAL' }, { externalFunctions: ['BAL'] }],
  ['=ОКРУГЛ(1;2;3)', { key: 'argumentCount', fn: 'ROUND', counts: { min: 2, max: 2 }, got: 3 }],
  ['=СУММЕСЛИ(A1:A3;">1";)', { key: 'argumentLeftOut', fn: 'SUMIF', position: 3 }],
  ['=[@Кол]', { key: 'tableRefOutside' }, TABLES],
  ['=Х[Кол]', { key: 'unknownTable', table: 'Х' }, TABLES],
  ['=Т[Нет]', { key: 'unknownColumn', table: 'Т', column: 'Нет' }, TABLES],
  ['=ФУ(1)', { key: 'unknownFunction', fn: 'ФУ' }],
  ['=ИСТОРИЯАКЦИЙ(1)', { key: 'unknownFunction', fn: 'STOCKHISTORY' }],
  ['=Нет!A1', { key: 'unknownSheet', sheet: 'Нет' }, { sheets: ['Лист'] }],
  ['=[1]Лист1:Лист3!A1', { key: 'externalAcrossSheets' }],
  ['=[1]Лист!A1#', { key: 'externalSpill' }],
  ['=[Книга.xlsx]Лист!A1', { key: 'unknownLink', book: 'Книга.xlsx' }, { externalBooks: [] }],
  ['=[1]Нет!A1', { key: 'unknownLinkSheet', book: 'Книга.xlsx', sheet: 'Нет' }, { externalBooks: [{ target: 'Книга.xlsx', sheets: [{ name: 'Лист' }] }] }],
  ['=ставка', { key: 'unknownName', name: 'ставка' }, { names: [] }],
  ["='Лист!A1", { key: 'unterminatedSheet' }],
  ['=СУММ(1.5)', { key: 'unexpectedText', text: '.' }],
  ['={1;x}', { key: 'unexpectedInArray', text: 'x' }],
  ['=СУММ(A1,5)', { key: 'notWritable', text: 'A1' }],
]

/** Stored formulas that do not show in ru-RU, and their messages. */
const SHOWN: [string, EngineMessage][] = [
  ['SUM(1;2)', { key: 'unexpectedText', text: ';' }],
  ['"abc', { key: 'unterminatedString' }],
  ["1ИСТИНА''", { key: 'notWritable', text: 'ИСТИНА', tag: 'ru-RU' }],
]

describe('a formula that does not read', () => {
  it('validateLocalFormula and toInvariant give the message in Russian in ru-RU, in English without messages of the locale', () => {
    for (const [text, m, ctx] of TYPED) {
      reached.add(m.key)
      const [r, e] = [validateLocalFormula(text, ruRU, ctx), validateLocalFormula(text, english, ctx)]
      expect(r.ok ? 'ok' : r.message, text).toBe(ru(m))
      expect(e.ok ? 'ok' : e.message, text).toBe(en(m))
      expect(!r.ok && !e.ok && r.at === e.at, text).toBe(true)
      const t = toInvariant(text, ruRU)
      if (!t.ok) expect([t.message, t.at], text).toEqual([ru(m), (r as { at: number }).at])
    }
  })

  it('toLocal gives the message in the locale shown in', () => {
    for (const [formula, m] of SHOWN) {
      reached.add(m.key)
      const [r, e] = [toLocal(formula, ruRU), toLocal(formula, english)]
      expect(r.ok ? 'ok' : r.message, formula).toBe(ru(m))
      expect(e.ok ? 'ok' : e.message, formula).toBe(en(m).replace('ru-RU', 'xx-YY'))
    }
  })

  it("parseEntry stops a formula with the message in the book's locale", () => {
    const stops = TYPED.filter(([text, , ctx]) => !ctx && text !== '=' && ['syntax', 'empty', 'too_long', 'too_deep'].includes((validateLocalFormula(text, ruRU) as { code: string }).code))
    expect(stops.length).toBeGreaterThan(30)
    // a formula Excel closes on Return is taken (\`docs/entry.md\`, "A formula left open")
    for (const [text, m] of stops) expect(parseEntry(text, { locale: 'ru-RU' }), text).toMatchObject(closeTyped(text)?.silent ? { ok: true } : { ok: false, message: ru(m) })
    expect(parseEntry('=1+')).toEqual({ ok: false, code: 'syntax', message: 'unexpected end of formula', at: 3 })
  })
})

const sheet = (extra: Partial<SheetInput> = {}): SheetInput => ({ name: 'Лист', cells: { A1: { value: 1 }, A2: { value: 2 } }, ...extra })
const table = (name: string, range: string, cols: [string, string]) => ({ name, range, columns: cols.map((c) => ({ name: c })) })
/** Two tables side by side, on rows 1 to 3 and 2 to 5, their headers in their first rows. */
const tables = (): SheetInput => ({
  name: 'Лист',
  cells: { A1: { value: 'А' }, B1: { value: 'Б' }, D2: { value: 'В' }, E2: { value: 'Г' } },
  tables: [table('Т1', 'A1:B3', ['А', 'Б']), table('Т2', 'D2:E5', ['В', 'Г'])],
})
const filtered = (): SheetInput => {
  const cells: SheetInput['cells'] = {}
  for (let r = 1; r <= 6; r++) Object.assign(cells, { [`A${r}`]: { value: r }, [`B${r}`]: { formula: `A${r}*10` } })
  return { name: 'Лист', cells, rows: { '2': { hidden: true }, '4': { hidden: true } }, filtered: true }
}

/** A book's sheets and names, an operation Excel refuses on it, and the message of the refusal. */
const REFUSED: [string, () => Pick<BookInput, 'sheets' | 'names'>, StructureOp, EngineMessage][] = [
  ['an empty sheet name', () => ({ sheets: [sheet()] }), { op: 'renameSheet', sheet: 'Лист', name: '' }, { key: 'sheetNameEmpty' }],
  ['a long sheet name', () => ({ sheets: [sheet()] }), { op: 'renameSheet', sheet: 'Лист', name: 'я'.repeat(32) }, { key: 'sheetNameLong', name: 'я'.repeat(32), length: 32 }],
  ['a colon in a sheet name', () => ({ sheets: [sheet()] }), { op: 'insertSheet', name: 'а:б' }, { key: 'sheetNameChar', char: ':' }],
  ['an apostrophe', () => ({ sheets: [sheet()] }), { op: 'renameSheet', sheet: 'Лист', name: "'Итоги" }, { key: 'sheetNameApostrophe' }],
  ['History', () => ({ sheets: [sheet()] }), { op: 'renameSheet', sheet: 'Лист', name: 'history' }, { key: 'sheetNameHistory' }],
  ['a sheet name taken', () => ({ sheets: [sheet(), { name: 'Итоги', cells: {} }] }), { op: 'insertSheet', name: 'итоги' }, { key: 'sheetNameTaken', name: 'итоги' }],
  ['the last sheet', () => ({ sheets: [sheet()] }), { op: 'deleteSheet', sheet: 'Лист' }, { key: 'lastSheet' }],
  ['the last visible sheet deleted', () => ({ sheets: [sheet(), { name: 'Скрытый', cells: {}, hidden: true }] }), { op: 'deleteSheet', sheet: 'Лист' }, { key: 'lastVisibleSheet', sheet: 'Лист' }],
  ['the last visible sheet hidden', () => ({ sheets: [sheet(), { name: 'Скрытый', cells: {}, hidden: 'very' }] }), { op: 'hideSheet', sheet: 'Лист' }, { key: 'lastVisibleSheet', sheet: 'Лист' }],
  ['a very hidden sheet deleted', () => ({ sheets: [sheet(), { name: 'Ключи', cells: {}, hidden: 'very' }] }), { op: 'deleteSheet', sheet: 'Ключи' }, { key: 'veryHidden', sheet: 'Ключи' }],
  ['a very hidden sheet moved', () => ({ sheets: [sheet(), { name: 'Ключи', cells: {}, hidden: 'very' }] }), { op: 'moveSheet', sheet: 'Ключи', before: 'Лист' }, { key: 'veryHidden', sheet: 'Ключи' }],
  ['Group past level 7', () => ({ sheets: [sheet({ rows: { '3:5': { outlineLevel: 7 } } })] }), { op: 'group', sheet: 'Лист', rows: '3:5' }, { key: 'outlineDeep', lines: 'rows', span: '3:5' }],
  ['Ungroup of columns in no group', () => ({ sheets: [sheet()] }), { op: 'ungroup', sheet: 'Лист', columns: 'B:D' }, { key: 'notGrouped', lines: 'columns', span: 'B:D' }],
  ['Hide Detail at a row in no group', () => ({ sheets: [sheet()] }), { op: 'hideDetail', sheet: 'Лист', row: 4 }, { key: 'noSummary', lines: 'rows', span: '4' }],
  ['Show Detail of a group shown', () => ({ sheets: [sheet({ rows: { '2:3': { outlineLevel: 1 } } })] }), { op: 'showDetail', sheet: 'Лист', row: 4 }, { key: 'detailShown', lines: 'rows', span: '4' }],
  ['Hide Detail of a group hidden', () => ({ sheets: [sheet({ columns: { 'B:C': { outlineLevel: 1, hidden: true } } })] }), { op: 'hideDetail', sheet: 'Лист', column: 'D' }, { key: 'detailHidden', lines: 'columns', span: 'D' }],
  ['a cell off the sheet', () => ({ sheets: [sheet({ cells: { A1048576: { value: 1 } } })] }), { op: 'insertRows', sheet: 'Лист', row: 1, count: 2 }, { key: 'insertOffSheet', count: 2, lines: 'rows', sheet: 'Лист', cell: 'A1048576' }],
  ['a merged area off the sheet', () => ({ sheets: [sheet({ merges: ['XFC1:XFD1'] })] }), { op: 'insertColumns', sheet: 'Лист', column: 'B' }, { key: 'insertMergeOffSheet', count: 1, lines: 'columns', sheet: 'Лист' }],
  ['part of an array', () => ({ sheets: [sheet({ cells: { A1: { value: 1 }, C1: { formula: 'A1:A3*2', array: 'C1:C3' } } })] }), { op: 'deleteRows', sheet: 'Лист', row: 2 }, { key: 'arrayPart', sheet: 'Лист', cell: 'C1', area: 'C1:C3' }],
  ['a move off the sheet', () => ({ sheets: [sheet()] }), { op: 'moveRange', from: { sheet: 'Лист', range: 'A1:A2' }, to: { sheet: 'Лист', cell: 'A1048576' } }, { key: 'moveOffSheet', sheet: 'Лист', cell: 'A1048576' }],
  ['a copy off the sheet', () => ({ sheets: [sheet()] }), { op: 'copyRange', from: { sheet: 'Лист', range: 'A1:B1' }, to: { sheet: 'Лист', range: 'XFD1' } }, { key: 'copyOffSheet', sheet: 'Лист', cell: 'XFD1' }],
  ['a move cutting a merged area', () => ({ sheets: [sheet({ merges: ['A1:B1'] })] }), { op: 'moveRange', from: { sheet: 'Лист', range: 'A1' }, to: { sheet: 'Лист', cell: 'D1' } }, { key: 'moveCutsMerge', sheet: 'Лист', area: 'A1:B1' }],
  ['a copy onto part of a merged area', () => ({ sheets: [sheet({ merges: ['D1:E1'] })] }), { op: 'copyRange', from: { sheet: 'Лист', range: 'A1' }, to: { sheet: 'Лист', range: 'D1' } }, { key: 'pasteOnMerge', sheet: 'Лист', area: 'D1:E1' }],
  ['a copy without hidden rows over hidden rows', () => ({ sheets: [filtered()] }), { op: 'copyRange', from: { sheet: 'Лист', range: 'A1:B6' }, to: { sheet: 'Лист', range: 'D1:E6' } }, { key: 'filteredPaste', sheet: 'Лист', area: 'D1:E6' }],
  ['a table onto a table', () => ({ sheets: [tables()] }), { op: 'moveRange', from: { sheet: 'Лист', range: 'A1:B3' }, to: { sheet: 'Лист', cell: 'D1' } }, { key: 'tableLands', table: 'Т1' }],
  ['rows into two tables', () => ({ sheets: [tables()] }), { op: 'insertRows', sheet: 'Лист', row: 3 }, { key: 'tableInsertTwo', lines: 'rows', tables: ['Т1', 'Т2'] }],
  ['rows out of two tables', () => ({ sheets: [tables()] }), { op: 'deleteRows', sheet: 'Лист', row: 3 }, { key: 'tableDeleteTwo', lines: 'rows', tables: ['Т1', 'Т2'] }],
  ['a header row alone', () => ({ sheets: [tables()] }), { op: 'deleteRows', sheet: 'Лист', row: 1 }, { key: 'tableHeader', table: 'Т1' }],
  ["a table's name by the rules", () => ({ sheets: [tables()] }), { op: 'renameTable', table: 'Т1', name: 'A1' }, { key: 'tableName', name: 'A1', rule: 'cell' }],
  ["a table's name taken", () => ({ sheets: [tables()] }), { op: 'renameTable', table: 'Т2', name: 'т1' }, { key: 'tableNameTaken', table: 'Т1' }],
  ['a defined name', () => ({ sheets: [tables()], names: [{ name: 'ставка', ref: '1' }] }), { op: 'renameTable', table: 'Т1', name: 'Ставка' }, { key: 'tableNameDefined', name: 'Ставка' }],
  ['a table over a table', () => ({ sheets: [tables()] }), { op: 'addTable', sheet: 'Лист', table: table('Т3', 'B2:C4', ['В', 'Г']) }, { key: 'tableOverTable', range: 'B2:C4', table: 'Т1' }],
  ['a table over a merged area', () => ({ sheets: [{ ...tables(), merges: ['G1:H1'] }] }), { op: 'addTable', sheet: 'Лист', table: table('Т3', 'G1:H3', ['В', 'Г']) }, { key: 'tableOverMerge', range: 'G1:H3', area: 'G1:H1' }],
  ['cells of a table merged', () => ({ sheets: [tables()] }), { op: 'mergeCells', sheet: 'Лист', range: 'B2:C2' }, { key: 'mergeInTable', range: 'B2:C2', table: 'Т1' }],
  ['an array formula onto a table', () => ({ sheets: [{ ...tables(), cells: { ...tables().cells, G1: { formula: '{1,2}', array: 'G1:H1' } } }] }), { op: 'copyRange', from: { sheet: 'Лист', range: 'G1:H1' }, to: { sheet: 'Лист', range: 'D3' } }, { key: 'arrayInTable', area: 'D3:E3', table: 'Т2' }],
]

const message = (r: { ok: boolean; message?: string }) => (r.ok ? 'done' : r.message)

describe('a refused structure operation', () => {
  it("editStructure gives the message in the book's locale", () => {
    for (const [what, book, op, m] of REFUSED) {
      reached.add(m.key)
      expect(message(editStructure({ locale: 'ru-RU', ...book() }, op)), what).toBe(ru(m))
      expect(message(editStructure({ ...book() }, op)), what).toBe(en(m))
      expect(message(editStructure({ locale: 'en-US', ...book() }, op)), what).toBe(en(m))
    }
    // ИСТИНА is the boolean of a ru-RU book, a name in en-US
    const named = (locale: string) => message(editStructure({ locale, sheets: [tables()] }, { op: 'renameTable', table: 'Т1', name: 'истина' }))
    expect([named('ru-RU'), named('en-US')]).toEqual([ru({ key: 'tableName', name: 'истина', rule: 'boolean' }), 'done'])
  })

  describe.each(['closures', 'eval'] as const)('in sessions, execution %s', (execution) => {
    const { keep } = closedAfter()

    it('a synchronous session', () => {
      for (const [what, book, op, m] of REFUSED) {
        const s = createSession({ locale: 'ru-RU', ...book() }, { execution })
        expect(message(s.editStructure(op)), what).toBe(ru(m))
        s.close()
      }
    })

    it.each([
      ['in the same thread', () => createInlineCalc()],
      ['in threads', () => pool({ threads: 1 })],
    ] as const)('a session %s', async (_, make) => {
      const calc = keep(await make())
      for (const [what, book, op, m] of REFUSED) {
        const s = await calc.createSession({ locale: 'ru-RU', ...book() }, { execution })
        expect(message(await s.editStructure(op)), what).toBe(ru(m))
        await s.close()
      }
    }, 60_000)
  })
})

/** A ru-RU book or another where text typed into a cell is refused: a table's header cell, part of an old-kind array, and a rule of data validation. */
const typedBook = (locale?: string): BookInput => ({
  ...(locale ? { locale } : {}),
  sheets: [
    {
      name: 'Лист',
      cells: { A1: { value: 'Кол' }, A2: { value: 1 }, C1: { formula: 'A2:A3*2', array: 'C1:C2' } },
      tables: [{ name: 'Т', range: 'A1:A2', columns: [{ name: 'Кол' }] }],
      validations: [{ range: 'E1', type: 'whole', operator: 'greaterThan', formula1: '0' }],
    },
  ],
})

/** Text typed into a cell, refused with the message. */
const REFUSED_TEXT: [string, string, EngineMessage][] = [
  ['A1', '=1+1', { key: 'headerFormula', sheet: 'Лист', cell: 'A1', table: 'Т' }],
  ['C2', '5', { key: 'arrayPart', sheet: 'Лист', cell: 'C1', area: 'C1:C2' }],
]

describe('text typed into a cell of a session', () => {
  afterAll(closeCalcs)
  it.each(KINDS)("refused, and a value the cell's rule does not take, in the book's locale: %s", async (kind) => {
    const open = await opener(kind)
    for (const [locale, say] of [['ru-RU', ru], [undefined, en]] as const) {
      const s = await open(typedBook(locale), {})
      for (const [cell, text, m] of REFUSED_TEXT) {
        reached.add(m.key)
        expect(await s.setCellText('Лист', cell, text), `${cell} ${text}`).toEqual({ ok: false, code: m.key === 'headerFormula' ? 'header' : 'array', message: say(m) })
      }
      reached.add('validationError')
      const r = await s.setCellText('Лист', 'E1', '-1')
      expect(r.ok && r.validation?.message).toBe(say({ key: 'validationError' }))
      expect(await s.validateEntry('Лист', 'E1', { value: 0 })).toMatchObject({ ok: false, message: say({ key: 'validationError' }) })
      await s.close()
    }
  }, 60_000)
})

/** A book whose explanation holds every piece: a value, a formula, a range, a spilled cell, an error and the cell it came from, a boolean, an empty cell. */
const explained = (locale?: string): BookInput => ({
  ...(locale ? { locale } : {}),
  sheets: [{ name: 'План', cells: { A1: { value: 100 }, A2: { value: 250.5 }, B1: { formula: 'ROUND(SUM(A1:A2)*1.2,0)+C2' }, C1: { formula: 'SEQUENCE(2)' }, D1: { formula: 'IF(E1,1.5,NA())' }, F1: { formula: 'D1*2' } } }],
})
const RU_TEXT = [
  "План!B1 = 423  ← ОКРУГЛ(СУММ(A1:A2)*1,2;0)+C2",
  "  План!A1:A2 (2 ячейки)",
  "  План!C2 = 2  ← перенос из План!C1",
  "    План!C1 = 1  ← ПОСЛЕД(2)",
].join('\n')
const EN_TEXT = [
  "План!B1 = 423  ← ROUND(SUM(A1:A2)*1.2,0)+C2",
  "  План!A1:A2 (2 cells)",
  "  План!C2 = 2  ← spilled from План!C1",
  "    План!C1 = 1  ← SEQUENCE(2)",
].join('\n')

describe('explain', () => {
  it("writes the text in the book's locale: words, numbers, formulas, errors and booleans; the tree as stored", () => {
    for (const key of ['explainCells', 'explainSpill', 'explainValue', 'explainFormula', 'explainErrorFrom', 'explainErrorHere'] as const) reached.add(key)
    const [r, e] = [createSession(explained('ru-RU')), createSession(explained())]
    expect(r.explain('План', 'B1').text).toBe(RU_TEXT)
    expect(e.explain('План', 'B1').text).toBe(EN_TEXT)
    expect(r.explain('План', 'F1').text).toBe("План!F1 = #Н/Д  ← D1*2  (ошибка из План!D1)\n  План!D1 = #Н/Д  ← ЕСЛИ(E1;1,5;НД())  (ошибка возникает здесь)\n    План!E1 = пусто")
    expect(e.explain('План', 'F1').text).toBe("План!F1 = #N/A  ← D1*2  (error from План!D1)\n  План!D1 = #N/A  ← IF(E1,1.5,NA())  (the error begins here)\n    План!E1 = empty")
    r.setCells([{ sheet: 'План', cell: 'E1', input: { value: true } }])
    expect(r.explain('План', 'D1').text).toBe("План!D1 = 1,5  ← ЕСЛИ(E1;1,5;НД())\n  План!E1 = ИСТИНА")
    const { text: _, ...tree } = r.explain('План', 'B1')
    expect(tree).toEqual((({ text: _t, ...rest }) => rest)(e.explain('План', 'B1')))
  })

  it('a formula a circular reference stops holds its value, in either locale', () => {
    for (const key of ['explainHeldFrom', 'explainHeldHere'] as const) reached.add(key)
    const cells = { A1: { formula: 'B1+1', held: 2.5 }, B1: { formula: 'A1' }, C1: { formula: 'A1*2' } }
    const [r, e] = [createSession({ sheets: [{ name: 'План', cells }], locale: 'ru-RU' }), createSession({ sheets: [{ name: 'План', cells }] })]
    expect(r.explain('План', 'C1').text).toBe('План!C1 = 0  ← A1*2  (значение удержано: циклическая ссылка через План!A1)\n  План!A1 = 2,5  ← B1+1  (значение удержано: циклическая ссылка)\n    План!B1 = 0  ← A1  (значение удержано: циклическая ссылка)\n      План!A1 = 2,5  ← B1+1  (значение удержано: циклическая ссылка)')
    expect(e.explain('План', 'C1').text).toBe('План!C1 = 0  ← A1*2  (held by a circular reference through План!A1)\n  План!A1 = 2.5  ← B1+1  (held by a circular reference)\n    План!B1 = 0  ← A1  (held by a circular reference)\n      План!A1 = 2.5  ← B1+1  (held by a circular reference)')
  })

  it('the same text in the same thread and in threads', async () => {
    for (const calc of [createInlineCalc(), await pool({ threads: 1 })]) {
      const s = await calc.createSession(explained('ru-RU'))
      expect((await s.explain('План', 'B1')).text).toBe(RU_TEXT)
      await calc.close()
    }
  }, 60_000)
})

it("an edit over limits.writes says it in the book's language, in a session as in a calculator", async () => {
  reached.add('writesLimit')
  const m: EngineMessage = { key: 'writesLimit', size: 1_048_576, max: 100_000 }
  const copy = { op: 'copyRange', from: { sheet: 'Лист', range: 'A1' }, to: { sheet: 'Лист', range: 'B:B' } } as const
  for (const [locale, text] of [['ru-RU', ru(m)], [undefined, en(m)]] as const) {
    const book: BookInput = { ...(locale ? { locale } : {}), sheets: [{ name: 'Лист', cells: { A1: { value: 1 } } }] }
    expect(() => createSession(book, { limits: { writes: 100_000 } }).editStructure(copy)).toThrow(text)
    const calc = createInlineCalc({ limits: { writes: 100_000 } })
    await expect((await calc.createSession(book)).editStructure(copy)).rejects.toThrow(text)
    await calc.close()
  }
})

it('the cases reach every message of the catalog but the rules of names', () => {
  expect([...Object.keys(EN_MESSAGES), ...Object.keys(EDIT_MESSAGES)].filter((k) => !reached.has(k as MessageKey))).toEqual(['nameRule'])
})
