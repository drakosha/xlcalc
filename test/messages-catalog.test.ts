/**
 * The catalog of messages: every message of en-US has its ru-RU text, which takes the same parameters and shows each
 * one the English shows, in Russian words and the locale's notation, written as an interface writes: no long dash, no
 * straight double quotes of its own (`«»`), no full stop at the end (but in Excel's own words), no lower-case start.
 */
import { describe, expect, it } from 'vitest'
import { EN_MESSAGES } from '../src/locales/en-US.js'
import { EDIT_MESSAGES } from '../src/locale/edit-messages.js'
import { toLocal } from '../src/locale/index.js'
import { ruRU } from './locales.js'
import { KEPT_MESSAGE_KEYS, ownText, say, XLSX_MESSAGE_KEYS, type Message, type MessageKey, type MessageParams } from '../src/locale/messages.js'
import type { XlsxMessage } from '../src/xlsx/issues.js'
import { issueText, XLSX_EN_MESSAGES } from '../src/xlsx/messages.js'
import { KEPT_SAMPLES, XLSX_SAMPLES } from './fixtures/xlsx-messages.js'
import { KEPT_EN_MESSAGES, keptText, type KeptMessage } from '../src/xlsx/keep/messages.js'

/** Parameters of each message: a key without samples does not compile. */
const SAMPLES: { [K in MessageKey]: MessageParams[K][] } = {
  rowOutOfRange: [{ row: '0' }],
  columnOutOfRange: [{ column: 'XFE' }],
  numberOutOfRange: [{ number: '1E400' }],
  spillAfterRange: [{}],
  spillAcrossSheets: [{}],
  unterminatedString: [{}],
  unknownErrorLiteral: [{}],
  externalReference: [{}],
  externalAcrossSheets: [{}],
  externalSpill: [{}],
  referenceAfterSheet: [{}],
  quotedSheetAlone: [{}],
  unexpectedCharacter: [{ char: '~' }],
  emptyFormula: [{}],
  tooLong: [{ max: 8192 }],
  tooDeep: [{ max: 64 }, { max: 21 }],
  unexpectedEnd: [{}],
  unexpected: [{ found: { t: 'fn', v: 'ROUND' } }, { found: { t: 'ident', v: 'TRUE' } }, { found: { t: 'op', v: '+' } }, { found: { t: 'num' } }, { found: { t: ',' } }, { found: { t: ')' } }],
  expected: [{ char: ')', found: { t: 'eof' } }, { char: ')', found: { t: 'str' } }],
  commaOutside: [{}],
  arraySeparator: [{ found: { t: 'num' } }, { found: { t: 'eof' } }],
  arrayRows: [{}],
  arrayValue: [{ found: { t: 'op', v: '+' } }, { found: { t: ',' } }],
  letTwice: [{ name: 'x' }],
  letName: [{ position: 3 }],
  letEmpty: [{}],
  letTooMany: [{ max: 254 }],
  specUnterminated: [{}],
  specBracketInColumn: [{}],
  specUnknown: [{ word: '#Data' }, { word: '#Foo' }],
  specSpan: [{}],
  specOpen: [{}],
  specTwoColumns: [{}],
  specTwice: [{ word: '#Headers' }],
  specClose: [{}],
  specTogether: [{ words: ['#All', '#Data'] }],
  externalLiterals: [{ fn: 'BALANCE' }],
  argumentCount: [
    { fn: 'ROUND', counts: { min: 2, max: 2 }, got: 3 },
    { fn: 'IF', counts: { min: 1, max: 3 }, got: 4 },
    { fn: 'SUMIFS', counts: { min: 3, max: 255, parity: 'odd' }, got: 4 },
    { fn: 'SORTBY', counts: { min: 2, max: 253, parity: 'odd', or: true }, got: 4 },
    { fn: 'PI', counts: { min: 0, max: 0 }, got: 1 },
  ],
  argumentLeftOut: [{ fn: 'SUMIF', position: 3 }],
  tableRefOutside: [{}],
  unknownTable: [{ table: 'Sales' }],
  unknownColumn: [{ table: 'Sales', column: 'Qty' }],
  unknownFunction: [{ fn: 'FOO' }, { fn: 'STOCKHISTORY' }],
  unknownSheet: [{ sheet: 'Data' }],
  unknownLink: [{ book: 'Book2.xlsx' }],
  unknownLinkSheet: [{ book: 'Book2.xlsx', sheet: 'Data' }],
  unknownName: [{ name: 'rate' }],
  unterminatedSheet: [{}],
  unexpectedText: [{ text: '.' }],
  unexpectedInArray: [{ text: 'x' }],
  notWritable: [{ text: 'A1' }, { text: 'x', tag: 'ru-RU' }],
  nameRule: [{ name: 'q', rule: 'chars' }, { name: 'A1', rule: 'cell' }, { name: 'TRUE', rule: 'true' }, { name: 'ЛОЖЬ', rule: 'boolean' }],
  sheetNameEmpty: [{}],
  sheetNameLong: [{ name: 'x'.repeat(32), length: 32 }, { name: 'y'.repeat(41), length: 41 }],
  sheetNameChar: [{ char: ':' }],
  sheetNameApostrophe: [{}],
  sheetNameHistory: [{}],
  sheetNameTaken: [{ name: 'Plan' }],
  lastSheet: [{}],
  lastVisibleSheet: [{ sheet: 'Feb' }],
  veryHidden: [{ sheet: 'Keys' }],
  outlineDeep: [{ lines: 'rows', span: '3:5' }],
  notGrouped: [{ lines: 'columns', span: 'A' }],
  noSummary: [{ lines: 'rows', span: '12' }],
  detailShown: [{ lines: 'rows', span: '8' }],
  detailHidden: [{ lines: 'columns', span: 'K' }],
  insertOffSheet: [{ count: 1, lines: 'rows', sheet: 'S', cell: 'A1048576' }, { count: 3, lines: 'columns', sheet: 'S', cell: 'XFD1' }, { count: 5, lines: 'rows', sheet: 'S', cell: 'A1048576' }],
  insertMergeOffSheet: [{ count: 2, lines: 'rows', sheet: 'S' }],
  arrayPart: [{ sheet: 'S', cell: 'C1', area: 'C1:C3' }],
  moveOffSheet: [{ sheet: 'S', cell: 'A1048575' }],
  copyOffSheet: [{ sheet: 'S', cell: 'XFD1' }],
  writesLimit: [{ size: 1048576, max: 100000 }],
  moveCutsMerge: [{ sheet: 'S', area: 'A1:B1' }],
  pasteOnMerge: [{ sheet: 'S', area: 'D1:E1' }],
  contentOnMerge: [{ sheet: 'S', area: 'F1:G1' }],
  filteredPaste: [{ sheet: 'S', area: 'D1:E6' }],
  tableLands: [{ table: 'Sales' }],
  tableInsertTwo: [{ lines: 'rows', tables: ['Sales', 'Costs'] }],
  tableDeleteTwo: [{ lines: 'columns', tables: ['Sales', 'Costs', 'Taxes'] }],
  tableHeader: [{ table: 'Sales' }],
  tableName: [{ name: 'A1', rule: 'cell' }, { name: 'ИСТИНА', rule: 'boolean' }],
  tableNameTaken: [{ table: 'Sales' }],
  tableNameDefined: [{ name: 'rate' }],
  tableOverTable: [{ range: 'A1:C9', table: 'Sales' }],
  tableOverMerge: [{ range: 'A1:C9', area: 'B2:C3' }],
  tableOverFilter: [{ range: 'A1:C9', filter: 'B1:D20' }],
  mergeInTable: [{ range: 'B2:D3', table: 'Sales' }],
  arrayInTable: [{ area: 'D2:E2', table: 'Sales' }],
  headerFormula: [{ sheet: 'Лист', cell: 'B2', table: 'Т' }],
  validationError: [{}],
  explainCells: [{ count: 1 }, { count: 2 }, { count: 11 }, { count: 25 }],
  explainSpill: [{ ref: "'S'!B2" }],
  explainValue: [{ value: 1.5 }, { value: true }, { value: null }, { value: { error: 'NA' } }],
  explainFormula: [{ formula: 'ROUND(B1*1.2,0)' }],
  explainErrorFrom: [{ ref: "'S'!A2" }],
  explainErrorHere: [{}],
  explainHeldFrom: [{ ref: "'S'!A2" }],
  explainHeldHere: [{}],
  ...XLSX_SAMPLES,
  ...KEPT_SAMPLES,
}

type Any = Message<MessageKey>
const messages = (): Any[] => Object.entries(SAMPLES).flatMap(([key, ps]) => (ps as object[]).map((p) => ({ key, ...p }) as Any))
const isXlsx = (m: Any): m is XlsxMessage => Object.hasOwn(XLSX_MESSAGE_KEYS, m.key)
const isKept = (m: Any): m is KeptMessage => Object.hasOwn(KEPT_MESSAGE_KEYS, m.key)
/** A message in English: en-US's of the engine, the xlsx subpath's of an issue, the xlsx/keep subpath's of a refusal. */
const en = (m: Any) => (isXlsx(m) ? issueText('en-US', m) : isKept(m) ? keptText('en-US', m) : say('en-US', m))
const ru = (m: Any) => ownText('ru-RU', m, (f) => (toLocal(f, ruRU) as { formula: string }).formula)!

/** The ways ru-RU writes a parameter's value: as it is, a function's or a specifier's name, a number, a formula. */
function forms(v: string | number): string[] {
  const s = String(v)
  const local = toLocal(s, ruRU)
  return [s, ruRU.functions[s], ruRU.tableItems?.[s as '#Data'], s.replace('.', ','), local.ok ? local.formula : undefined].filter((x): x is string => x !== undefined)
}

/** The parameters whose value picks the message's words rather than standing in it: rows or columns, a rule of names, a part of a style, a rule's text. */
const WORDS = new Set(['key', 'lines', 'rule', 'part', 'field'])

/** The parameters' values that a message's text may show: strings, numbers, the strings of lists and those of a refusal's part of the file. */
const values = (m: Any): (string | number)[] =>
  Object.entries(m).flatMap(([k, v]) => (WORDS.has(k) ? [] : k === 'what' ? values(v as Any) : Array.isArray(v) ? v : typeof v === 'string' || typeof v === 'number' ? [v] : []))

/** The pieces of the explanation's lines, which stand inside a line rather than alone. */
/** Excel's own words, as its window shows them, a full stop kept: the alert of a rule without a message of its own (`ui/fill-held`). */
const EXCEL_WORDS = new Set<MessageKey>(['validationError'])

const PIECES = new Set<MessageKey>(['explainCells', 'explainSpill', 'explainValue', 'explainFormula', 'explainErrorFrom', 'explainErrorHere', 'explainHeldFrom', 'explainHeldHere'])

describe('the catalog of messages', () => {
  it("ru-RU writes every message of en-US and of the xlsx subpaths, whose keys registerLocale takes", () => {
    const keys = [...Object.keys(EN_MESSAGES), ...Object.keys(EDIT_MESSAGES), ...Object.keys(XLSX_EN_MESSAGES), ...Object.keys(KEPT_EN_MESSAGES)]
    expect(keys.filter((k) => ruRU.messages?.[k as MessageKey] === undefined)).toEqual([])
    expect(Object.keys(SAMPLES).sort()).toEqual([...keys].sort())
    expect(Object.keys(XLSX_MESSAGE_KEYS).sort()).toEqual(Object.keys(XLSX_EN_MESSAGES).sort())
    expect(Object.keys(KEPT_MESSAGE_KEYS).sort()).toEqual(Object.keys(KEPT_EN_MESSAGES).sort())
  })

  it('the Russian text shows every parameter the English one shows, in the locale', () => {
    const missing: string[] = []
    for (const m of messages()) {
      const [english, text] = [en(m), ru(m)]
      for (const v of values(m)) if (english.includes(String(v)) && !forms(v).some((f) => text.includes(f))) missing.push(`${m.key}: ${v} in ${text}`)
    }
    expect(missing).toEqual([])
  })

  it('the Russian text is Russian, written as an interface writes', () => {
    const bad: string[] = []
    for (const m of messages()) {
      const text = ru(m)
      if (text === en(m)) bad.push(`${m.key}: the English text`)
      if (/[—–]|\p{Extended_Pictographic}/u.test(text)) bad.push(`${m.key}: a dash or a pictograph in ${text}`)
      if (PIECES.has(m.key)) continue
      if (!/\p{Script=Cyrillic}/u.test(text)) bad.push(`${m.key}: no Russian word in ${text}`)
      if (text.includes('"')) bad.push(`${m.key}: straight quotes in ${text}`)
      if (/^\p{Ll}/u.test(text) || (/\.$/.test(text) && !EXCEL_WORDS.has(m.key))) bad.push(`${m.key}: a lower-case start or a full stop in ${text}`)
    }
    expect(bad).toEqual([])
  })

  it('numbers take the noun in its form: 1 ячейка, 2 ячейки, 11 ячеек, 25 ячеек', () => {
    expect(SAMPLES.explainCells.map((p) => ru({ key: 'explainCells', ...p }))).toEqual(['1 ячейка', '2 ячейки', '11 ячеек', '25 ячеек'])
    expect(SAMPLES.insertOffSheet.map((p) => ru({ key: 'insertOffSheet', ...p }).split(':')[0])).toEqual(['Нельзя вставить 1 строку', 'Нельзя вставить 3 столбца', 'Нельзя вставить 5 строк'])
    expect(SAMPLES.argumentCount.map((p) => ru({ key: 'argumentCount', ...p }))).toEqual([
      'Функция ОКРУГЛ принимает 2 аргумента, а указано 3',
      'Функция ЕСЛИ принимает от 1 до 3 аргументов, а указано 4',
      'Функция СУММЕСЛИМН принимает нечётное число аргументов от 3 до 255, а указано 4',
      'Функция СОРТПО принимает 2 аргумента или нечётное число от 3 до 253, а указано 4',
      'Функция ПИ принимает 0 аргументов, а указано 1',
    ])
    expect(ru({ key: 'tooDeep', max: 21 })).toBe('Вложенность формулы больше 21 уровня')
  })

  it('a locale without its own text of a message takes the loaded locale\'s of its tag, else the English one', () => {
    const own = { ...ruRU, tag: 'xx-YY', messages: undefined }
    expect(say(own, { key: 'lastSheet' })).toBe('a workbook must contain at least one sheet')
    expect(say('xx-YY', { key: 'lastSheet' })).toBe('a workbook must contain at least one sheet')
    expect(say({ ...ruRU, messages: undefined }, { key: 'lastSheet' })).toBe('В книге должен остаться хотя бы один лист')
    expect(say({ ...ruRU, tag: 'xx-YY' }, { key: 'lastSheet' })).toBe('В книге должен остаться хотя бы один лист')
  })
})
