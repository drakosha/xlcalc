/**
 * Messages for people, in the language of a locale: why a formula does not read, why Excel refuses a structure
 * operation, the words of `explain`, what an xlsx file read or written lost. Code makes a message as a key with its
 * parameters (`Message`); the locale's `messages` write it (`Messages`), and a locale without its own text of a message
 * takes the English one: en-US's (`EN_MESSAGES`) for the engine's, the subpath's (`src/xlsx/messages.ts`) for the
 * issues of xlsx. Text is made only where a message is shown: a refusal, a formula that does not read, an explanation,
 * an issue of xlsx.
 */
import type { CellValue } from '../contract.js'
import { errorIn } from '../syntax/errors.js'
import { english } from './english.js'
import { LOCALES, localeOf } from './registry.js'
import type { Locale, TableItem } from './types.js'

/** What the parser found where it expected something else: a token's kind, with its text for an operator, a name or a function. */
export type Found = { t: 'op' | 'ident' | 'fn'; v: string } | { t: 'num' | 'str' | 'err' | 'ref' | 'table' | 'eof' | '(' | ')' | ',' | ';' | '{' | '}' }

/** The argument counts a function takes: `min` to `max`, past `min` of one parity (SUMIFS 3, 5, …), `or`: `min` too (SORTBY 2, 3, 5, …). */
export interface Counts {
  min: number
  max: number
  parity?: 'odd' | 'even'
  or?: boolean
}

/** Why Excel does not take a name: its characters, a cell's address, TRUE or FALSE, the book's locale's boolean. */
export type NameRule = 'chars' | 'cell' | 'true' | 'boolean'

/** Rows or columns. */
export type Lines = 'rows' | 'columns'

/** What a style, or a conditional format's differential format, holds that the model's style does not. */
export type StylePart =
  | 'accountingUnderline'
  | 'underlineRemoved'
  | 'superscript'
  | 'subscript'
  | 'fontName'
  | 'fontSize'
  | 'fontVertAlign'
  | 'fontOutline'
  | 'fontShadow'
  | 'fontCondense'
  | 'fontExtend'
  | 'diagonalBorder'
  | 'gradientFill'
  | 'gradientBesidePattern'
  | 'patternFill'
  | 'fillRemoved'
  | 'fillWithoutColor'
  | 'textRotation'
  | 'shrinkToFit'
  | 'readingOrder'
  | 'justifyLastLine'
  | 'alignment'
  | 'protection'

/** A part of a style the model does not hold, or a border line of a kind it does not have (the file's name of it). */
export type StyleLoss = StylePart | { border: string }

/** A text of a rule of data validation that Excel limits. */
export type RuleText = 'promptTitle' | 'prompt' | 'errorTitle' | 'error'

/** The parameters of the messages of `readXlsx`'s and `writeXlsx`'s issues, by key. */
export interface XlsxMessageParams {
  // reading: the workbook
  localeUnknown: { tag: string }
  builtInName: object
  nameOfSheetLeftOut: object
  hiddenName: object
  manualCalculation: object
  autoNoTable: object
  precisionAsDisplayed: object
  iterationCount: { count: string }
  iterationChange: { change: string }
  workbookProtection: object
  externalLinks: object
  pivotTables: object
  macros: object
  connections: object
  defaultFont: { name?: string; size?: number }
  sheetKind: object
  sheetPartMissing: object
  sheetNameRepaired: { name: string; sheet: string }
  stringsOpen: object
  allSheetsHidden: object
  tableNameTwice: { table: string }
  // reading: a sheet
  sharedWithoutMaster: object
  sharedCopy: object
  outlineStyles: object
  mergeOverlap: object
  autoFilterCriteria: object
  autoFilterPlain: object
  advancedFilter: object
  tableFilter: object
  sheetProtection: object
  drawings: object
  sparklines: object
  printSettings: object
  defaultWidth: { width: number }
  errorTyped: object
  sheetDataOpen: object
  // reading: formulas
  userFunction: object
  tableRefWithoutTable: object
  formulaExternal: object
  letPrefix: object
  formulaUnparsed: object
  // reading: cells
  cellOutside: object
  widthBad: object
  heightBad: object
  dateInvalid: object
  cellOpen: object
  valueMetadata: object
  richText: object
  errorUnknown: object
  errorUnspelled: object
  dataTable: object
  styleLost: { part: StyleLoss }
  eastAsianFormat: object
  unknownFormat: object
  // reading: tables
  tablePartBad: object
  /** `reason`: the check's own words (`InputError`), which are English. */
  tableLeftOut: { table: string; reason: string }
  headerGiven: { name: string }
  tableStyleCustom: { table: string; name: string }
  // reading: conditional formats
  condRangeOutside: object
  condStyleLost: { part: StyleLoss }
  condOperator: { operator: string }
  condNoFormulas: object
  condNoFormula: object
  condPeriod: { period: string }
  condAverage: object
  condColorScale: object
  condDataBar: object
  condIconSet: { set: string }
  condIcons: object
  condType: { type: string }
  condPivot: object
  // reading: data validation
  ruleListCommas: object
  ruleType: { type: string }
  ruleNoFormula: object
  ruleNoSecond: object
  ruleIme: object
  ruleNoRange: object
  ruleShared: object
  // reading: notes, hyperlinks, panes
  splitPanes: object
  frozenScrolled: object
  noteCell: object
  noteThread: object
  noteFormatted: object
  noteShown: object
  linkAddress: object
  linkCells: object
  // writing
  nameNotTaken: object
  nameNoSheet: object
  nameTwice: object
  nameBadRef: object
  noResult: object
  formulaNoValue: object
  formulaEmpty: object
  formulaAsValue: object
  tableFormulaOut: { table: string }
  cycleZero: object
  cellMeta: object
  textCut: object
  noteCut: object
  tooltipCut: object
  ruleTextCut: { field: RuleText; max: number }
  ruleFormulaOut: object
  condFormulaOut: object
  filterModeOut: object
}

/** The parameters of each message, by key. */
export interface MessageParams extends XlsxMessageParams {
  // reading a formula: the tokens
  rowOutOfRange: { row: string }
  columnOutOfRange: { column: string }
  numberOutOfRange: { number: string }
  spillAfterRange: object
  spillAcrossSheets: object
  unterminatedString: object
  unknownErrorLiteral: object
  externalReference: object
  externalAcrossSheets: object
  externalSpill: object
  referenceAfterSheet: object
  quotedSheetAlone: object
  unexpectedCharacter: { char: string }
  // the parser
  emptyFormula: object
  tooLong: { max: number }
  tooDeep: { max: number }
  unexpectedEnd: object
  unexpected: { found: Found }
  expected: { char: '(' | ')'; found: Found }
  commaOutside: object
  arraySeparator: { found: Found }
  arrayRows: object
  arrayValue: { found: Found }
  letTwice: { name: string }
  letName: { position: number }
  letEmpty: object
  letTooMany: { max: number }
  // structured references
  specUnterminated: object
  specBracketInColumn: object
  specUnknown: { word: string }
  specSpan: object
  specOpen: object
  specTwoColumns: object
  specTwice: { word: string }
  specClose: object
  specTogether: { words: string[] }
  // the checks of a formula
  externalLiterals: { fn: string }
  argumentCount: { fn: string; counts: Counts; got: number }
  argumentLeftOut: { fn: string; position: number }
  tableRefOutside: object
  unknownTable: { table: string }
  unknownColumn: { table: string; column: string }
  unknownFunction: { fn: string }
  unknownSheet: { sheet: string }
  unknownLink: { book: string }
  unknownLinkSheet: { book: string; sheet: string }
  unknownName: { name: string }
  // translating formula text
  unterminatedSheet: object
  unexpectedText: { text: string }
  unexpectedInArray: { text: string }
  /** `tag`: the locale written into, none for the stored text. */
  notWritable: { text: string; tag?: string }
  // names
  nameRule: { name: string; rule: NameRule }
  // structure operations
  sheetNameEmpty: object
  sheetNameLong: { name: string; length: number }
  sheetNameChar: { char: string }
  sheetNameApostrophe: object
  sheetNameHistory: object
  sheetNameTaken: { name: string }
  lastSheet: object
  lastVisibleSheet: { sheet: string }
  veryHidden: { sheet: string }
  outlineDeep: { lines: Lines; span: string }
  notGrouped: { lines: Lines; span: string }
  noSummary: { lines: Lines; span: string }
  detailShown: { lines: Lines; span: string }
  detailHidden: { lines: Lines; span: string }
  insertOffSheet: { count: number; lines: Lines; sheet: string; cell: string }
  insertMergeOffSheet: { count: number; lines: Lines; sheet: string }
  arrayPart: { sheet: string; cell: string; area: string }
  moveOffSheet: { sheet: string; cell: string }
  copyOffSheet: { sheet: string; cell: string }
  writesLimit: { size: number; max: number }
  moveCutsMerge: { sheet: string; area: string }
  pasteOnMerge: { sheet: string; area: string }
  filteredPaste: { sheet: string; area: string }
  tableLands: { table: string }
  tableInsertTwo: { lines: Lines; tables: string[] }
  tableDeleteTwo: { lines: Lines; tables: string[] }
  tableHeader: { table: string }
  tableName: { name: string; rule: NameRule }
  tableNameTaken: { table: string }
  tableNameDefined: { name: string }
  tableOverTable: { range: string; table: string }
  tableOverMerge: { range: string; area: string }
  mergeInTable: { range: string; table: string }
  arrayInTable: { area: string; table: string }
  // typing into a cell
  headerFormula: { sheet: string; cell: string; table: string }
  validationError: object
  // explain
  explainCells: { count: number }
  explainSpill: { ref: string }
  explainValue: { value: CellValue | undefined }
  explainFormula: { formula: string }
  explainErrorFrom: { ref: string }
  explainErrorHere: object
  explainHeldFrom: { ref: string }
  explainHeldHere: object
}

export type MessageKey = keyof MessageParams

/** The keys of the engine's messages, whose English text is en-US's (`EN_MESSAGES`). */
export type EngineKey = Exclude<MessageKey, keyof XlsxMessageParams>

/**
 * The keys of the messages of xlsx issues, which a locale's `messages` holds beside the engine's (`registerLocale` takes
 * both); their English text is the subpath's, out of the bundles that read no files.
 */
export const XLSX_MESSAGE_KEYS: Readonly<Record<keyof XlsxMessageParams, 0>> = {
  localeUnknown: 0, builtInName: 0, nameOfSheetLeftOut: 0, hiddenName: 0, manualCalculation: 0, autoNoTable: 0, precisionAsDisplayed: 0,
  iterationCount: 0, iterationChange: 0, workbookProtection: 0, externalLinks: 0, pivotTables: 0, macros: 0, connections: 0, defaultFont: 0,
  sheetKind: 0, sheetPartMissing: 0, sheetNameRepaired: 0, stringsOpen: 0, allSheetsHidden: 0, tableNameTwice: 0, sharedWithoutMaster: 0, sharedCopy: 0, outlineStyles: 0,
  mergeOverlap: 0, autoFilterCriteria: 0, autoFilterPlain: 0, advancedFilter: 0, tableFilter: 0, sheetProtection: 0, drawings: 0, sparklines: 0,
  printSettings: 0, defaultWidth: 0, errorTyped: 0, sheetDataOpen: 0, userFunction: 0, tableRefWithoutTable: 0, formulaExternal: 0, letPrefix: 0, formulaUnparsed: 0,
  cellOutside: 0, widthBad: 0, heightBad: 0, dateInvalid: 0, cellOpen: 0, valueMetadata: 0, richText: 0, errorUnknown: 0, errorUnspelled: 0, dataTable: 0, styleLost: 0, eastAsianFormat: 0, unknownFormat: 0,
  tablePartBad: 0, tableLeftOut: 0, headerGiven: 0, tableStyleCustom: 0, condRangeOutside: 0, condStyleLost: 0, condOperator: 0, condNoFormulas: 0, condNoFormula: 0,
  condPeriod: 0, condAverage: 0, condColorScale: 0, condDataBar: 0, condIconSet: 0, condIcons: 0, condType: 0, condPivot: 0, ruleListCommas: 0,
  ruleType: 0, ruleNoFormula: 0, ruleNoSecond: 0, ruleIme: 0, ruleNoRange: 0, ruleShared: 0, splitPanes: 0, frozenScrolled: 0, noteCell: 0,
  noteThread: 0, noteFormatted: 0, noteShown: 0, linkAddress: 0, linkCells: 0, nameNotTaken: 0, nameNoSheet: 0, nameTwice: 0, nameBadRef: 0,
  noResult: 0, formulaNoValue: 0, formulaEmpty: 0, formulaAsValue: 0, tableFormulaOut: 0, cycleZero: 0, cellMeta: 0, textCut: 0, noteCut: 0,
  tooltipCut: 0, ruleTextCut: 0, ruleFormulaOut: 0, condFormulaOut: 0, filterModeOut: 0,
}

/** A message: its key and its parameters; the engine's by default (an issue of xlsx is `XlsxMessage`). */
export type Message<K extends MessageKey = EngineKey> = { [P in K]: { key: P } & MessageParams[P] }[K]

/** What a locale's messages read the locale by: its formula text and the words of formulas in it. */
export interface Words {
  /** The locale's formula text: its separators, booleans and error literals. */
  readonly loc: Locale
  /** A function's name as the locale writes it (`ROUND` is `ОКРУГЛ` in ru-RU). */
  fn(name: string): string
  /** A specifier of structured references as the locale writes it (`#Data` is `#Данные` in ru-RU). */
  item(word: string): string
  /** A number with the locale's decimal separator. */
  num(n: number): string
  /** A value as a formula of the locale writes it: a number, text in quotes (a quote doubled), a boolean, an error. */
  value(v: Exclude<CellValue, null>): string
  /** A stored formula as the locale shows it (`toLocal`), where the caller gives the way. */
  formula(text: string): string
}

/** A locale's messages: the text of each, in its language. */
export type Messages = { readonly [K in MessageKey]: (p: MessageParams[K], w: Words) => string }

/** en-US's messages of the engine, which read nothing of a locale: what every locale falls back to. */
export type EnglishMessages = { readonly [K in EngineKey]: (p: MessageParams[K]) => string }

function wordsOf(loc: Locale, formula: (text: string) => string): Words {
  // every digit of the double, as JSON has them, with the exponent a formula writes (`1E+21`, not `1e+21`)
  const num = (n: number) => String(n).replace('e', 'E').replace('.', loc.decimalSeparator)
  return {
    loc,
    fn: (name) => loc.functions[name] ?? name,
    item: (word) => loc.tableItems?.[word as TableItem] ?? word,
    num,
    value: (v) => {
      if (typeof v === 'number') return num(v)
      if (typeof v === 'string') return `"${v.replaceAll('"', '""')}"`
      if (typeof v === 'boolean') return v ? loc.booleans.TRUE : loc.booleans.FALSE
      return errorIn(v.error, loc.errors)
    },
    formula,
  }
}

/**
 * A locale's own text of a message: a loaded locale's tag (en-US without one) or a locale object, its own messages or
 * else those of the loaded locale of its tag; undefined when neither has one. `formula` shows a stored formula in the
 * locale, for the messages that hold one.
 */
export function ownText<K extends MessageKey>(where: string | Locale | undefined, m: { key: K } & MessageParams[K], formula: (text: string) => string = (f) => f): string | undefined {
  const loc = localeOf(where)
  const own = (loc.messages ?? LOCALES.get(loc.tag)?.messages)?.[m.key]
  return own?.(m, wordsOf(loc, formula))
}

/** A message of the engine in the language of a locale (`ownText`), else in English. */
export function say<K extends EngineKey>(where: string | Locale | undefined, m: { key: K } & MessageParams[K], formula?: (text: string) => string): string {
  return ownText(where, m, formula) ?? english(m)
}
