/**
 * Messages for people, in the language of a locale: why a formula does not read, why Excel refuses a structure
 * operation, the words of `explain`, what an xlsx file read or written lost. Code makes a message as a key with its parameters (`Message`); the locale's
 * `messages` write it (`Messages`), en-US's every one of them, and a locale without its own text of a message takes
 * en-US's. Text is made only where a message is shown: a refusal, a formula that does not read, an explanation, an
 * issue of xlsx.
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
  // reading: formulas
  userFunction: object
  tableRefWithoutTable: object
  formulaExternal: object
  letPrefix: object
  formulaUnparsed: object
  // reading: cells
  cellOutside: object
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
  spillAfterRange: object
  spillAcrossSheets: object
  unterminatedString: object
  unknownErrorLiteral: object
  externalReference: object
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
  insertOffSheet: { count: number; lines: Lines; sheet: string; cell: string }
  insertMergeOffSheet: { count: number; lines: Lines; sheet: string }
  arrayPart: { sheet: string; cell: string; area: string }
  moveOffSheet: { sheet: string; cell: string }
  copyOffSheet: { sheet: string; cell: string }
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
}

export type MessageKey = keyof MessageParams

/** A message: its key and its parameters. */
export type Message<K extends MessageKey = MessageKey> = { [P in K]: { key: P } & MessageParams[P] }[K]

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
  /** A value as a formula of the locale writes it: a number, text in quotes, a boolean, an error. */
  value(v: Exclude<CellValue, null>): string
  /** A stored formula as the locale shows it (`toLocal`), where the caller gives the way. */
  formula(text: string): string
}

/** A locale's messages: the text of each, in its language. */
export type Messages = { readonly [K in MessageKey]: (p: MessageParams[K], w: Words) => string }

/** en-US's messages, which read nothing of a locale: what every locale falls back to. */
export type EnglishMessages = { readonly [K in MessageKey]: (p: MessageParams[K]) => string }

function wordsOf(loc: Locale, formula: (text: string) => string): Words {
  const num = (n: number) => JSON.stringify(n).replace('.', loc.decimalSeparator)
  return {
    loc,
    fn: (name) => loc.functions[name] ?? name,
    item: (word) => loc.tableItems?.[word as TableItem] ?? word,
    num,
    value: (v) => {
      if (typeof v === 'number') return num(v)
      if (typeof v === 'string') return JSON.stringify(v)
      if (typeof v === 'boolean') return v ? loc.booleans.TRUE : loc.booleans.FALSE
      return errorIn(v.error, loc.errors)
    },
    formula,
  }
}

/**
 * A message in the language of a locale: a loaded locale's tag (en-US without one) or a locale object, its own messages
 * or else those of the loaded locale of its tag. `formula` shows a stored formula in the locale, for the messages that
 * hold one.
 */
export function say<K extends MessageKey>(where: string | Locale | undefined, m: { key: K } & MessageParams[K], formula: (text: string) => string = (f) => f): string {
  const loc = localeOf(where)
  const own = (loc.messages ?? LOCALES.get(loc.tag)?.messages)?.[m.key]
  return own ? own(m, wordsOf(loc, formula)) : english(m)
}
