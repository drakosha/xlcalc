/**
 * en-US: formula text is the invariant syntax itself, as Excel with the English interface shows formulas. Its rules of
 * computation are the engine's own code (`src/locale/rules.ts`), so the file has no regional data. Its messages are
 * every message the engine gives, the text a locale without its own falls back to.
 */
import type { Counts, EnglishMessages, Found, NameRule } from '../locale/messages.js'
import type { Locale } from '../locale/types.js'
import { ERROR_TEXT } from '../syntax/errors.js'

const found = (f: Found) => (f.t === 'op' || f.t === 'ident' || f.t === 'fn' ? JSON.stringify(f.v) : f.t === 'eof' ? 'end of formula' : f.t)

const counts = ({ min, max, parity, or }: Counts) => {
  const range = min === max ? `${min}` : `${min} to ${max}`
  return !parity ? range : or ? `${min}, or ${parity} ${min + 1} to ${max}` : `${range} (${parity})`
}

const NAME_RULES: Record<NameRule, (name: string) => string> = {
  chars: () => 'a name starts with a letter of any script, _ or \\, and has only letters, digits, _, ., \\ and ? (255 characters at most)',
  cell: () => 'a name must not look like a cell address or an R1C1 reference',
  true: () => 'TRUE and FALSE are reserved',
  boolean: (name) => `${name} is a boolean in the book's locale`,
}

export const EN_MESSAGES: EnglishMessages = {
  rowOutOfRange: (p) => `row ${p.row} out of range`,
  columnOutOfRange: (p) => `column ${p.column} out of range`,
  spillAfterRange: () => 'a spill reference (#) follows a single cell',
  spillAcrossSheets: () => 'a spill reference (#) follows a single cell of one sheet',
  unterminatedString: () => 'unterminated string',
  unknownErrorLiteral: () => 'unknown error literal',
  externalReference: () => 'external workbook reference',
  referenceAfterSheet: () => 'reference expected after sheet name',
  quotedSheetAlone: () => 'quoted sheet name without reference',
  unexpectedCharacter: (p) => `unexpected character ${JSON.stringify(p.char)}`,
  emptyFormula: () => 'empty formula',
  tooLong: (p) => `formula longer than ${p.max}`,
  tooDeep: (p) => `nesting deeper than ${p.max}`,
  unexpectedEnd: () => 'unexpected end of formula',
  unexpected: (p) => `unexpected ${found(p.found)}`,
  expected: (p) => `expected ${p.char}, found ${found(p.found)}`,
  commaOutside: () => 'comma outside of a function call',
  arraySeparator: (p) => `expected , ; or } in an array constant, found ${found(p.found)}`,
  arrayRows: () => 'the rows of an array constant must have the same number of values',
  arrayValue: (p) => `an array constant holds numbers, text, booleans and errors only, found ${found(p.found)}`,
  letTwice: (p) => `LET names ${p.name} twice`,
  letName: (p) => `LET takes a name as its argument ${p.position}`,
  letEmpty: () => 'LET takes a name and its value at least',
  letTooMany: (p) => `LET takes at most ${p.max} arguments`,
  specUnterminated: () => 'unterminated structured reference',
  specBracketInColumn: () => "unexpected [ in a column name (write it as '[)",
  specUnknown: (p) => `unknown table specifier ${p.word}`,
  specSpan: () => 'expected [ after : in a structured reference',
  specOpen: () => 'expected [ in a structured reference',
  specTwoColumns: () => 'two columns in a structured reference (write a span as [A]:[B])',
  specTwice: (p) => `${p.word} twice in a structured reference`,
  specClose: () => 'expected ] to close a structured reference',
  specTogether: (p) => `a structured reference cannot take ${p.words.join(' and ')} together`,
  externalLiterals: (p) => `arguments of ${p.fn} must be literals`,
  argumentCount: (p) => `${p.fn} takes ${counts(p.counts)} arguments, got ${p.got}`,
  argumentLeftOut: (p) => `${p.fn} takes a reference as its argument ${p.position}, which cannot be left out`,
  tableRefOutside: () => 'a structured reference without a table name outside a table',
  unknownTable: (p) => `unknown table ${p.table}`,
  unknownColumn: (p) => `the table ${p.table} has no column ${p.column}`,
  unknownFunction: (p) => `unknown function ${p.fn}`,
  unknownSheet: (p) => `unknown sheet ${p.sheet}`,
  unknownName: (p) => `unknown name ${p.name}`,
  unterminatedSheet: () => 'unterminated quoted sheet name',
  unexpectedText: (p) => `unexpected ${JSON.stringify(p.text)}`,
  unexpectedInArray: (p) => `unexpected ${JSON.stringify(p.text)} in an array constant`,
  notWritable: (p) => `${JSON.stringify(p.text)} cannot be written in ${p.tag ?? 'the invariant syntax'}`,
  nameRule: (p) => NAME_RULES[p.rule](p.name),
  sheetNameEmpty: () => 'a sheet name must not be empty',
  sheetNameLong: (p) => `a sheet name has at most 31 characters, ${JSON.stringify(p.name)} has ${p.length}`,
  sheetNameChar: (p) => `a sheet name must not contain ${JSON.stringify(p.char)} (Excel forbids : \\ / ? * [ ])`,
  sheetNameApostrophe: () => 'a sheet name must not begin or end with an apostrophe',
  sheetNameHistory: () => 'Excel reserves the sheet name "History" (in any case)',
  sheetNameTaken: (p) => `there is a sheet named ${JSON.stringify(p.name)} already (sheet names are case-insensitive)`,
  lastSheet: () => 'a book keeps at least one sheet',
  lastVisibleSheet: (p) => `a book keeps at least one visible sheet, and ${JSON.stringify(p.sheet)} is its last`,
  veryHidden: (p) => `the very hidden sheet ${JSON.stringify(p.sheet)} is neither deleted nor moved: show it first`,
  insertOffSheet: (p) => `inserting ${p.count} ${p.lines} would push ${p.sheet}!${p.cell} off the sheet`,
  insertMergeOffSheet: (p) => `inserting ${p.count} ${p.lines} would push a merged area of ${p.sheet} off the sheet`,
  arrayPart: (p) => `the operation changes part of the array formula ${p.sheet}!${p.cell} (${p.area})`,
  moveOffSheet: (p) => `the range does not fit on ${p.sheet} at ${p.cell}`,
  copyOffSheet: (p) => `the copy does not fit on ${p.sheet} at ${p.cell}`,
  moveCutsMerge: (p) => `the moved cells cut the merged area ${p.sheet}!${p.area}`,
  pasteOnMerge: (p) => `the pasted cells land on part of the merged area ${p.sheet}!${p.area}`,
  filteredPaste: (p) => `the copy leaves out rows a filter hid and cannot be pasted over the hidden rows of ${p.sheet}!${p.area}`,
  tableLands: (p) => `the table ${p.table} would land on another table`,
  tableInsertTwo: (p) => `inserting ${p.lines} there would shift cells of the tables ${p.tables.join(' and ')} at once`,
  tableDeleteTwo: (p) => `deleting these ${p.lines} would shift cells of the tables ${p.tables.join(' and ')} at once`,
  tableHeader: (p) => `the header row of the table ${p.table} cannot be deleted without the rest of the table`,
  tableName: (p) => `${JSON.stringify(p.name)}: ${NAME_RULES[p.rule](p.name)}`,
  tableNameTaken: (p) => `there is a table named ${p.table} already (table names are case-insensitive)`,
  tableNameDefined: (p) => `${JSON.stringify(p.name)} is a defined name (tables and names share one namespace)`,
  tableOverTable: (p) => `${p.range} meets the table ${p.table}`,
  tableOverMerge: (p) => `${p.range} meets the merged area ${p.area}`,
  explainCells: (p) => `${p.count} cells`,
  explainSpill: (p) => `spilled from ${p.ref}`,
  // the value as JSON, an error as a cell shows it (`#DIV/0!`)
  explainValue: ({ value: v }) => (v === undefined ? '' : v === null ? 'empty' : typeof v === 'object' ? ERROR_TEXT[v.error] : JSON.stringify(v)),
  explainFormula: (p) => p.formula,
  explainErrorFrom: (p) => `(error from ${p.ref})`,
  explainErrorHere: () => '(the error begins here)',
}

/** en-US's formula text, the invariant syntax; its messages are `EN_MESSAGES`, which every locale falls back to. */
export const EN_US: Locale = {
  tag: 'en-US',
  decimalSeparator: '.',
  listSeparator: ',',
  arrayColumnSeparator: ',',
  arrayRowSeparator: ';',
  booleans: { TRUE: 'TRUE', FALSE: 'FALSE' },
  errors: {},
  functions: {},
}
