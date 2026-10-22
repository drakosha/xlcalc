/**
 * en-US: formula text is the invariant syntax itself, as Excel with the English interface shows formulas. Its rules of
 * computation are the engine's own code (`src/locale/rules.ts`), so the file has no regional data. Its messages, with
 * those of an edit (`src/locale/edit-messages.ts`), are every message the engine gives, the text a locale without its own
 * falls back to.
 */
import type { Counts, EnglishMessages, Found, NameRule } from '../locale/messages.js'
import type { Locale } from '../locale/types.js'
import { OUTLINE_MESSAGES } from '../locale/outline-messages.js'

const found = (f: Found) => (f.t === 'op' || f.t === 'ident' || f.t === 'fn' ? JSON.stringify(f.v) : f.t === 'eof' ? 'end of formula' : f.t)

const counts = ({ min, max, parity, or }: Counts) => {
  const range = min === max ? `${min}` : `${min} to ${max}`
  return !parity ? range : or ? `${min}, or ${parity} ${min + 1} to ${max}` : `${range} (${parity})`
}

const NAME_RULES: Record<NameRule, (name: string) => string> = {
  chars: () => 'a name must start with a letter (of any script), _ or \\ and contain only letters, digits, _, ., \\ and ? (255 characters at most)',
  cell: () => 'a name must not look like a cell address or an R1C1 reference',
  true: () => 'TRUE and FALSE are reserved',
  boolean: (name) => `${name} is a boolean in the workbook's locale`,
}

export const EN_MESSAGES: EnglishMessages = {
  rowOutOfRange: (p) => `row ${p.row} out of range`,
  columnOutOfRange: (p) => `column ${p.column} out of range`,
  numberOutOfRange: (p) => `number ${p.number} out of range`,
  spillAfterRange: () => 'a spill reference (#) must follow a single cell',
  spillAcrossSheets: () => 'a spill reference (#) must follow a single cell on one sheet',
  unterminatedString: () => 'unterminated string',
  unknownErrorLiteral: () => 'unknown error literal',
  externalReference: () => 'external workbook reference',
  externalAcrossSheets: () => 'a reference across the sheets of another workbook is not supported',
  externalSpill: () => 'a spill reference into another workbook is not supported',
  referenceAfterSheet: () => 'reference expected after sheet name',
  quotedSheetAlone: () => 'quoted sheet name without a reference',
  unexpectedCharacter: (p) => `unexpected character ${JSON.stringify(p.char)}`,
  emptyFormula: () => 'empty formula',
  tooLong: (p) => `formula longer than ${p.max} characters`,
  tooDeep: (p) => `nesting deeper than ${p.max} levels`,
  unexpectedEnd: () => 'unexpected end of formula',
  unexpected: (p) => `unexpected ${found(p.found)}`,
  expected: (p) => `expected ${p.char}, found ${found(p.found)}`,
  commaOutside: () => 'comma outside a function call',
  arraySeparator: (p) => `expected , ; or } in an array constant, found ${found(p.found)}`,
  arrayRows: () => 'the rows of an array constant must have the same number of values',
  arrayValue: (p) => `an array constant can hold only numbers, text, booleans and errors, found ${found(p.found)}`,
  letTwice: (p) => `LET defines the name ${p.name} twice`,
  letName: (p) => `LET expects a name as argument ${p.position}`,
  letEmpty: () => 'LET needs at least a name and its value',
  letTooMany: (p) => `LET takes at most ${p.max} arguments`,
  specUnterminated: () => 'unterminated structured reference',
  specBracketInColumn: () => "unexpected [ in a column name (write it as '[)",
  specUnknown: (p) => `unknown table specifier ${p.word}`,
  specSpan: () => 'expected [ after : in a structured reference',
  specOpen: () => 'expected [ in a structured reference',
  specTwoColumns: () => 'two columns in a structured reference (write a span as [A]:[B])',
  specTwice: (p) => `${p.word} twice in a structured reference`,
  specClose: () => 'expected ] to close a structured reference',
  specTogether: (p) => `a structured reference cannot combine ${p.words.join(' and ')}`,
  externalLiterals: (p) => `arguments of ${p.fn} must be literals`,
  argumentCount: (p) => `${p.fn} takes ${counts(p.counts)} arguments, got ${p.got}`,
  argumentLeftOut: (p) => `${p.fn} expects a reference as argument ${p.position}, which cannot be omitted`,
  tableRefOutside: () => 'a structured reference without a table name outside a table',
  unknownTable: (p) => `unknown table ${p.table}`,
  unknownColumn: (p) => `the table ${p.table} has no column ${p.column}`,
  unknownFunction: (p) => `unknown function ${p.fn}`,
  unknownSheet: (p) => `unknown sheet ${p.sheet}`,
  unknownLink: (p) => `the workbook has no link to ${p.book}`,
  unknownLinkSheet: (p) => `the workbook ${p.book} has no sheet ${p.sheet}`,
  unknownName: (p) => `unknown name ${p.name}`,
  unterminatedSheet: () => 'unterminated quoted sheet name',
  unexpectedText: (p) => `unexpected ${JSON.stringify(p.text)}`,
  unexpectedInArray: (p) => `unexpected ${JSON.stringify(p.text)} in an array constant`,
  notWritable: (p) => `${JSON.stringify(p.text)} cannot be written in ${p.tag ?? 'the invariant syntax'}`,
  nameRule: (p) => NAME_RULES[p.rule](p.name),
  sheetNameEmpty: () => 'a sheet name must not be empty',
  sheetNameLong: (p) => `a sheet name can have at most 31 characters, and ${JSON.stringify(p.name)} has ${p.length}`,
  sheetNameChar: (p) => `a sheet name must not contain ${JSON.stringify(p.char)} (Excel forbids : \\ / ? * [ ])`,
  sheetNameApostrophe: () => 'a sheet name must not begin or end with an apostrophe',
  sheetNameHistory: () => 'Excel reserves the sheet name "History" (regardless of case)',
  veryHidden: (p) => `the very hidden sheet ${JSON.stringify(p.sheet)} cannot be deleted or moved: unhide it first`,
  ...OUTLINE_MESSAGES,
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
