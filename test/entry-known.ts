/**
 * The cases of the entry sets (`test/excel/entry/`, `test/excel/win/entry/`) where `parseEntry` knowingly differs from
 * Excel's answer, by `<tag>:<set>:<id>`, with the reason. The sets typed through `Formula2Local` (`typed`, `typed-held`,
 * `typed-held-2`) answer as Excel's object model takes text; where it takes text otherwise than typing, the sets typed
 * on the keyboard (`typed-keys`, `typed-keys-held`, `typed-keys-held-2`) give the answer `parseEntry` follows.
 */

export const ENTRY_REASONS = {
  sign: "set through Formula2Local, Excel takes text begun by + or - that is no number as text; typed on the keyboard it makes a formula of it (-C1, +1+2 is 1+2), or brings its window for one that does not parse (-1,5 in en-US), as parseEntry does (typed-keys, typed-keys-held)",
  percent: "set through Formula2Local, a number goes into a percent cell as it is (5 is 500%); typed on the keyboard Excel's editor adds % after a number begun by a digit (5 is 5%, 5% becomes the text 5%%), as parseEntry does (typed-keys, typed-keys-held)",
  empty: 'set through Formula2Local, empty text makes a cell formatted as text hold empty text; typed, an empty entry clears the value, as parseEntry does',
  english: "Excel ru-RU takes an English function name or TRUE in a typed formula as an unknown name (#NAME?, no format); parseEntry reads them as toInvariant does, as the function and the boolean",
  reference: "a formula that refers to a cell with a date format takes that format in Excel (=C2, SUM(C2,1), C1+C2); parseEntry has no book and gives a reference no format",
  gettingData: '#GETTING_DATA typed (#ОЖИДАНИЕ_ДАННЫХ in ru-RU) is an error value in Excel; the engine has no such error and keeps the text',
  plusPercent: 'typed on the keyboard, +5%+1 takes the percent format of its first number; parseEntry gives a formula begun by + the format of what it computes, none here',
  plusInside: 'typed on the keyboard, Excel drops a + before a number further into a formula begun by + (+-+5 is +-5); parseEntry drops only the first',
  nameCase: 'Excel writes an unknown name in the case it was first typed in the book (=Foo*2 after =foo is foo*2, Mac and Windows alike); parseEntry has no book and keeps the case typed',
  letAfter: "a cell's word used as LET's name and again after the LET in one formula: Excel reads the later one as a defined name (_xlnm.IW1, #NAME?); parseEntry reads it as the cell",
  macLongS: "Excel for Mac reads the long s as s: =falſe and the value falſe typed are FALSE, =ſum(1,2) is SUM and ſ is LET's name S; parseEntry follows Excel for Windows, where ſ is a letter of its own (falſe a name or text, ſum an unknown function)",
  typedI: 'Excel reads a typed function\'s name with the dotted İ as I (=İF(TRUE,1,2) is IF), and Excel for Windows with the dotless ı too (=ıf(TRUE,1,2)); parseEntry reads a function\'s name by its own letters, as Excel reads it in a file (İF(TRUE,1,2) is #NAME? there)',
  macCode: "Excel for Mac's own built-in short date (en-US m/d/yy) and short time (ru-RU hh:mm) are not Windows's, so it keeps the codes m/d/yyyy and h:mm as formats of the book's own, which a typed form never replaces; parseEntry follows Excel for Windows, which takes them as the built-in formats 14 and 20",
} as const

export type EntryReason = keyof typeof ENTRY_REASONS

/** `<tag>:<set>` and its case ids → the same reason. */
function cases(reason: EntryReason, bySet: Record<string, string[]>): [string, EntryReason][] {
  return Object.entries(bySet).flatMap(([set, ids]) => ids.map((id): [string, EntryReason] => [`${set}:${id}`, reason]))
}

const SIGN_TYPED = ['sign-1', 'sign-2', 'sign-3', 'sign-4', 'sign-5', 'sign-7', 'sign-8', 'sign-14', 'sign-15', 'sign-16', 'sign-18', 'sign-19']
const PERCENT_TYPED = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '15', '16', '17', '18'].flatMap((n) => [`in-pct-${n}`, `in-pct2-${n}`])
const PERCENT_CODE = ['1', '2', '3', '4', '5', '6', '7', '8', '9'].map((n) => `in-code-pct-${n}`)
const PERCENT_CODE2 = ['1', '2', '3', '4', '7', '8'].map((n) => `in-code-pct2-${n}`)

/** Against Excel for Mac and Excel for Windows alike. */
export const ENTRY_KNOWN: ReadonlyMap<string, EntryReason> = new Map([
  ...cases('sign', {
    'en-US:typed': ['num-12', 'grp-13', 'cur-4', 'sci-5', ...SIGN_TYPED, 'sign-20'],
    'ru-RU:typed': ['num-13', 'grp-14', 'cur-11', 'sci-6', ...SIGN_TYPED, 'sign-21'],
    'en-US:typed-held': ['more-6'],
    'ru-RU:typed-held': ['more-5'],
    'en-US:typed-held-2': ['pct2-2'],
    'ru-RU:typed-held-2': ['pct2-2'],
    'en-US:typed-pct-held': ['pru-12'],
    'ru-RU:typed-pct-held': ['pen-12'],
  }),
  ...cases('percent', {
    'en-US:typed': PERCENT_TYPED,
    'ru-RU:typed': PERCENT_TYPED,
    'en-US:typed-held': PERCENT_CODE,
    'ru-RU:typed-held': PERCENT_CODE,
    'en-US:typed-held-2': PERCENT_CODE2,
    'ru-RU:typed-held-2': PERCENT_CODE2,
  }),
  ...cases('empty', { 'en-US:typed': ['in-text-22'], 'ru-RU:typed': ['in-text-22'] }),
  ...cases('english', {
    'ru-RU:typed': ['fx-13', 'fx-17', 'fx-21'],
    'ru-RU:typed-held': ['ffx-1', 'ffx-2', 'ffx-4', 'ffx-22', 'ffx-23', 'ffx-24', 'in-code-pct-11', ...['b41', 'b45', 'b46', 'b48', 'b13', 'b18', 'b7', 'b5'].map((f) => `in-${f}-11`)],
    'ru-RU:typed-held-2': ['ffx2-1', 'ffx2-8', 'ffx2-9', ...['sci', 'frac', 'mmss', 'pct2', 'elapsed'].map((f) => `in-code-${f}-5`)],
  }),
  ...cases('reference', {
    'en-US:typed': ['fx-23', 'fx-24'],
    'ru-RU:typed': ['fx-23', 'fx-24'],
    'en-US:typed-held': ['ffx-14', 'ffx-15', 'ffx-16'],
    'ru-RU:typed-held': ['ffx-16', 'ffx-32'],
  }),
  ...cases('gettingData', { 'en-US:typed': ['err-19'], 'ru-RU:typed': ['err-20'] }),
  ...cases('plusPercent', { 'en-US:typed-keys-held-2': ['key2-sign-13'], 'ru-RU:typed-keys-held-2': ['key2-sign-13'] }),
  ...cases('plusInside', { 'en-US:typed-keys-held-2': ['key2-sign-15'], 'ru-RU:typed-keys-held-2': ['key2-sign-15'] }),
  ...cases('nameCase', { 'en-US:names-held': ['unknown-3'], 'en-US:name-case-held': ['let-4', 'let-7'] }),
  ...cases('typedI', { 'en-US:name-case-held': ['fn-6'] }),
  ...cases('letAfter', { 'en-US:let-grid-held': ['scope-2'] }),
])

/** Against Excel for Mac alone, where its regional data part from Windows's. */
export const ENTRY_MAC_KNOWN: ReadonlyMap<string, EntryReason> = new Map([
  ...cases('macCode', {
    'en-US:typed-held': ['in-code-mdy-4', 'in-code-mdy-5', 'in-code-mdy-6', 'in-code-mdy-8', 'in-code-mdy-10'],
    'ru-RU:typed': ['in-time-6', 'in-time-8', 'in-time-15', 'in-time-16', 'in-time-18'],
    'ru-RU:edit-held': ['time-cell-3'],
  }),
  ...cases('macLongS', {
    'en-US:name-case-held': ['bool-1', 'bool-2', 'bool-3', 'bool-4', 'bool-5', 'bool-6', 'fn-1', 'let-1', 'let-2'],
    'en-US:name-case-values-held': ['value-1', 'value-2'],
  }),
])

/** Against Excel for Windows alone. */
export const ENTRY_WIN_KNOWN: ReadonlyMap<string, EntryReason> = new Map([
  ...cases('nameCase', { 'en-US:name-case-held': ['bool-2'] }),
  ...cases('typedI', { 'en-US:name-case-held': ['fn-5'] }),
])
