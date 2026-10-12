/**
 * The English text of the messages of xlsx issues: the subpath's own, so that a bundle without it (`compute`, the root,
 * `xlcalc/text`) does not carry them. A locale writes them in its `messages` beside the engine's (ru-RU's), and a locale
 * without its own text of one gives this one.
 */
import { ownText, type StyleLoss, type StylePart, type XlsxMessageParams } from '../locale/messages.js'
import type { Locale } from '../locale/types.js'
import type { XlsxMessage } from './issues.js'

const STYLE_PARTS: Record<StylePart, string> = {
  accountingUnderline: 'accounting underline',
  underlineRemoved: 'an underline taken away',
  superscript: 'superscript',
  subscript: 'subscript',
  fontName: 'font name',
  fontSize: 'font sz',
  fontVertAlign: 'font vertAlign',
  fontOutline: 'font outline',
  fontShadow: 'font shadow',
  fontCondense: 'font condense',
  fontExtend: 'font extend',
  diagonalBorder: 'diagonal border',
  gradientFill: 'gradient fill',
  gradientBesidePattern: 'gradient fill beside its pattern',
  patternFill: 'pattern fill',
  fillRemoved: 'fill taken away',
  fillWithoutColor: 'fill without its colour',
  textRotation: 'text rotation',
  shrinkToFit: 'shrink to fit',
  readingOrder: 'reading order',
  justifyLastLine: 'justify last line',
  alignment: 'alignment',
  protection: 'protection',
}

const stylePart = (x: StyleLoss) => (typeof x === 'string' ? STYLE_PARTS[x] : `border ${x.border}`)

/** The messages of xlsx issues in English, what a locale without its own text of one falls back to. */
export const XLSX_EN_MESSAGES: { readonly [K in keyof XlsxMessageParams]: (p: XlsxMessageParams[K]) => string } = {
  // reading: the workbook
  localeUnknown: (p) => `the locale ${JSON.stringify(p.tag)} is not one of the package's, left out`,
  builtInName: () => 'a built-in name (a print area or titles, the range of a filter) is not in the model',
  nameOfSheetLeftOut: () => 'a name of a sheet left out is left out',
  hiddenName: () => 'a hidden name is read as a visible one',
  manualCalculation: () => 'the manual calculation mode is not kept',
  autoNoTable: () => 'the calculation mode "automatic except for data tables" is not kept',
  precisionAsDisplayed: () => 'precision as displayed is not in the engine: values keep their full precision',
  iterationCount: (p) => `the maximum of iterations ${JSON.stringify(p.count)} is not one of 1 to 32,767: the default, 100, is read`,
  iterationChange: (p) => `the maximum change of iterations ${JSON.stringify(p.change)} is not a number from 0: the default, 0.001, is read`,
  workbookProtection: () => 'workbook protection is not in the model',
  externalLinks: () => 'a link to the data of another program (DDE, OLE), not to a workbook, or a link whose part is missing, is not in the model: its number is kept by a book without sheets, and formulas that refer to it compute to an error',
  pivotTables: () => 'pivot tables are not in the model: their cells are read as stored',
  macros: () => 'macros (a VBA project) are not in the model',
  connections: () => 'data connections and queries are not in the model',
  defaultFont: (p) => `the default font ${p.name ?? '(no name)'} ${p.size ?? '(no size)'} is read as Calibri 11, the model's default`,
  sheetKind: () => 'a chart, dialog or macro sheet is left out',
  sheetPartMissing: () => 'a sheet whose part is missing is left out',
  allSheetsHidden: () => 'every worksheet of the file is hidden: the first is read as a visible one',
  tableNameTwice: (p) => `the table ${p.table} is left out: its name is another table's or a defined name`,
  // reading: a sheet
  sharedWithoutMaster: () => 'a shared formula without its master cell is left out',
  sharedCopy: () => 'a copy of a shared formula that a paste could not make is left out',
  outlineStyles: () => 'automatic styles of an outline are not kept',
  mergeOverlap: () => 'a merged area overlapping another is left out',
  autoFilterCriteria: () => "an autofilter's range, criteria and sort are not in the model: the sheet is read as filtered, the rows it hid as hidden",
  autoFilterPlain: () => 'an autofilter without criteria (its buttons and sort) is not in the model',
  advancedFilter: () => 'an advanced filter is not in the model: the sheet is read as filtered, the rows it hid as hidden',
  tableFilter: () => "a table's filter criteria are not in the model: the sheet is read as filtered, the rows they hid as hidden",
  sheetProtection: () => 'sheet protection is not in the model',
  drawings: () => 'pictures, charts and shapes are not in the model',
  sparklines: () => 'sparklines are not in the model',
  printSettings: () => 'print settings (headers and footers, page breaks, page setup) are not in the model',
  defaultWidth: (p) => `the sheet's default column width (${p.width} in xlsx units) is not in the model, which takes Excel's standard of 8 digits`,
  errorTyped: () => 'an error typed into a cell is read as a formula of that error',
  // reading: formulas
  userFunction: () => 'a user-defined function (a macro or an add-in) is not in the engine',
  tableRefWithoutTable: () => 'a structured reference without a table name is not in a file: the formula is read as written and computes to an error',
  formulaExternal: () => 'a reference to another workbook that is not one of the links the file holds, or that the engine does not parse: the formula is read as written and computes to an error',
  letPrefix: () => 'a LET name that reads as a reference or a boolean once its _xlpm. prefix goes: the formula changes its meaning',
  formulaUnparsed: () => 'a formula the engine does not parse is read as written (its cell computes to an error)',
  // reading: cells
  cellOutside: () => 'a cell outside the sheet (past its last row or column) is left out',
  valueMetadata: () => "a cell's value metadata (a data type, a picture in the cell) is not in the model",
  richText: () => 'formatted runs of text are read as plain text',
  errorUnknown: () => 'a stored error the model has no code for, read as #ERROR!',
  errorUnspelled: () => 'an error stored as a value that a formula cannot spell (#CALC!, #GETTING_DATA and the like), left out',
  dataTable: () => 'a what-if data table (TABLE) is read as its stored values',
  styleLost: (p) => `a style's ${stylePart(p.part)} is not in the model`,
  eastAsianFormat: () => 'a built-in East Asian date format the model does not have (50 to 58), read as the code Excel for Mac en-US shows',
  unknownFormat: () => 'a built-in number format this reader does not know, left out',
  // reading: tables
  tablePartBad: () => 'a table part without a name or a range is left out',
  tableLeftOut: (p) => `the table ${p.table} is left out: ${p.reason}`,
  headerGiven: (p) => `a header cell that does not hold its column's name as text is read with the name ${JSON.stringify(p.name)}`,
  // reading: conditional formats
  condRangeOutside: () => 'a range of a conditional format outside the sheet is left out',
  condStyleLost: (p) => `a conditional format's ${stylePart(p.part)} is not in the model`,
  condOperator: (p) => `a cell value rule with the operator ${p.operator} is not in the model`,
  condNoFormulas: () => 'a cell value rule without its formulas is left out',
  condNoFormula: () => 'a formula rule without its formula is left out',
  condPeriod: (p) => `a time period ${p.period} is not in the model`,
  condAverage: () => 'an average rule with both equal and standard deviations is read with equal alone, as Excel keeps it',
  condColorScale: () => 'a colour scale without its points or colours is left out',
  condDataBar: () => 'a data bar without its thresholds is left out',
  condIconSet: (p) => `the icon set ${p.set} is not in the model`,
  condIcons: () => 'an icon set without a threshold for each icon is left out',
  condType: (p) => `a conditional format of type ${p.type || '(none)'} is not in the model`,
  condPivot: () => "a pivot table's conditional formatting is not in the model",
  // reading: data validation
  ruleListCommas: () => "a list's items holding a comma are read split at it",
  ruleType: (p) => `a rule of an unknown type ${JSON.stringify(p.type)} is left out`,
  ruleNoFormula: () => 'a rule without the formula its type needs is left out',
  ruleNoSecond: () => 'a rule between two bounds without its second formula is left out',
  ruleIme: () => "a rule's IME mode is not kept",
  ruleNoRange: () => 'a rule without a range of the sheet is left out',
  ruleShared: () => 'cells of a rule that an earlier rule has keep the earlier one',
  // reading: notes, hyperlinks and panes
  splitPanes: () => 'split panes that are not frozen are not in the model',
  frozenScrolled: () => 'panes frozen over a sheet scrolled down or right are read from the first row and column',
  noteCell: () => 'a note or a thread of comments whose cell is no one cell of the sheet is left out',
  noteThread: () => "a thread of comments is read as a note: its first comment's author, every comment's text with the replies' authors; dates, mentions and the thread's state are not kept",
  noteFormatted: () => "a note's formatted text is read as its plain text",
  noteShown: () => 'a note shown all the time is read as one shown when the pointer rests on its cell',
  linkAddress: () => "a hyperlink's address that is no external relationship of the sheet is left out",
  linkCells: () => 'a hyperlink over more than 4,096 cells is read on the cells of its range that hold something',
  // writing
  nameNotTaken: () => 'a name Excel does not take is left out of the file',
  nameNoSheet: () => 'a name of a sheet the book does not have is left out of the file',
  nameTwice: () => 'a name given twice is left out of the file',
  nameBadRef: () => 'a name whose reference does not parse or names a table, a column or a link to another workbook the book does not have is left out of the file',
  noResult: () => 'no result was given: the file stores no values of formulas, and Excel computes them when it opens the file',
  formulaNoValue: () => 'a formula without a value in the result is written without one: Excel computes it when it opens the file',
  formulaEmpty: () => 'a formula without a value in the result is written as an empty cell',
  formulaAsValue: () => 'a formula the file cannot hold (it does not parse, or names a sheet, a table, a column or a link to another workbook the book does not have) is written as its value',
  tableFormulaOut: (p) => `a formula of table ${p.table} that the file cannot hold is left out`,
  cycleZero: () => 'a circular reference (#CYCLE!) is written as 0, as Excel stores it',
  cellMeta: () => 'cell meta has no place in a file',
  textCut: () => "text over Excel's 32,767 characters is cut",
  noteCut: () => "a note over Excel's 32,767 characters is cut",
  tooltipCut: () => "a hyperlink's tooltip over Excel's 255 characters is cut",
  ruleTextCut: (p) => `a rule's ${p.field} over Excel's ${p.max} characters is cut`,
  ruleFormulaOut: () => 'a rule whose formula the file cannot hold (it does not parse, or names a sheet, a table or a column the book does not have) is left out',
  condFormulaOut: () => 'a conditional format with a formula the file cannot hold (it does not parse, or names a sheet, a table or a column the book does not have) is left out of the file',
  filterModeOut: () => "the filter mode of a sheet whose tables have no header row is left out: an autofilter over the sheet's cells would meet them",
}

/** An issue's message in a locale (en-US without one): the locale's own text, else the English one. */
export function issueText(locale: string | Locale | undefined, m: XlsxMessage): string {
  return ownText(locale, m) ?? (XLSX_EN_MESSAGES[m.key] as (p: XlsxMessage) => string)(m)
}
