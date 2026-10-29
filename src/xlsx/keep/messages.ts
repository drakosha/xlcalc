/**
 * The English text of `writeKept`'s refusals (`KeptRefusal`): the `./xlsx/keep` subpath's own, as the xlsx issues' are
 * `./xlsx`'s. A locale writes them in its `messages` (ru-RU's), and a locale without its own text of one gives this one.
 */
import { ownText, type KeptMessageParams, type KeptPart, type KeptSheetPart, type Message } from '../../locale/messages.js'
import type { Locale } from '../../locale/types.js'

/** A message of a refusal of `writeKept`. */
export type KeptMessage = Message<keyof KeptMessageParams>

const OPEN_KEPT = 'open the file with openKept and keep the hidden names it adds'

/** The parts of a sheet named by the sheet alone. */
const OF_SHEET: Record<KeptSheetPart, string> = {
  autoFilter: 'the AutoFilter',
  sort: 'the sort',
  pageBreaks: 'the page breaks',
  protectedRanges: 'the protected ranges',
  sparklines: 'the sparklines',
  threads: 'the comment threads',
  formControls: 'the form controls',
}

/** A part of the file in English. */
function partText(w: KeptPart): string {
  switch (w.part) {
    case 'name':
      return `the name ${w.name}`
    case 'chart':
      return `the chart ${w.path}`
    case 'pivotChart':
      return `the pivot chart ${w.path}`
    case 'pivotSource':
      return `the source of the pivot cache ${w.path}`
    case 'drawing':
      return `a drawing on '${w.sheet}' (${w.path})`
    case 'pivotTable':
      return `the pivot table ${w.path} on '${w.sheet}'`
    case 'element':
      return `the ${w.element} of '${w.sheet}'`
    case 'extension':
      return `an extension of '${w.sheet}' (${w.uri})`
    case 'sheetKind':
      return `the ${w.kind} '${w.name}'`
    case 'unreadRule':
      return `an unread conditional format of '${w.sheet}' (${w.range})`
    default:
      return `${OF_SHEET[w.part]} of '${w.sheet}'`
  }
}

const rules = (sheet: string, why: string) => `the conditional formats of '${sheet}' changed, and ${why}`

/** The refusals of `writeKept` in English, what a locale without its own text of one falls back to. */
export const KEPT_EN_MESSAGES: { readonly [K in keyof KeptMessageParams]: (p: KeptMessageParams[K]) => string } = {
  keptStrict: () => 'the file is a Strict Open XML workbook, which writeKept does not write over',
  keptSheetName: (p) => `the file names a sheet '${p.name}', which Excel does not take (it was read as '${p.sheet}')`,
  keptFont: () => 'the default font changed, which writeKept does not write',
  keptValidation: (p) => `the data validation of '${p.sheet}' changed, and the file has rules of it the workbook could not read, which writing them again would lose`,
  keptThreads: (p) => `the notes of '${p.sheet}' changed, and it has comment threads, which writeKept does not write`,
  keptTableParts: (p) => `the table ${p.table} was deleted, and it has parts the workbook does not hold (${p.path})`,
  keptLinkKind: (p) => `the link [${p.link}] to another workbook changed, and its part is missing or is not a workbook's (DDE, OLE)`,
  keptLinks: () => "a link to another workbook was removed, moved or pointed at another file, which the file's formulas would read by another number",
  keptNoCarrier: (p) => `the workbook has no name ${p.name}, which tells what became of the sheet '${p.sheet}': ${OPEN_KEPT}`,
  keptLayout: (p) => `the workbook's name ${p.name}, which tells how openKept laid out the hidden names, is missing or was given by another version of openKept or for another file: open the original file again with openKept`,
  keptNoCellCarrier: (p) => `the workbook has no name ${p.name}, which tells where the cell ${p.cell} of '${p.sheet}' went: ${OPEN_KEPT}`,
  keptNoPartCarrier: (p) => `the workbook has no name ${p.name}, which holds a reference of ${partText(p.what)}: ${OPEN_KEPT}`,
  keptUnread: (p) => `the sheet '${p.sheet}' had rows or columns inserted or deleted, or was renamed or deleted, and ${partText(p.what)} holds references writeKept does not read`,
  keptNamedSheet: (p) => `the sheet '${p.sheet}' was deleted, and ${partText(p.what)} names it`,
  keptSheetParts: (p) => `the sheet '${p.sheet}' was deleted, and it has a part the workbook does not hold (${p.kind} ${p.path})`,
  keptSourceGone: (p) => `${partText(p.what)} is the table or name ${p.name}, which the workbook no longer has`,
  keptNoRef: (p) => `${partText(p.what)} is no reference after the change`,
  keptMovedSheet: (p) => `${partText(p.what)} moved to another sheet, which writeKept does not write`,
  keptFirstCell: (p) => `${partText(p.what)} lost its first cell, and its formulas have relative references, which writeKept does not make for the new first cell`,
  keptTakenApart: (p) => `${partText(p.what)} has a formula whose references the change took apart`,
  keptRulesPivot: (p) => rules(p.sheet, "the file has a pivot table's rules, which the pivot table names by priority"),
  keptRulesPart: (p) => rules(p.sheet, `the workbook read a rule over ${p.range} without what it holds (${p.lost})`),
  keptRulesPrefixes: (p) => rules(p.sheet, `a rule of Excel 2010 over ${p.range} is in prefixes of its own`),
  keptRulesBar: (p) => rules(p.sheet, `the workbook read the Excel 2010 part of a data bar over ${p.range} without its rule`),
  keptNamespace: (p) => `the worksheet's default namespace is ${p.namespace}, not SpreadsheetML's`,
  keptRelPrefix: (p) => `the worksheet's prefix r is ${p.namespace}, not the namespace of relationships`,
  keptAlternate: (p) => `${partText(p.what)} lies in cells the change deleted, in alternate content, which writeKept does not delete`,
  keptDrawingMoved: (p) => `${partText(p.what)} lies in a range that moved, which writeKept does not move a drawing with`,
  keptDrawingDeleted: (p) => `${partText(p.what)} lies in cells deleted with all those below or beside it`,
  keptCut: (p) => `${partText(p.what)} would be cut by the change, which Excel refuses`,
  keptCacheSheet: (p) => `${partText(p.what)} is on a sheet the change deleted, which Excel makes a link to the file by its path`,
  keptCellsDeleted: (p) => `${partText(p.what)} lies in cells the change deleted`,
  keptSparkline: (p) => `a sparkline's cell on '${p.sheet}' was deleted`,
  keptWouldMove: (p) => `${partText(p.what)} (${p.path}) would move, which writeKept does not write`,
}

/** A refusal's message in a locale (en-US without one): the locale's own text, else the English one. */
export function keptText(locale: string | Locale | undefined, m: KeptMessage): string {
  return ownText(locale, m) ?? (KEPT_EN_MESSAGES[m.key] as (p: KeptMessage) => string)(m)
}
