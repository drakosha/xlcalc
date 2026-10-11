/**
 * The cases of the ui sets (`test/excel/ui/`, `test/excel/win/ui/`) where the engine knowingly differs from what Excel did,
 * by `<tag>:<set>:<id>`, with the reasons. A paste is answered by `setCellTexts`, which reads each text as typed; Excel
 * reads pasted text otherwise in places, alike on the Mac and on Windows. An entry into a selection (Control+Return) is
 * answered by `setCellTexts` with the same text in each cell, which Excel reads once, in the active cell.
 */

export const UI_REASONS = {
  pasteWidth: 'a paste widens no column in Excel (a long number shows in scientific form, a date as ###); setCellTexts widens it as typing does',
  pasteApostrophe: "a pasted apostrophe is part of the text in Excel ('007 holds '007); setCellTexts drops it as typing does",
  pasteFormula: 'Excel pastes a text it does not read as a formula as text (=1+, =SUM(A1,1.5) in ru-RU) and writes the other cells; setCellTexts refuses the whole call',
  pasteRectangle: 'a paste fills the rectangle of the clipboard: a missing field of a short row clears its cell, and a line break after the last row adds an empty row that clears the cell below; the host gave setCellTexts only the texts it had',
  pastePercent: "a paste into a percent cell has no percent editor: 5 is 5 (500%), 5% is 0.05, and 0.5 in ru-RU is text; setCellTexts reads a number as typing does (5 is 5%, 5% the text 5%%)",
  pasteValidation: 'a paste checks no rule of data validation in Excel, no window; setCellTexts gives the alert of each value its rule does not take',
  euro: "a pasted €5 takes the euro money format ([$€-2]\\ #,##0;[Red]\\-[$€-2]\\ #,##0) in Excel; parseEntry drops the sign with no format, as Formula2Local takes it (the keyboard's layout has no €)",
  fillShift: "Control+Return enters a formula into each cell of the selection with its relative references moved, as a copy from the active cell (=A1*2 in B2 is =A2*2); setCellTexts writes the same text in each cell",
  reference: "a formula that refers to a cell with a date format takes that format in Excel (=A1+A2), and widens the column for it; parseEntry has no book and gives a reference no format",
  fillPercent: "Control+Return reads the text once, by the active cell's format: with a percent cell active 5 is 5% in every cell, a cell in General too; with a cell in General active 5 is 5 in every cell, 500% in a percent cell; setCellTexts reads each cell by its own format",
  fillValidation: "Control+Return checks only the active cell's rule of data validation, once: a cell of the selection under a rule of its own takes the value with no window, and a stop alert takes it in no cell; setCellTexts checks each cell's rule (a stop alert taken back by the host)",
  macShortDate: "Excel for Mac's own short date in en-US is m/d/yy, narrower than Windows's m/d/yyyy that the engine follows: a typed date widens the column in the engine, not in Excel for Mac",
  arrayArea: "every cell of an old-kind array formula's area opens with the area's formula (F2), as its formula bar shows it in braces; editText gives it only in the area's first cell",
} as const

export type UiReason = keyof typeof UI_REASONS

/** Cases in both locales, and in each alone. */
const both = (all: string[], en: string[] = [], ru: string[] = []) => ({ 'en-US': [...all, ...en], 'ru-RU': [...all, ...ru] })

/** Each reason's cases by `<set>` and locale. */
const BY_REASON: Record<UiReason, Record<string, Record<string, string[]>>> = {
  pasteWidth: { 'paste-held': both(['wide-1', 'wide-2', 'dates-1', 'money-1', 'date-cell-1'], ['grid-1']) },
  pasteApostrophe: { 'paste-held': both(['apostrophe-1']) },
  pasteFormula: { 'paste-held': both(['bad-1'], ['formula-ru-1'], ['formula-en-1']) },
  pasteRectangle: { 'paste-held': both(['empty-1', 'trailing-1']) },
  pastePercent: { 'paste-held': both(['percent-cell-1']) },
  pasteValidation: { 'paste-held': both(['valid-1', 'valid-2']) },
  euro: { 'paste-held': both(['money-1']) },
  fillShift: { 'fill-held': both(['formula-1', 'formula-2', 'formula-3', 'formula-4']) },
  reference: { 'fill-held': both(['formula-3']) },
  fillPercent: { 'fill-held': both(['formats-1', 'formats-2']) },
  fillValidation: { 'fill-held': both(['valid-1', 'valid-2', 'valid-3']) },
  macShortDate: { 'fill-held': { 'en-US': ['date-1'] } },
  arrayArea: { 'arrays-held': both(['legacy-1']) },
}

/** `<tag>:<set>:<id>` → the reasons the case differs for. */
export const UI_KNOWN: ReadonlyMap<string, UiReason[]> = (() => {
  const out = new Map<string, UiReason[]>()
  for (const [reason, sets] of Object.entries(BY_REASON) as [UiReason, Record<string, Record<string, string[]>>][])
    for (const [set, tags] of Object.entries(sets)) for (const [tag, ids] of Object.entries(tags)) for (const id of ids) out.set(`${tag}:${set}:${id}`, [...(out.get(`${tag}:${set}:${id}`) ?? []), reason])
  return out
})()
