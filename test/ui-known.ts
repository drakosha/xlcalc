/**
 * The cases of the ui sets (`test/excel/ui/`, `test/excel/win/ui/`) where the engine knowingly differs from what Excel did,
 * by `<tag>:<set>:<id>`, a set of Excel for Windows alone as `win/<set>`, with the reasons; a cell of the ribbon's sets
 * by `<tag>:<set>:<id>:<cell>`. A paste is answered
 * by `pasteText`, an entry into a selection (Control+Return) by `fillCellText`, a cell of an array by `editText`; what is
 * left is what typing differs in, which these calls share. The ribbon's box by `numberFormatCategory`, its list by
 * `numberFormatChoices`.
 */

export const UI_REASONS = {
  macShortDate: "Excel for Mac's own short date in en-US is m/d/yy, narrower than Windows's m/d/yyyy that the engine follows: a typed date widens the column in the engine, not in Excel for Mac",
  shortDateCode: "Excel for Mac's built-in short date (14) is m/d/yy in en-US, so it takes the code m/d/yyyy as one of the book's own (Custom); the package shows 14 by Excel for Windows's m/d/yyyy, and numberFormatCategory takes that code for the short date (Date), as builtinFormatCode(14) gives it",
  macSyntax: "Excel for Mac in a locale other than en-US compares the code, kept in en-US syntax (0.00), with its formats written in the locale's (0,00), and shows Custom ((все форматы)) for nearly every code; Excel for Windows writes the code in the locale's syntax first (NumberFormatLocal) and finds 0.00 among Числовой's, as numberFormatCategory does (test/format-category-win.test.ts). A rule, not a list: every cell of the Mac's ribbon sets in such a locale where the Mac shows Custom and the function another category (an item applied: the box after it)",
  ruCommaWidth: "Comma Style in ru-RU (_-* # ##0,00_-, a hyphen's width of padding at each end): Excel for Mac keeps the default column for 1 234,50, where the engine fits it to 54 points and widens it to 9; the format is the same, and the accounting format with ₽ widens the column to 10 in both",
  winGeometry: "Excel for Windows widens a column by its own geometry, 7 pixels a character at 96 dpi (12.5703125 for 11.86 characters), and keeps its default of 64 pixels where a number fits it (12345679), where the engine widens by Excel for Mac's, the geometry the library keeps (docs/formats.md, Fitting columns): the format is the same",
  winDialogUnread: "Excel for Windows refused AutoSum's SUM() as the Mac does, the cell left blank, but the step's press.ps1 did not find the window it answered, so no window is recorded",
  oneCellFill: "Control+Return in a selection of one cell (Excel for Mac, the harness's fill) left a formula's value in the cell as a constant, twice in two runs (=A5-A3 is 98, =A1 0.1 without its format): it types no formula there, and the case tells nothing of the format a formula takes",
  dateItems: "the list's Short Date, Long Date and Time items are not the same in every Excel session: in this run Excel offered Date (yyyy-mm-dd;@), Date (d/m;@) and Time (h:mm;@) whatever the cell held; the same case run in another session that day offered Short Date (built-in 14), Long Date ([$-F800]dddd, mmmm dd, yyyy) and Time ([$-F400]h:mm:ss AM/PM), the items numberFormatChoices gives",
} as const

export type UiReason = keyof typeof UI_REASONS

/** Cases in both locales, and in each alone. */
const both = (all: string[], en: string[] = [], ru: string[] = []) => ({ 'en-US': [...all, ...en], 'ru-RU': [...all, ...ru] })

/** Each reason's cases by `<set>` and locale. */
const BY_REASON: Record<UiReason, Record<string, Record<string, string[]>>> = {
  macShortDate: { 'fill-held': { 'en-US': ['date-1', 'formula-3'] } },
  oneCellFill: { 'fill-refs-held': { 'en-US': ['fill-5', 'fill-7'] } },
  ruCommaWidth: { 'press-held': { 'ru-RU': ['money-1', 'money-2'] } },
  // a set of Excel for Windows by `win/<set>`
  winGeometry: {
    'win/press-third-held': { 'en-US': ['whole-3', 'whole-4', 'whole-5', 'whole-6'] },
    'win/press-held': { 'en-US': [...['5', '7', '8', '11'].map((n) => `general-${n}`), 'code-19', 'code-41', 'code-42', ...['1', '2', '4', '5', '7'].map((n) => `money-${n}`)] },
    'win/press-more-held': { 'en-US': ['1', '3', '5', '7', '8', '9', '10', '13', '15', '17', '19'].map((n) => `general-${n}`) },
  },
  winDialogUnread: { 'win/press-held': { 'en-US': ['sum-none-1', 'sum-dates-1'] }, 'win/press-more-held': { 'en-US': ['sum-below-1'] } },
  shortDateCode: { ribbon: { 'en-US': ['code-1:B15', ...['B3', 'B4', 'B5', 'B6', 'B7', 'B12'].map((x) => `code-2:${x}`)] } },
  // by the rule in test/ui-excel.test.ts, as the cells are those where the function's category is not Custom
  macSyntax: {},
  dateItems: { ribbon: { 'en-US': [...['B5', 'B6', 'B7'].map((x) => `choose-1:${x}`), ...['B7', 'B8', 'B9'].map((x) => `choose-2:${x}`)] } },
}

/** `<tag>:<set>:<id>` → the reasons the case differs for. */
export const UI_KNOWN: ReadonlyMap<string, UiReason[]> = (() => {
  const out = new Map<string, UiReason[]>()
  for (const [reason, sets] of Object.entries(BY_REASON) as [UiReason, Record<string, Record<string, string[]>>][])
    for (const [set, tags] of Object.entries(sets)) for (const [tag, ids] of Object.entries(tags)) for (const id of ids) out.set(`${tag}:${set}:${id}`, [...(out.get(`${tag}:${set}:${id}`) ?? []), reason])
  return out
})()
