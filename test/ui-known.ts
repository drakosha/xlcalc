/**
 * The cases of the ui sets (`test/excel/ui/`, `test/excel/win/ui/`) where the engine knowingly differs from what Excel did,
 * by `<tag>:<set>:<id>`, with the reasons. A paste is answered by `pasteText`, an entry into a selection (Control+Return)
 * by `fillCellText`, a cell of an array by `editText`; what is left is what typing differs in, which these calls share.
 */

export const UI_REASONS = {
  reference: "a formula that refers to a cell with a date format takes that format in Excel (=A1+A2), and widens the column for it; parseEntry has no book and gives a reference no format",
  macShortDate: "Excel for Mac's own short date in en-US is m/d/yy, narrower than Windows's m/d/yyyy that the engine follows: a typed date widens the column in the engine, not in Excel for Mac",
} as const

export type UiReason = keyof typeof UI_REASONS

/** Cases in both locales, and in each alone. */
const both = (all: string[], en: string[] = [], ru: string[] = []) => ({ 'en-US': [...all, ...en], 'ru-RU': [...all, ...ru] })

/** Each reason's cases by `<set>` and locale. */
const BY_REASON: Record<UiReason, Record<string, Record<string, string[]>>> = {
  reference: { 'fill-held': both(['formula-3']) },
  macShortDate: { 'fill-held': { 'en-US': ['date-1'] } },
}

/** `<tag>:<set>:<id>` → the reasons the case differs for. */
export const UI_KNOWN: ReadonlyMap<string, UiReason[]> = (() => {
  const out = new Map<string, UiReason[]>()
  for (const [reason, sets] of Object.entries(BY_REASON) as [UiReason, Record<string, Record<string, string[]>>][])
    for (const [set, tags] of Object.entries(sets)) for (const [tag, ids] of Object.entries(tags)) for (const id of ids) out.set(`${tag}:${set}:${id}`, [...(out.get(`${tag}:${set}:${id}`) ?? []), reason])
  return out
})()
