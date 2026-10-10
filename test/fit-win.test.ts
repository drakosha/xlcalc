/**
 * AutoFit and typing against Excel for Windows's answers to the fit sets (`test/excel/win/fit/`). Excel for Windows lays
 * text out in pixels (a digit 7 at 96 dpi) where the Mac and the engine count points (a digit 6), so the widths it
 * fits differ from the Mac's in most cases (`WIN_SAME_WIDTH`, a known difference, `docs/formats.md`); what it does is
 * the same: a column with something to show is fitted and marked `bestFit`, a hidden one shown, a column with nothing to
 * show keeps its width, and text set through the object model widens no column.
 */
import { describe, expect, it } from 'vitest'
import type { FitAnswer, FitCase } from '../bench/coverage/fit/cases.js'
import { engineAnswer, excelWidth, fitSets } from './fit-sets.js'

const SETS = fitSets('win/')

/** What became of a case's column in Excel for Windows: fitted (`bestFit`), kept (a width without it), or the standard width. */
const fate = (a: FitAnswer) => (a.saved?.bestFit ? 'fitted' : a.saved ? 'kept' : 'standard')

/** The cases whose width Excel for Windows saved as the Mac did, by `<tag>:<set>`. */
const WIN_SAME_WIDTH: Record<string, number> = {
  'en-US:autofit': 16,
  'ru-RU:autofit': 16,
  'en-US:typed': 52,
  'ru-RU:typed': 52,
  'en-US:glyphs': 34,
}

describe.each(SETS)('fit $tag $name (Excel for Windows)', (set) => {
  it.each(set.cases.map((c) => [c.id, c] as const))('%s: fitted, kept or shown as the engine does', (_, c: FitCase) => {
    const a = set.answers[c.id]
    const e = engineAnswer(set, c)
    // a column fitted to Excel for Windows's standard width (64 pixels) is saved without one, as one it kept
    const wanted = e.fitted ? ['fitted', ...(c.column?.width === undefined || e.width === null ? ['standard'] : [])] : [c.column?.width === undefined ? 'standard' : 'kept']
    expect(wanted).toContain(fate(a))
    expect(a.saved?.hidden ?? false).toBe(e.fitted ? false : (c.column?.hidden ?? false))
  })
  it('widths in pixels: the cases Excel for Windows fits as the Mac does', () => {
    const mac = fitSets('').find((s) => s.tag === set.tag && s.name === set.name)!
    const same = set.cases.filter((c) => excelWidth(set.answers[c.id]) === excelWidth(mac.answers[c.id])).length
    expect(same).toBe(WIN_SAME_WIDTH[`${set.tag}:${set.name}`])
  })
})
