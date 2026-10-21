/** The contract's types name what the calls return: the refusals of typed text include a reference to a workbook not among the links. */
import { describe, expect, it } from 'vitest'
import { createSession } from '../src/index.js'

const OTHER = '=[Book9.xlsx]Sheet1!A1'

describe.each(['closures', 'eval'] as const)('%s', (execution) => {
  it("the calls that type text refuse a workbook not among the links with 'external_ref', a code of their result types", () => {
    const s = createSession({ sheets: [{ name: 'S', cells: { A1: { value: 'x' } } }] }, { clock: { now: '2026-10-04T12:00:00Z', seed: 7 }, execution })
    const set = s.setCellText('S', 'B1', OTHER)
    const preview = s.previewCellText('S', 'B1', OTHER)
    const texts = s.setCellTexts([{ sheet: 'S', cell: 'C1', text: OTHER }])
    const textsPreview = s.previewCellTexts([{ sheet: 'S', cell: 'C1', text: OTHER }])
    const fill = s.fillCellText('S', 'D1:D2', OTHER)
    const fillPreview = s.previewFillCellText('S', 'D1:D2', OTHER)
    const replaced = s.replaceInCells([{ sheet: 'S', range: 'A1' }], 'x', OTHER)
    // each comparison compiles only where the result's code can be 'external_ref'
    expect([
      !set.ok && set.code === 'external_ref',
      !preview.ok && preview.code === 'external_ref',
      !texts.ok && texts.code === 'external_ref',
      !textsPreview.ok && textsPreview.code === 'external_ref',
      !fill.ok && fill.code === 'external_ref',
      !fillPreview.ok && fillPreview.code === 'external_ref',
      replaced.refused?.[0].code === 'external_ref',
    ]).toEqual(Array(7).fill(true))
    s.close()
  })
})
